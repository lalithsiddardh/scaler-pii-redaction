"""detect.py - PII detection over the text surface of a .docx package.

Every detector works on the *concatenated* text of a paragraph (all ``w:t`` and
``w:instrText`` nodes joined), because Word splits names and numbers across runs
and drops the space at the seam: ``Kushal`` + ``Subbayya Hegde`` reads as
``KushalSubbayya Hegde``.  Matchers are therefore space tolerant.

Detection order matters.  Simple pattern detectors run first, then names, then
companies; addresses run last because they use the already-accepted name and
company spans as left boundaries (an address glued to its issuer, e.g.
"ICICI Securities Limited ICICI Venture House ...", must not swallow the
company name).  A final resolver keeps the highest priority type on any
overlap.
"""

import re
import unicodedata
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

# --------------------------------------------------------------------------
# PII taxonomy
# --------------------------------------------------------------------------

EMAIL = "EMAIL"
PHONE = "PHONE"
PERSON_NAME = "PERSON_NAME"
COMPANY = "COMPANY"
ADDRESS = "ADDRESS"
SSN = "SSN"
CREDIT_CARD = "CREDIT_CARD"
DATE_OF_BIRTH = "DATE_OF_BIRTH"
IP_ADDRESS = "IP_ADDRESS"
DIN = "DIN"
URL = "URL"

# Higher wins when two candidates overlap.  An address beats the company name
# printed inside it; a name beats the company-ish run it sits next to.
PRIORITY: Dict[str, int] = {
    ADDRESS: 100,
    PERSON_NAME: 90,
    EMAIL: 85,
    COMPANY: 80,
    PHONE: 70,
    DIN: 60,
    URL: 58,
    CREDIT_CARD: 55,
    SSN: 54,
    IP_ADDRESS: 50,
    DATE_OF_BIRTH: 45,
}

ALL_TYPES = (
    ADDRESS, PERSON_NAME, EMAIL, COMPANY, PHONE, DIN, URL,
    CREDIT_CARD, SSN, IP_ADDRESS, DATE_OF_BIRTH,
)


class Candidate:
    __slots__ = ("part", "paragraph", "start", "end", "ptype", "text",
                 "source", "group", "entity")

    def __init__(self, part: str, paragraph: int, start: int, end: int,
                 ptype: str, text: str, source: str = "regex", group: str = "",
                 entity: str = ""):
        self.part = part
        self.paragraph = paragraph
        self.start = start
        self.end = end
        self.ptype = ptype
        self.text = text
        self.source = source
        self.group = group
        #: The whole entity, when this candidate is only a piece of one.  A
        #: table cell can hold "Distriparks" out of "KSH Distriparks Private
        #: Limited"; the pieces must share one pseudonym, so they carry the
        #: joined value here.
        self.entity = entity or text

    @property
    def key(self) -> Tuple[str, int, int, int]:
        return (self.part, self.paragraph, self.start, self.end)

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return "Candidate(%s, %s|%d, %d:%d, %r, %s)" % (
            self.ptype, self.part, self.paragraph, self.start, self.end,
            self.text[:48], self.source,
        )


class Unit:
    """One paragraph of one part: the surface text detection runs on."""

    __slots__ = ("part", "index", "text", "cell")

    def __init__(self, part: str, index: int, text: str, cell: int = -1):
        self.part = part
        self.index = index
        self.text = text
        #: Index of the enclosing table cell, or ``-1`` for body text.
        self.cell = cell


class ScanContext:
    def __init__(self, units: List[Unit]):
        self.units = units
        self._by_part: Dict[str, List[Unit]] = {}
        for unit in units:
            self._by_part.setdefault(unit.part, []).append(unit)

    def parts(self) -> List[str]:
        return list(self._by_part)

    def units_of(self, part: str) -> List[Unit]:
        return self._by_part.get(part, [])

    def position(self, unit: Unit) -> int:
        return self.units_of(unit.part).index(unit)


# --------------------------------------------------------------------------
# Vocabulary
# --------------------------------------------------------------------------

#: Organisations that are regulators, exchanges or government bodies.  Per the
#: stated scope these are *not* redacted.
KEEP_ORGS: Set[str] = {
    "sebi", "bse", "nse", "roc", "mca", "nsdl", "cdsl",
    "securities and exchange board of india",
    "securities and exchange board",
    "bse limited", "national stock exchange of india limited",
    "national stock exchange", "stock exchanges",
    "national securities depository limited",
    "central depository services limited",
    "registrar of companies", "ministry of corporate affairs",
    "government of india", "central government", "state government",
    "government of maharashtra", "reserve bank of india", "rbi",
    "financial express", "jansatta", "loksatta", "economic times",
    "indian accounting standards", "indian gaap", "us gaap",
}

#: Definitional shorthand.  These end in a legal suffix but name no entity.
COMPANY_DENY_EXACT: Set[str] = {
    "company", "our company", "the company", "certain corp", "corp",
    "private limited", "limited", "public limited company",
    "sponsor bank", "refund bank", "escrow collection bank",
    "bankers to the issue", "bankers to the offer", "banker",
    "issuing bank", "bank", "bankers", "lead manager", "lead managers",
    "book running lead manager", "book running lead managers",
    "share transfer agent", "share transfer agents", "registrar to the offer",
    "designated stock exchange", "stock exchange", "stock exchanges",
    "total income", "total inc", "gross national disposable income",
    "government inc", "production linked income",
}

#: Tokens that may never *start* a company span.  Used to stop the
#: left-expansion of "Offer Escrow Collection Bank HDFC Bank Limited" at HDFC.
COMPANY_STOP_BEFORE: Set[str] = {
    "the", "a", "an", "formerly", "our", "its", "their", "such", "certain",
    "any", "each", "every", "all", "gross", "total", "net", "other", "offer",
    "escrow", "collection", "refund", "sponsor", "public", "issue", "account",
    "bank", "banker", "bankers", "company", "designated", "book", "running",
    "lead", "manager", "managers", "syndicate", "member", "members",
    "underwriter", "registrar", "transfer", "agent", "agents", "stock",
    "exchange", "exchanges", "government", "in", "of", "for", "from", "to",
    "by", "at", "on", "and", "as", "is", "be", "been", "by", "with", "will",
    "have", "has", "had", "was", "were", "that", "which", "who", "name",
    "names", "acting", "acting", "sole", "principal", "lead", "bankers",
}

#: Word sequences that introduce a legal entity name.
COMPANY_SUFFIXES: Tuple[str, ...] = (
    "Private Limited", "Limited", "LLP", "Inc.", "Inc", "Corporation",
    "Incorporated", "Ltd.", "Ltd",
)

#: Titles / designations.  Used as role anchors (a gazetteer-free way to find
#: person names) and as a structural signal in the Board table.
DESIGNATIONS: Tuple[str, ...] = (
    "Chairman", "Executive Director", "Managing Director", "Joint Managing Director",
    "Whole-time Director", "Whole Time Director", "Independent Director",
    "Non-Executive Director", "Director", "Company Secretary", "Compliance Officer",
    "Chief Executive Officer", "CEO", "Chief Financial Officer", "CFO",
    "Chief Operating Officer", "Technical Director", "President", "Vice President",
    "Company Secretary and Compliance Officer", "Executive Director and Chief",
)

#: Markers that a token belongs to a postal address.
ADDRESS_MARKERS: Set[str] = {
    "road", "marg", "lane", "street", "marg.", "nagar", "colony", "society",
    "soc", "apartment", "apartments", "bunglow", "bungalow", "villa", "park",
    "complex", "chambers", "centre", "center", "plaza", "gardens", "estate",
    "heights", "residency", "mansion", "layout", "cross", "avenue", "place",
    "terrace", "tower", "building", "wing", "floor", "block", "plot", "gat",
    "survey", "village", "taluka", "district", "dist", "unit", "flat", "house",
    "premises", "off", "near", "next", "opposite", "behind", "beside", "above",
    "co-operative", "housing", "soc.", "no", "nos", "no.", "s", "s.", "s.no",
    "&", "and", "at", "from", "1st", "2nd", "3rd", "4th", "5th", "6th", "7th",
    "8th", "9th", "10th", "11th", "12th", "a", "b", "c", "side", "street",
}

#: Cities / localities that appear in Indian postal addresses.
ADDRESS_CITIES: Set[str] = {
    "pune", "mumbai", "bhopal", "baner", "bandra", "kanjurmarg", "akurdi",
    "vikhroli", "churchgate", "parel", "thane", "navi", "delhi", "gurugram",
    "gurgaon", "bengaluru", "bangalore", "chennai", "hyderabad", "ahmedabad",
    "surat", "jaipur", "kolkata", "ahilyanagar", "ahmednagar", "chakan",
    "khed", "birdewadi", "supa", "parner", "palve", "khurd", "mauje",
    "taloja", "padghe", "panvel", "raigad", "khalumbre", "kothrud", "koregaon",
    "shivajinagar", "shivaji", "nagar", "deccan", "gymkhana", "erandawane",
    "pashan", "panchvati", "huzur", "govindpura", "model", "prabhat",
    "prabhadevi", "bhonde", "railyard", "railway", "station", "monarch",
    "world", "capital", "kothrud", "sadashiv", "peth", "high", "main", "old",
    "baner", "mahalunge", "balewadi", "sus", "pashan", "aundh", "bhosari",
}

#: State / country tail of an address.
ADDRESS_TAIL: Tuple[str, ...] = (
    "maharashtra", "madhya pradesh", "gujarat", "karnataka", "tamil nadu",
    "telangana", "west bengal", "uttar pradesh", "delhi", "haryana", "punjab",
    "rajasthan", "odisha", "assam", "bihar", "jharkhand", "goa", "kerala",
    "india",
)

#: "Maharashtra, India" is the tail of an address block when it is the whole
#: paragraph (a table cell).  The same two words inside a sentence ("banks in
#: Maharashtra, India are open for business") name a location, not an address.
BARE_REGION_RE = re.compile(r"(?i)^[A-Z][a-z]+(?: [A-Z][a-z]+)?,\s*India\.?$")

#: Phrases whose presence means the paragraph is prose, not an address block.
ADDRESS_STOP_PHRASES: Tuple[str, ...] = (
    "counsel", "auditor", "engineer", "bankers to", "as to indian law",
    "namely", "being", "secretary", "officer", "consent", "certificate",
    "chartered accountants", "company secretary", "below:", "set forth",
    "practicing", "practising",
)

#: Tokens that make a capitalised run an organisation, not a person.
ORGANISATION_HINT_RE = re.compile(
    r"(?i)\b(?:TRUST|FAMILY|HUF|FOUNDATION|FUND|PARTNERS|VENTURES?|HOLDINGS?|"
    r"CAPITAL|ASSET|MANAGEMENT|SECURITIES|BANK|LIMITED|PRIVATE|LLP|INC|CORPORATION|"
    r"INDUSTRIES|INDUSTRIAL|ENGINEERING|ASSOCIATES|ENTERPRISES|SOLUTIONS|"
    r"TECHNOLOGIES|INDIA)\b"
)

#: Function words that show a capitalised run is prose or a business term, not
#: a personal name.  "THE REGIONAL LANGUAGE OF", "Finance Act", "Equity Shares".
NAME_STOPWORDS = frozenset("""
    the of and or in to for from with within into under over also such any all
    each shall should would could may might must have has had been being is are
    was were these those there here then than when whom whose which whither
    whereas pursuant given hereinafter provided whereby
    shares securities share promoters directors director shareholders kmp kmps
    ssm auditors administrator administrators parties counsel bankers leaders
    members members employees beneficiaries partners partner company group
    trust trusts consultants underwriters managing joint finance act officer
    officers secretary chairman chief executive general head vice president
    nominee representative trustee alternate additional independent
    non-executive independent director whole only mere
""".split())

#: A legal-entity qualifier that turns a gazetteer name into an organisation:
#: "Karunakar Hegde HUF", "Rajesh Hegde Family Trust".
ORG_QUALIFIER_RE = re.compile(
    r"\s*(?:HUF|HUF\b|Family\s+Trust|Family\s+HUF|Trust\b|Holdings?\b|"
    r"Trustees\b|Foundation\b)", re.IGNORECASE
)

#: Text that must never be treated as an address.
ADDRESS_NEGATIVE = re.compile(
    r"(?i)\b(?:page|regulation|section|schedule|form|rounded|total|share|"
    r"price band|crore|million|billion|per cent|%)\b"
)


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def _norm_key(text: str) -> str:
    """Normalisation used to decide that two mentions are the same entity."""
    return re.sub(r"\s+", " ", text).strip().lower()


def _space_tolerant(pattern: str) -> str:
    """Turn ``\\s+`` in a pattern into ``[ \\u00a0]*`` so run seams do not break it.

    A Word run seam removes the space entirely (``KushalSubbayya``); the
    pattern must therefore also match with no space at all.
    """
    return pattern.replace(r"\s+", r"[ \u00a0]*")


def _clean_span(text: str, start: int, end: int) -> Tuple[int, int]:
    """Shrink a span so it does not start or end on whitespace or a connector."""
    while start < end and (text[start].isspace() or text[start] in ",;|-&/"):
        start += 1
    while end > start and (text[end - 1].isspace() or text[end - 1] in ",;|-&/"):
        end -= 1
    # A leading connector is the tail of the previous list item, not part of
    # this name: "Kanchenjunga Family Trust and Waterloo Industrial Park VI
    # Private Limited" ends at the trust, so the name starts after "and".
    # A connector *inside* a name ("CARE Analytics and Advisory Private
    # Limited") is left alone - only a leading one is trimmed.
    while True:
        word = re.match(r"[ \u00a0]*(and|of|&)[ \u00a0]+", text[start:end], re.I)
        if word is None:
            break
        head = text[start + word.end():]
        if not head[:1].isupper():
            break
        start += word.end()
    return start, end


def _split_pieces(segment: str, base: int) -> List[Tuple[str, int]]:
    """Split "A/ B /C" into pieces, keeping each piece's offset in the paragraph.

    ``re.split`` loses positions, which silently maps every name in a
    "Name1/Name2" contact line onto the first name's position.
    """
    pieces: List[Tuple[str, int]] = []
    cursor = 0
    for piece in SUFFIX_SPLIT.split(segment):
        found = segment.find(piece, cursor) if piece else cursor
        if found < 0:
            found = cursor
        stripped = piece.strip()
        lead = len(piece) - len(piece.lstrip())
        cleaned = SPLIT_LABEL.sub("", stripped)
        # SPLIT_LABEL only removes a trailing connector, so the head offset
        # does not move; the piece still starts where the raw piece did.
        pieces.append((cleaned, base + found + lead))
        cursor = found + len(piece)
    return pieces


def looks_like_organisation_name(text: str) -> bool:
    """True when a capitalised run is a legal entity rather than a person."""
    lowered = _norm_key(text)
    if lowered in COMPANY_DENY_EXACT or lowered in KEEP_ORGS:
        return True
    if any(text.rstrip().endswith(suffix) for suffix in COMPANY_SUFFIXES):
        return True
    return bool(ORGANISATION_HINT_RE.search(text))


# --------------------------------------------------------------------------
# EMAIL
# --------------------------------------------------------------------------

EMAIL_RE = re.compile(
    r"(?<![A-Za-z0-9._%+\-])"
    r"[A-Za-z0-9](?:[A-Za-z0-9._%+\-]*[A-Za-z0-9])?"
    r"@[A-Za-z0-9](?:[A-Za-z0-9\-]*[A-Za-z0-9])?"
    r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9\-]*[A-Za-z0-9])?)*"
    r"\.[A-Za-z]{2,24}"
    r"(?![A-Za-z0-9\-])"
)


class EmailDetector:
    ptype = EMAIL
    name = "email"

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for unit in ctx.units:
            for match in EMAIL_RE.finditer(unit.text):
                found.append(Candidate(
                    unit.part, unit.index, match.start(), match.end(),
                    EMAIL, match.group(0), "regex",
                ))
        return found


# --------------------------------------------------------------------------
# URL
# --------------------------------------------------------------------------

#: A web address, either with a scheme or written as "www.x.y".  A bare host
#: ("nuvama.com") is *not* matched: inside an e-mail address it belongs to the
#: e-mail candidate, and on its own it is indistinguishable from prose.
#: Word splits a host at a run seam - "www.kshinternational. com" - so one
#: space after a dot is tolerated, and the last label has to be a real top
#: level domain so that "www. of the. report" stays prose.
URL_TLD = (
    "com|net|org|edu|gov|mil|int|info|biz|io|co|in|uk|us|eu|au|ca|de|fr|jp|"
    "cn|ch|nl|se|no|dk|fi|be|at|it|es|pt|gr|il|tr|za|br|mx|ru|kr|tw|hk|sg|"
    "ae|sa|ng|ke|gh|nz|pl|cz|hu|ro|ph|my|th|vn|ar|cl|co|pe"
)
URL_RE = re.compile(
    r"(?i)\b(?:https?://|www\.)"
    r"[a-z0-9][a-z0-9\-]*(?:\.[ ]?[a-z0-9][a-z0-9\-]*)*"
    r"\.[ ]?(?:" + URL_TLD + r")\b"
    r"(?:\.[ ]?(?:" + URL_TLD + r"))?"
    r"(?:/[^\s\"'<>,;]*)?"
)


#: Hosts that belong to a regulator or an exchange.  Their web sites are
#: public reference, not a party's contact details, so they stay.
KEEP_URL_HOSTS = (
    "sebi.gov.in", "bseindia.com", "nseindia.com", "cdslindia.com",
    "nsdl.co.in", "rbi.org.in", "mca.gov.in", "indiacode.nic.in",
    "nclt.gov.in", "ibbi.gov.in", "npstrust.org.in", "income-tax.gov.in",
)


def _url_host(value: str) -> str:
    """The bare host of a URL, without the scheme, "www." or any path."""
    rest = value
    for prefix in ("https://", "http://"):
        if rest.lower().startswith(prefix):
            rest = rest[len(prefix):]
            break
    rest = rest.split("/", 1)[0]
    rest = rest.replace(" ", "").lower()
    if rest.startswith("www."):
        rest = rest[4:]
    return rest


class UrlDetector:
    """Web addresses, which carry the entity's identity in the host name.

    "www.kshinternational.com" and the Mimecast tracking wrapper
    "https://url.uk.m.mimecastprotect.com/...?domain=nuvama.com" both name the
    organisation the rest of the document had to be redacted for.
    """

    ptype = URL
    name = "url"

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for unit in ctx.units:
            for match in URL_RE.finditer(unit.text):
                value = match.group(0)
                if value.lower().startswith("www.") and "@" in value:
                    continue
                if _url_host(value) in KEEP_URL_HOSTS:
                    continue
                found.append(Candidate(
                    unit.part, unit.index, match.start(), match.end(),
                    URL, value, "regex",
                ))
        return found


# --------------------------------------------------------------------------
# PHONE
# --------------------------------------------------------------------------

#: "+ 91", "+91", "+91-" ... then a run of digits and Indian grouping marks.
INTERNATIONAL_START = re.compile(r"\+\s*91[\s\-.]*")
#: STD landline written with a separator: 022-68052182.
STD_RE = re.compile(r"(?<![\w\-])0\d{2,4}[\s\-]\d{8}(?![\d])")
PHONE_CONTEXT = re.compile(
    r"(?i)\b(?:telephone|tel|mobile|phone|fax|contact\s*person|contact\s*no)"
    r"\s*\.?\s*:?\s*$"
)


class PhoneDetector:
    ptype = PHONE
    name = "phone"

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for unit in ctx.units:
            found.extend(self._international(unit))
            found.extend(self._std(unit))
        found.sort(key=lambda c: (c.part, c.paragraph, c.start))
        return found

    @staticmethod
    def _digits_only(text: str) -> str:
        return "".join(ch for ch in text if ch.isdigit())

    def _international(self, unit: Unit) -> List[Candidate]:
        text = unit.text
        out: List[Candidate] = []
        for match in INTERNATIONAL_START.finditer(text):
            start = match.start()
            cursor = match.end()
            while cursor < len(text) and (text[cursor].isdigit() or text[cursor] in " -()."):
                cursor += 1
            body = text[match.end():cursor]
            digits = self._digits_only(body)
            if len(digits) < 10:
                continue
            if len(digits) > 10:
                # Keep the last ten national digits and move the start with them.
                drop = len(digits) - 10
                seen = 0
                offset = 0
                while offset < len(body) and seen < drop:
                    if body[offset].isdigit():
                        seen += 1
                    offset += 1
                start = match.end() + offset
            # The span ends immediately after the last digit of *this* number,
            # so a trailing full stop or comma, and any following number, are
            # not swallowed.
            last = -1
            for index in range(min(cursor, len(text)) - 1, start - 1, -1):
                if text[index].isdigit():
                    last = index
                    break
            if last == -1:
                continue
            start, end = _clean_span(text, start, last + 1)
            if end <= start:
                continue
            out.append(Candidate(unit.part, unit.index, start, end, PHONE,
                                 text[start:end], "regex"))
        return out

    def _std(self, unit: Unit) -> List[Candidate]:
        out: List[Candidate] = []
        for match in STD_RE.finditer(unit.text):
            start, end = _clean_span(unit.text, match.start(), match.end())
            out.append(Candidate(unit.part, unit.index, start, end, PHONE,
                                 unit.text[start:end], "regex"))
        return out


# --------------------------------------------------------------------------
# GOVERNMENT / FINANCIAL IDENTIFIERS
# --------------------------------------------------------------------------

SSN_RE = re.compile(r"(?<!\d)\d{3}-\d{2}-\d{4}(?!\d)")
CARD_CANDIDATE = re.compile(r"(?<![\d.])(?:\d[ -]?){12,18}\d(?![\d])")
IPV4_RE = re.compile(
    r"(?<![\d.])(?:25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)"
    r"(?:\.(?:25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}(?![\d.])"
)
DOB_KEYWORDS = re.compile(
    r"(?i)\b(?:date\s+of\s+birth|d\.?o\.?b\.?|born\s+on|age(?:d)?\s*:?\s*\d{1,3})\b"
)
DATE_TOKEN = re.compile(
    r"\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}"
    r"|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}"
    r"|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?,?\s+\d{4})\b"
)
DIN_CONTEXT = re.compile(r"(?i)\bDIN\b")
DIGIT_RUN = re.compile(r"(?<![\d])\d{8}(?![\d])")


def luhn_ok(digits: str) -> bool:
    total = 0
    for index, char in enumerate(reversed(digits)):
        value = int(char)
        if index % 2 == 1:
            value *= 2
            if value > 9:
                value -= 9
        total += value
    return total % 10 == 0


class SsnDetector:
    ptype = SSN
    name = "ssn"

    def find(self, ctx: ScanContext) -> List[Candidate]:
        return [
            Candidate(u.part, u.index, m.start(), m.end(), SSN, m.group(0))
            for u in ctx.units for m in SSN_RE.finditer(u.text)
        ]


class CreditCardDetector:
    ptype = CREDIT_CARD
    name = "credit_card"

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for unit in ctx.units:
            for match in CARD_CANDIDATE.finditer(unit.text):
                raw = match.group(0)
                digits = "".join(ch for ch in raw if ch.isdigit())
                if not 13 <= len(digits) <= 19 or not luhn_ok(digits):
                    continue
                found.append(Candidate(unit.part, unit.index, match.start(),
                                       match.end(), CREDIT_CARD, raw))
        return found


class IpAddressDetector:
    ptype = IP_ADDRESS
    name = "ip_address"

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for unit in ctx.units:
            for match in IPV4_RE.finditer(unit.text):
                octets = match.group(0).split(".")
                # A version-like or year-like run is not an address.
                if all(int(o) < 10 for o in octets) and match.group(0).count(".") == 3:
                    continue
                found.append(Candidate(unit.part, unit.index, match.start(),
                                       match.end(), IP_ADDRESS, match.group(0)))
        return found


class DateOfBirthDetector:
    ptype = DATE_OF_BIRTH
    name = "date_of_birth"

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for unit in ctx.units:
            for keyword in DOB_KEYWORDS.finditer(unit.text):
                window = unit.text[keyword.end(): keyword.end() + 40]
                for match in DATE_TOKEN.finditer(window):
                    start = keyword.end() + match.start()
                    end = keyword.end() + match.end()
                    found.append(Candidate(unit.part, unit.index, start, end,
                                           DATE_OF_BIRTH, unit.text[start:end]))
        return found


class DinDetector:
    """Director Identification Number.

    Only fires on a bare 8 digit paragraph (the DIN column of the Board table)
    or next to an explicit DIN label.  Bare 8 digit runs elsewhere in a
    prospectus are share counts and registration numbers, not PII.
    """

    ptype = DIN
    name = "din"

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for part in ctx.parts():
            units = ctx.units_of(part)
            for position, unit in enumerate(units):
                text = unit.text.strip()
                if re.fullmatch(r"\d{8}", text):
                    offset = unit.text.index(text)
                    found.append(Candidate(
                        part, unit.index, offset, offset + len(text), DIN, text,
                        "bare-8-digit",
                    ))
                    continue
                if DIN_CONTEXT.search(text):
                    for match in DIGIT_RUN.finditer(text):
                        found.append(Candidate(part, unit.index, match.start(),
                                               match.end(), DIN, match.group(0)))
        return found


# --------------------------------------------------------------------------
# PERSON NAMES
# --------------------------------------------------------------------------

CONTACT_PERSON_RE = re.compile(
    r"(?i)\bcontact\s*persons?\s*[:.]?\s*(?P<names>[^;|\n]{2,120})"
)
NAME_TOKEN = r"[A-Z][A-Za-z.'\u2019-]*"
NAME_RUN = re.compile(
    r"(?<![A-Za-z])(?:[A-Z]\.|[A-Z][a-z]+)"
    r"(?:[ \u00a0]*(?:[A-Z]\.|[A-Z][a-z]+)){1,3}"
)
SUFFIX_SPLIT = re.compile(r"\s*/\s*")
TRAILING_ROLE = re.compile(
    r"(?i)[,\s]*(?:company secretary(?: and compliance officer)?|compliance officer|"
    r"chairman(?: and executive director)?|managing director|executive director|"
    r"independent director|whole-?time director|chief executive officer|"
    r"chief financial officer|ceo|cfo)\b.*$"
)
LEADING_LABEL = re.compile(r"(?i)^.*?(?:contact\s*persons?\s*[:.]?\s*)")
SPLIT_LABEL = re.compile(r"(?i)\b(?:or|and)\b\s*$")


def _looks_like_name(run: str) -> bool:
    """Reject capitalised runs that are prose or business vocabulary."""
    tokens = run.split()
    if not 2 <= len(tokens) <= 4:
        return False
    if any(len(tok) == 1 and not tok.endswith(".") for tok in tokens):
        return False
    for token in tokens:
        bare = token.strip(".,'").lower()
        if bare in NAME_STOPWORDS or bare in ("ltd", "llp", "inc", "pvt", "sebi",
                                             "bse", "nse", "roc", "mca"):
            return False
        if looks_like_organisation_name(run):
            return False
    return True


class LabelNameDetector:
    """A personal name used as a label: "Rashi Patil: John Doe".

    This is the shape the assignment's own example uses, and it is the one
    name shape the other detectors have no reason to fire on.  A capitalised
    bigram on its own is not evidence of a name - in a prospectus "Equity
    Shares", "Red Herring Prospectus" and "Stock Exchanges" outnumber the
    people thousands to one, which is why no bare-bigram name rule ships.
    The evidence here is the *pair*: a capitalised run, a colon, and another
    capitalised run.  That is what a name-to-name mapping looks like, and
    offering vocabulary does not look like it: "Corporate Identity Number:
    U28129PN1979PLC141039", "Order Id: A-5561" and "Price Band: 98-100" are
    all left alone, because what follows the colon is not a second run.

    Only the left side is reported.  In the brief's example the left side is
    the real person and the right side is the fake to substitute, so
    redacting both would replace one name with another.
    """

    ptype = PERSON_NAME
    name = "name/label-pair"

    PAIR = re.compile(
        NAME_RUN.pattern + r"[ \u00a0]*:[ \u00a0]*" + NAME_RUN.pattern)
    #: "Contact Person: Eric Bacha", "Book Running Lead Managers: X" - the
    #: left side is a role label, not a name.  The contact-anchor detector
    #: owns the first and reports the name on the right, so the pair rule
    #: must not report the label as a person.
    LABEL_TAIL = frozenset("""
        person persons contact name names details detail address phone
        email website signatory nominee
        manager managers banker bankers auditor auditors adviser advisers
        advisor advisors representative representatives consultant
        consultants underwriter underwriters trustee trustees officer
        officers director directors secretary leaders partner partners
    """.split())

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for unit in ctx.units:
            for match in self.PAIR.finditer(unit.text):
                left = match.group(0).split(":")[0].strip()
                tokens = left.split()
                if not 2 <= len(tokens) <= 4:
                    continue
                if tokens[-1].strip(".,'\u2019").lower() in self.LABEL_TAIL:
                    continue
                start, end = _clean_span(unit.text, match.start(),
                                         match.start() + len(left))
                if end <= start:
                    continue
                found.append(Candidate(unit.part, unit.index, start, end,
                                       PERSON_NAME, unit.text[start:end],
                                       "label-pair"))
        return found


class ContactPersonNameDetector:
    """Gazetteer-free: names printed after an explicit "Contact Person" label."""

    ptype = PERSON_NAME
    name = "name/contact-anchor"

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for unit in ctx.units:
            for match in CONTACT_PERSON_RE.finditer(unit.text):
                segment = match.group("names")
                base = match.start("names")
                segment = TRAILING_ROLE.sub("", segment)
                if re.search(r"(?i)HYPERLINK|SEBI\s+Registration|E-?mail|Website", segment):
                    break
                for piece, piece_base in _split_pieces(segment, base):
                    piece = SPLIT_LABEL.sub("", piece.strip()).strip()
                    if not piece:
                        continue
                    for run in NAME_RUN.finditer(piece):
                        candidate_text = run.group(0)
                        if not _looks_like_name(candidate_text):
                            continue
                        start = piece_base + run.start()
                        end = piece_base + run.end()
                        start, end = _clean_span(unit.text, start, end)
                        if end > start:
                            found.append(Candidate(
                                unit.part, unit.index, start, end, PERSON_NAME,
                                unit.text[start:end], "contact-anchor",
                            ))
        return found


class RoleAnchorNameDetector:
    """Gazetteer-free: "being <Name>", "namely, <Name>", "<Name> is our <Role>"."""

    ptype = PERSON_NAME
    name = "name/role-anchor"

    ANCHORS = (
        # Scoped (?i:...) on the keyword only: the name part must stay
        # case sensitive or "our principal raw material" matches.
        re.compile(r"\b(?i:being|namely,?|is our|are our)\s+(?P<name>" + NAME_TOKEN +
                   r"(?:[ \u00a0]*" + NAME_TOKEN + r"){1,3})"),
        re.compile(r"(?P<name>" + NAME_TOKEN + r"(?:[ \u00a0]*" + NAME_TOKEN +
                   r"){1,3})\s+(?i:is|are)\s+our\s+(?i:company\s+secretary|"
                   r"chairman|managing director|executive director|promoter)"),
    )

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for unit in ctx.units:
            for pattern in self.ANCHORS:
                for match in pattern.finditer(unit.text):
                    name = match.group("name")
                    if looks_like_organisation_name(name):
                        continue
                    if not _looks_like_name(name):
                        continue
                    start, end = _clean_span(unit.text, match.start("name"),
                                             match.end("name"))
                    if end > start:
                        found.append(Candidate(
                            unit.part, unit.index, start, end, PERSON_NAME,
                            unit.text[start:end], "role-anchor",
                        ))
        return found


class DesignationTableNameDetector:
    """Gazetteer-free: a short capitalised cell whose neighbour is a designation.

    This is how the Board table is structured (Name | Designation | DIN |
    Address), so the signal is structural rather than a list of names.
    """

    ptype = PERSON_NAME
    name = "name/designation-table"

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        designation = re.compile(
            r"(?i)\b(?:" + "|".join(re.escape(d) for d in DESIGNATIONS) + r")\b"
        )
        for part in ctx.parts():
            units = ctx.units_of(part)
            for position, unit in enumerate(units):
                text = unit.text.strip()
                if not text or len(text) > 60:
                    continue
                if not NAME_RUN.fullmatch(text):
                    continue
                if not _looks_like_name(text):
                    continue
                if looks_like_organisation_name(text):
                    continue
                window = [u.text.strip() for u in units[position + 1: position + 3]]
                # The designation must *be* the next cell, not merely appear
                # somewhere in the row, or "Joint Managing Director" followed
                # by a designation cell is mistaken for a name.
                if not any(designation.match(cell) for cell in window if cell):
                    continue
                start = unit.text.index(text)
                found.append(Candidate(part, unit.index, start,
                                       start + len(text), PERSON_NAME, text,
                                       "designation-table"))
        return found


class AllCapsListNameDetector:
    """Gazetteer-free: the cover banner, "OUR PROMOTERS: A NAME, B NAME, ...".

    Only fires after an explicit all-caps label so that ordinary all-caps
    headings ("RISKS IN RELATION TO THE FIRST OFFER") are never touched.
    """

    ptype = PERSON_NAME
    name = "name/all-caps-list"

    LABEL = re.compile(
        r"(?i)\b(?:our\s+)?(?:promoters?|promoter\s+selling\s+shareholders?|"
        r"directors?|individual\s+promoters?)\s*:\s*"
    )
    NAME = re.compile(r"[A-Z][A-Z]*(?:[ \u00a0]+[A-Z][A-Z]*){1,3}")

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for unit in ctx.units:
            for label in self.LABEL.finditer(unit.text):
                tail = unit.text[label.end():]
                if len(tail) > 400:
                    tail = tail[:400]
                for match in self.NAME.finditer(tail):
                    text = match.group(0)
                    # "… KANCHENJUNGA FAMILY TRUST AND" - the list runs on.
                    text = re.sub(r"(?:\s+(?:AND|OR|&))+$", "", text)
                    if looks_like_organisation_name(text):
                        continue
                    if not self._plausible(text):
                        continue
                    start = label.end() + match.start()
                    end = label.end() + match.start() + len(text)
                    start, end = _clean_span(unit.text, start, end)
                    if end > start:
                        found.append(Candidate(unit.part, unit.index, start, end,
                                               PERSON_NAME, unit.text[start:end],
                                               "all-caps-list"))
        return found

    @staticmethod
    def _plausible(text: str) -> bool:
        tokens = text.split()
        if not 2 <= len(tokens) <= 4:
            return False
        # A person name in this document is 2-4 tokens, none of them a
        # securities term longer than 6 characters.
        for token in tokens:
            if len(token) > 14:
                return False
            if token in {"THE", "OUR", "AND", "FOR", "SHARES", "EQUITY", "FACE",
                         "VALUE", "AGGREGATING", "MILLION", "LIMITED", "PRIVATE",
                         "COMPANY", "OFFER", "OFFICE", "SECRETARY", "DIRECTOR",
                         "AGAINST", "STOCK", "EXCHANGES", "SHARE", "HOLDER",
                         "HOLDERS", "PROCEEDS", "NET", "ISSUE", "FRESH"}:
                return False
        return True



class GazetteerNameDetector:
    """Curated list of the people named in this document.

    Disabled with ``--no-gazetteer``; the gazetteer-free numbers in the
    evaluation report come from that mode.
    """

    ptype = PERSON_NAME
    name = "name/gazetteer"

    def __init__(self, entries: Sequence[str]):
        # Longest first so "Kushal Subbayya Hegde" wins over "Kushal".
        self.entries = sorted(set(entries), key=lambda e: (-len(e), e))
        self._patterns: List[Tuple[str, re.Pattern]] = []
        for entry in self.entries:
            tokens = [re.escape(tok) for tok in entry.split()]
            # A run seam can delete the space between two tokens, so the
            # separator is optional: the gazetteer spells "Sunil Nagayya
            # Shetty" and the document prints "SunilNagayya Shetty".
            body = r"[ \u00a0]*".join(tokens)
            self._patterns.append((entry, re.compile(
                r"\b" + body + r"\b"
            )))
            if len(tokens) >= 2:
                # The cover page prints the promoters in capitals
                # ("KUSHAL SUBBAYYA HEGDE").  Two or more tokens is specific
                # enough to match case-insensitively; a single given name is
                # not, so it stays case-sensitive.
                self._patterns.append((entry, re.compile(
                    r"\b" + body + r"\b", re.IGNORECASE
                )))
        self.first_names = sorted({entry.split()[0] for entry in self.entries
                                   if 2 <= len(entry.split()) <= 4})

    def _family_patterns(self, ctx: ScanContext) -> List[Tuple[str, "re.Pattern"]]:
        """Given name + the surname this document actually gives it.

        A given name on its own is not enough: "Kushal" also opens the company
        "Kushal Motors and Electricals Private Limited", and a one token
        person span outranks the company span and breaks the company in half.
        A surname that the document pairs with a gazetteer given name in two
        places is a family name, though - "Rajesh Branch, Sangeeta Branch,
        Rakhi Branch and Rohit Branch" - so the pair is what is matched, and
        it is case sensitive so that "sarthak.malvadkar@kshinterantional.com"
        in a mail address is not a person.
        """
        pairs: Dict[str, int] = {}
        for unit in ctx.units:
            for name in self.first_names:
                for match in re.finditer(
                        r"(?<![A-Za-z])" + re.escape(name) + r"(?![a-z])",
                        unit.text):
                    tail = re.match(r"[ \u00a0]+([A-Z][A-Za-z'\u2019]*)",
                                    unit.text[match.end():])
                    if tail is None or tail.group(1) in LEGAL_FORM_ABBREV:
                        continue
                    pairs["%s %s" % (name, tail.group(1))] = \
                        pairs.get("%s %s" % (name, tail.group(1)), 0) + 1
        return [(pair, re.compile(r"\b" + r"[ \u00a0]*".join(
            re.escape(token) for token in pair.split()) + r"\b"))
            for pair, count in sorted(pairs.items()) if count >= 2]

    #: What follows a learned pair and means the name continues into a legal
    #: entity: "Kushal Motors and Electricals Private Limited", "Bank of
    #: Baroda Limited".  A person span here would outrank the company span.
    #: No ``^``: the search starts at the end of the match, and ``^`` does not
    #: match at an offset.  A legal form further along the same clause means
    #: the pair is the front of a company name ("Kushal Motors and Electricals
    #: Private Limited").  A list of people has no such tail ("Rakhi Branch
    #: and Rohit Branch"), so the conjunction alone is not a warning.
    _CONTINUES_ENTITY = re.compile(
        r"(?i)[^,;.()]{0,140}?\b(?:private|limited|llp|inc\.?|incorporated"
        r"|corporation|ltd\.?)\b"
    )

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        # The full entries first: they are the longest, so a learned family
        # pair such as "Kushal Subbayya" can never shorten "Kushal Subbayya
        # Hegde" to two tokens.
        patterns = self._patterns + self._family_patterns(ctx)
        for unit in ctx.units:
            taken: List[Tuple[int, int]] = []
            for index, (entry, pattern) in enumerate(patterns):
                learned = index >= len(self._patterns)
                for match in pattern.finditer(unit.text):
                    if any(match.start() < end and match.end() > start
                           for start, end in taken):
                        continue
                    # The guard is for a guess.  A gazetteer entry is a name
                    # that was read off the document, so "Narayna B. Shetty
                    # and Jayaram N. Shetty" keeps both.
                    if learned and self._CONTINUES_ENTITY.match(unit.text,
                                                                match.end()):
                        continue
                    end = match.end()
                    qualifier = ORG_QUALIFIER_RE.match(unit.text, end)
                    if qualifier:
                        # "Family Trust" is an organisation, not a person.
                        taken.append((match.start(), qualifier.end()))
                        found.append(Candidate(unit.part, unit.index,
                                               match.start(), qualifier.end(),
                                               COMPANY,
                                               unit.text[match.start():qualifier.end()],
                                               "gazetteer"))
                        continue
                    taken.append((match.start(), end))
                    found.append(Candidate(unit.part, unit.index, match.start(),
                                           end, PERSON_NAME, match.group(0),
                                           "gazetteer"))
        return found


# --------------------------------------------------------------------------
# COMPANIES
# --------------------------------------------------------------------------

#: A token that can be part of a legal-entity name.  A lowercase word ends
#: the walk; a digit does not - "KSH Infra Park 5 Private Limited".
NAME_LIKE_TOKEN = re.compile(r"[A-Z0-9]")
CONNECTOR_TOKEN = frozenset({"and", "of"})

#: Words that end a company name when they are read from the left.
COMPANY_LEFT_STOP = frozenset("""
    collectively however therefore additionally together hereby therein
    aforesaid respectively thereafter notwithstanding whereas henceforth
    provided also
    statutory auditors secretarial auditors
""".split())

#: Capitalised clause words that can never be part of a legal-entity name.
#: Words like "and", "&" or "or" are deliberately *not* here: "Kanj and Co"
#: is a real name, and the KEEP_ORGS test then recognises it and leaves it
#: alone.
COMPANY_NAME_STOP = frozenset("""
    the a an our its their such any each every all former both same other
    others said aforesaid hereby company
""".split())

#: Legal-suffix words.  A connector is not a joiner when the token on its left
#: is itself a suffix: "Kanj and Co LLP" is two companies.
SUFFIX_WORD = frozenset("""
    limited ltd llp inc incorporated private
""".split())

#: Abbreviations that take a trailing comma inside the name itself
#: ("Co.,"), as opposed to a list separator.
#: The tail of a name, reached from the paragraph before it.
NAME_TAIL_RE = re.compile(
    r"(?:Family|HUF|Holdings?|Trusts?)[ ]+(?:Trust|TRUST)\b", re.IGNORECASE
)

#: A company name is at most this long; anything longer is a list.
MAX_COMPANY_CHARS = 120

LEGAL_FORM_ABBREV = frozenset({
    "limited", "ltd", "ltd.", "private", "llp", "inc", "inc.", "incorporated",
    "corporation", "corp", "corp.", "co", "co.", "pvt", "pvt.",
})

#: A comma right after a legal form separates two *names*, it is not a
#: grammatical comma: "Kirtane & Pandit, LLP" is a list of two companies, not
#: one company with a long name.  Without this the leftward walk swallows the
#: whole list.
LEGAL_FORM_WORDS = frozenset({
    "limited", "ltd", "llp", "inc", "incorporated", "private", "corporation",
    "corp", "co", "company", "pvt",
})

#: Intermediaries are named in a role cell: "Bankers to the Offer".  The role
#: phrase is a hard left-hand boundary for a legal-suffix name.  The pattern
#: is *not* anchored to the end of the prefix: an anchored tail has to reach
#: the end of the string, so it swallows the company name that follows the
#: role words and the boundary is lost.
ROLE_PREFIX_RE = re.compile(
    r"(?i)(?:"
    r"book\s*running\s*lead\s*managers?|lead\s*managers?|arrangers?|"
    r"bankers?|editors?|registrars?|"
    r"(?:escrow\s*collection|public\s*offer\s*account|refund)\s*banks?|"
    r"sponsor\s*banks?|collection\s*agents?|clearing\s*banks?|"
    r"principal\s*accountants?|statutory\s*auditors?|central\s*advisors?|"
    r"legal\s*advisors?"
    r")(?:\s+to\s+the\s+offer)?"
    r"(?:\s+(?:" + r"book\s*running\s*lead\s*managers?|lead\s*managers?|"
    r"arrangers?|bankers?|editors?|registrars?|"
    r"(?:escrow\s*collection|public\s*offer\s*account|refund)\s*banks?|"
    r"sponsor\s*banks?|collection\s*agents?|clearing\s*banks?|"
    r"principal\s*accountants?|statutory\s*auditors?|central\s*advisors?|"
    r"legal\s*advisors?"
    r")(?:\s+to\s+the\s+offer)?)*"
)

#: How far after a role phrase a company name may start and still belong to
#: that cell ("Bankers to the Offer | HDFC Bank Limited").
ROLE_MAX_GAP = 60

#: How far after a role phrase a company name may start and still belong to
#: that cell ("Bankers to the Offer | HDFC Bank Limited").
ROLE_MAX_GAP = 60


def _joins_a_name(text: str, token_start: int, token_end: int) -> bool:
    """True for "Analytics and Advisory", false for "Company and our Promoters".

    A connector is part of a name only when capitalised tokens sit on both
    sides of it.
    """
    prefix = text[:token_start].rstrip(" \u00a0,&")
    left_token = prefix.rsplit(" ", 1)[-1] if prefix else ""
    if left_token.lower().strip(".,") in SUFFIX_WORD:
        return False  # "X Limited, and Y Limited" - the connector lists entities
    right_token = ""
    for char in text[token_end:].lstrip(" \u00a0,&"):
        if char in " .&'\u2019-":
            break
        right_token += char
    return (NAME_LIKE_TOKEN.match(left_token) is not None
            and NAME_LIKE_TOKEN.match(right_token) is not None)
SUFFIX_RE = re.compile(
    r"[ ]*,?[ ]*(?:Private[ ]+Limited|Limited|LLP|Inc\.?|Incorporated|"
    r"Corporation|Ltd\.?|"
    # All-caps spellings: the cover page and the promoter list print every
    # name in capitals ("KSH INTERNATIONAL LIMITED").
    r"PRIVATE[ ]+LIMITED|LIMITED|LLP|INC\.?|INCORPORATED|CORPORATION|LTD\.?)"
    # A form word directly followed by another form word is not the end of the
    # name: in "State Financial Corporation Limited", "Corporation" is not the
    # suffix, "Limited" is.  Without this the name is emitted twice.
    r"(?![\s,]+(?:Private[ ]+Limited|Limited|LLP|Inc\.?|Incorporated|"
    r"Corporation|Ltd\.?|PRIVATE[ ]+LIMITED|LIMITED|LLP|INC\.?|"
    r"INCORPORATED|CORPORATION|LTD\.?)(?![A-Za-z]))"
    r"(?![A-Za-z])"
)

#: "Annapurna Family Trust" - a promoter entity that carries no legal suffix.
#: The gazetteer lists the trusts that recur, but the promoter list introduces
#: the group, so the shape is matched too - in capitals as well, because the
#: same list is printed twice, once in title case and once in capitals.
_NAME_TOKEN = r"(?:[A-Z][a-z]{2,}|[A-Z]{3,})"
FAMILY_TRUST_RE = re.compile(
    r"(?<![A-Za-z])" + _NAME_TOKEN + r"(?:[ ]" + _NAME_TOKEN + r")?"
    r"[ ]+(?:Family|FAMILY)[ ]+(?:Trust|TRUST)(?![A-Za-z])"
)


class CompanyDetector:
    """Legal-suffix names, grown left over name-like tokens, stopped at roles."""

    ptype = COMPANY
    name = "company/legal-suffix"

    def __init__(self, boundary_re: "re.Pattern" = None):
        self.boundary_re = boundary_re

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for unit in ctx.units:
            text = unit.text
            for match in SUFFIX_RE.finditer(text):
                # The pattern swallows the comma and spaces in front of the
                # legal form ("Kirtane & Pandit, LLP"); the walk has to start
                # at the form itself, or that comma looks like a list
                # separator and the walk stops before the name.
                raw_suffix = match.group(0)
                word_start = match.start() + len(raw_suffix) - len(
                    raw_suffix.lstrip(" , "))
                start = word_start
                grown = self._grow_left(text, start, self.boundary_re,
                                        end=match.end())
                end = match.end()
                grown = self._grow_right(text, grown, end)
                if grown >= end:
                    continue
                start, end = _clean_span(text, grown, end)
                if end <= start:
                    continue
                value = text[start:end]
                if NAME_TAIL_RE.match(value):
                    # "Family Trust, Annapurna Family Trust, ... " starts in
                    # the middle of the trust named in the previous
                    # paragraph; a name does not begin with its own tail.
                    continue
                if len(value) > MAX_COMPANY_CHARS:
                    # A legal form inside a list of promoters turns the whole
                    # comma separated row into one "name" when no gazetteer
                    # marks the individual entities.  A real company name is
                    # short; a 400 character row is a list, and one pseudonym
                    # for it would also be laid out in the wrong paragraphs.
                    continue
                if _norm_key(value) in KEEP_ORGS or _norm_key(value) in COMPANY_DENY_EXACT:
                    continue
                if self._blocked(value):
                    continue
                found.append(Candidate(unit.part, unit.index, start, end,
                                       COMPANY, value, "legal-suffix"))
            for match in FAMILY_TRUST_RE.finditer(text):
                value = match.group(0)
                if _norm_key(value) in KEEP_ORGS or _norm_key(value) in COMPANY_DENY_EXACT:
                    continue
                found.append(Candidate(unit.part, unit.index, match.start(),
                                       match.end(), COMPANY, value, "family-trust"))
        return found

    @staticmethod
    def _grow_left(text: str, start: int, boundary_re: "re.Pattern" = None,
                  end: int = None) -> int:
        """Walk left over the tokens of a legal-entity name.

        Three boundaries matter: a role or clause word ("Collectively,",
        "Statutory Auditors,"), a token that is followed by a comma (a list
        separator: "Switchgear Co., Bharat Bijlee Limited"), and a lowercase
        word.  Lowercase connectors join the name instead of ending it, so
        "CARE Analytics and Advisory Private Limited" and "Kanj and Co LLP"
        are captured whole.
        """
        cursor = start
        # "Bankers to the Offer Escrow Collection Bank HDFC Bank Limited" -
        # the role phrase is a boundary, not a reason to give up: the company
        # name still starts right after it.
        role = ROLE_PREFIX_RE.search(text, 0, start)
        limit = role.end() if (role and start - role.end() <= ROLE_MAX_GAP) else 0
        while cursor > 0:
            previous = cursor
            while previous > 0 and text[previous - 1] in " \u00a0,&":
                previous -= 1
            token_start = previous
            while token_start > 0 and (text[token_start - 1].isalnum()
                                       or text[token_start - 1] in ".'\u2019-&"):
                token_start -= 1
            if token_start < limit:
                break
            raw = text[token_start:cursor]
            token = raw.strip(" \u00a0,&")
            if not token:
                if raw.strip(" \u00a0") == "&" and token_start > 0:
                    # "Kirtane & Pandit, LLP": the ampersand is inside the name
                    cursor = token_start
                    continue
                break
            lowered = token.lower()
            # A lowercase connector joins the name instead of ending it.
            if not NAME_LIKE_TOKEN.match(token):
                if lowered in CONNECTOR_TOKEN and _joins_a_name(text, token_start, cursor):
                    cursor = token_start
                    continue
                break
            if lowered in COMPANY_LEFT_STOP or lowered in COMPANY_NAME_STOP:
                break
            if cursor < len(text) and text[cursor] == ",":
                break
            # "…being, CARE Analytics and Advisory Private Limited" is a
            # grammatical comma and is crossed; "Switchgear Co., Bharat
            # Bijlee Limited" is a name that starts after the comma, and
            # "… Kanchenjunga Family Trust, and Waterloo Industrial Park VI
            # Private Limited" starts after a *known* entity.
            before = token_start - 1
            while before >= 0 and text[before] in " \u00a0":
                before -= 1
            if before >= 0 and text[before] == ",":
                word = before
                while word > 0 and (text[word-1].isalnum() or text[word-1] == "."):
                    word -= 1
                preceding = text[word:before].lower()
                if preceding in LEGAL_FORM_WORDS:
                    # end of the previous list item: this name starts here
                    cursor = token_start
                    break
                item_start = before + 1
                while item_start < len(text) and text[item_start] in " \u00a0":
                    item_start += 1
                item = boundary_re.match(text, item_start) if boundary_re else None
                if item is not None:
                    # The list item is a different known entity: this name
                    # starts where that entity ends.  A comma before an
                    # unknown item is not a boundary - "…, Kanj and Co LLP" is
                    # one name.  When the known entity *contains* the legal
                    # suffix we grew from, the two are the same name
                    # ("Kirtane & Pandit, LLP") and the name starts where the
                    # known entity does.
                    if item.end() <= start:
                        cursor = item.end()
                    else:
                        cursor = item.start()
                    break
                word = before
                while word > 0 and (text[word-1].isalnum() or text[word-1] == "."):
                    word -= 1
                if text[word:before].lower() in LEGAL_FORM_ABBREV:
                    cursor = item_start
                    break
            if len(token) > 40:
                break
            if boundary_re is not None and boundary_re.search(text, token_start, cursor):
                # The token is part of a *different* known entity, so the
                # legal-suffix name starts to the right of it.
                break
            cursor = token_start
        if boundary_re is not None:
            # A *known* entity that covers the position the walk stopped at
            # is a name of its own, and this name starts where it ends:
            # "Kanchenjunga Family Trust, and Waterloo Industrial Park VI
            # Private Limited".  The entity that ends the same way the legal
            # suffix does is the name we just matched, so it is left alone.
            limit = end if end is not None else cursor
            for known in boundary_re.finditer(text, max(0, cursor - 240)):
                if known.start() <= cursor < known.end() and known.end() < limit:
                    return known.end()
        return cursor

    @staticmethod
    def _grow_right(text: str, start: int, end: int) -> int:
        # "…Private Limited (formerly …)" - never swallow following prose.
        return start

    @staticmethod
    def _blocked(value: str) -> bool:
        tokens = value.split()
        if len(tokens) < 2:
            return True
        # A bare suffix with no distinctive part ("Private Limited").
        if all(token.lower() in {"private", "limited", "llp", "inc", "corporation"}
               for token in tokens):
            return True
        if tokens[0].lower() in {"the", "our", "a", "an", "formerly"}:
            return True
        if value.endswith(" Inc") and " ".join(tokens[:-1]).lower() in {
            "total", "government", "gross national disposable",
            "production linked",
        }:
            return True
        return False


class SplitEntityDetector:
    """An entity that a narrow table column broke across several cells.

    The related-party table prints "KSH", "Distriparks", "Private",
    "Limited" in four columns, and the promoter list breaks
    "DHAULAGIRI FAMILY TRUST" between two paragraphs.  Neither the
    paragraph text nor the paragraph regexes can see the name, so the
    consecutive cells are joined and matched against the gazetteer.  Every
    piece is returned as its own candidate carrying the joined value, so the
    pieces get one shared pseudonym.
    """

    ptype = COMPANY
    name = "company/split-cells"

    #: Longest run of cells a name may be broken across.
    WINDOW = 5

    #: Lowercase words that may appear inside a proper name.
    JOINERS = frozenset(("&", "and", "of", "the", "de", "&amp;"))

    #: Table roles and column labels.  One of these in a cell means the cell
    #: belongs to the table furniture, not to a name that spills over from
    #: the neighbouring column.
    HEADING_WORDS = frozenset("""
        account address anchor banker bankers book books brlm brlms
        consortium contact corporate custodian depositary director directors
        email lead leadbankers leadmanagers manager managers name
        promoter promoters registered registrar registrars representative
        shareholders sponsor sponsors syndicate telephone trustee trustees
        underwriter underwriters
    """.split())

    def __init__(self, config: "RedactionConfig"):
        self.config = config
        self._brands = None

    def _entries(self) -> List[Tuple[str, str]]:
        entries = [(value, COMPANY) for value in self.config.company_gazetteer]
        entries += [(value, PERSON_NAME) for value in self.config.person_gazetteer]
        # The legal form is not always in the gazetteer ("KSH Distriparks
        # Private Limited" recurs, "KSH Integrated Logistics Private
        # Limited" does not), so the shape is matched as well.
        return sorted(entries, key=lambda item: (-len(item[0]), item[0]))

    def _match(self, joined: str, entries):
        """Known entities that straddle a cell boundary, longest name first.

        Yields every candidate match; the caller keeps the first one that
        really is torn across two cells.
        """
        for entry, ptype in entries:
            tokens = [t for t in re.split(r"[ \t,]+", entry) if t]
            if len(tokens) < 2:
                continue
            pattern = r"[ \t,]*".join(re.escape(token) for token in tokens)
            for match in re.finditer(pattern, joined, re.I):
                if " " in match.group(0) or "," in match.group(0):
                    yield (match.start(), match.end(), ptype, match.group(0))

    def _brand_tokens(self) -> set:
        if self._brands is None:
            brands = set()
            for value in self.config.company_gazetteer:
                first = value.split(" ", 1)[0].strip(".,")
                if first:
                    brands.add(first.lower())
            self._brands = brands
        return self._brands

    @staticmethod
    def _one_cell(window) -> bool:
        """True when every line of the window belongs to the same cell."""
        cell = window[0].cell
        return cell >= 0 and all(unit.cell == cell for unit in window)

    def _shape_entity(self, cells: List[str], joined: str):
        """A legal-suffix name torn apart inside one table cell.

        The pieces must share a cell: "KSH" | "Distriparks" | "Private
        Limited" is one name printed as three lines of a single cell, while
        "Sponsor Banks" | "ICICI Bank Limited" is two cells and two names.
        Inside the cell the run must end in a legal form that starts in an
        earlier line, and no line may read as a table role or heading.
        """
        if not self._name_like(cells):
            return None
        last = cells[-1]
        if any(self._heading_word(cell) for cell in cells):
            return None
        last_offset = len(joined) - len(last) - 1
        for match in SUFFIX_RE.finditer(joined):
            if match.end() == len(joined) and match.start() <= last_offset:
                return joined
        return None

    @staticmethod
    def _heading_word(cell: str) -> bool:
        """True when a cell reads as a table role or column label."""
        tokens = {t.strip("().,").lower() for t in cell.split()}
        return bool(tokens & SplitEntityDetector.HEADING_WORDS)

    def _name_like(self, texts: Sequence[str]) -> bool:
        """Every cell looks like a fragment of a proper name.

        "Bankers to the Offer" is a column heading, not half a company
        name, and a lowercase function word is the giveaway.
        """
        for cell in texts:
            stripped = cell.strip()
            words = stripped.split()
            if not stripped or len(words) > 5:
                return False
            if re.search(r"[.;:!?]", stripped):
                return False
            for token in words:
                clean = token.strip("(),")
                if clean.lower() in self.JOINERS:
                    continue
                if not clean[:1].isupper() and not clean[:1].isdigit():
                    return False
        return True

    @staticmethod
    def _joins_name(left: str, right: str) -> bool:
        """Cheap pre-filter: could a name run across this cell boundary?

        A paragraph of prose meeting a table of figures ("... on page 250" |
        "0.00") is skipped before any gazetteer pattern is tried.
        """
        tail = left.rsplit(None, 1)[-1] if left.split() else ""
        head = right.split()[0] if right.split() else ""
        return (SplitEntityDetector._token_joins(tail)
                and SplitEntityDetector._token_joins(head))

    @staticmethod
    def _token_joins(token: str) -> bool:
        clean = token.strip("(),;:*&#“”\"'’-")
        if not clean:
            return True
        if clean.lower() in SplitEntityDetector.JOINERS:
            return True
        if clean.lower() in SUFFIX_WORD or clean.lower() in LEGAL_FORM_WORDS:
            return True
        return clean[:1].isupper() or clean[:1].isdigit()

    def find(self, ctx: ScanContext) -> List[Candidate]:
        # No gazetteer is fine: the joined cells are then matched on shape
        # alone ("KSH" | "Distriparks" | "Private" | "Limited").
        entries = self._entries()
        found: List[Candidate] = []
        for part in ctx.parts():
            units = ctx.units_of(part)
            for position in range(len(units)):
                for width in range(2, self.WINDOW + 1):
                    if position + width > len(units):
                        break
                    window = units[position: position + width]
                    cells = [u.text for u in window]
                    if any(not c.strip() for c in cells):
                        continue
                    if not all(self._joins_name(cells[i], cells[i + 1])
                               for i in range(width - 1)):
                        continue
                    joined = " ".join(cells)
                    if len(joined) > 400:
                        continue
                    hit = None
                    for match in self._match(joined, entries):
                        pieces = self._to_pieces(window, match[0], match[1],
                                                 match[3], match[2])
                        if pieces and len(pieces) >= 2:
                            hit = (pieces, match[3])
                            break
                    if hit is None:
                        if not self._one_cell(window):
                            continue
                        value = self._shape_entity(cells, joined)
                        if value is None:
                            continue
                        pieces = self._to_pieces(window, 0, len(joined),
                                                 value, COMPANY)
                        if not pieces or len(pieces) < 2:
                            # A name that fits in one cell is somebody else's job.
                            continue
                        hit = (pieces, value)
                    found.extend(hit[0])
        return found

    @staticmethod
    def _to_pieces(window, start: int, end: int, value: str, ptype: str):
        """Cut the joined match back into one candidate per cell."""
        pieces: List[Candidate] = []
        offset = 0
        for unit in window:
            length = len(unit.text)
            low = max(start, offset)
            high = min(end, offset + length)
            if high > low:
                pieces.append(Candidate(
                    unit.part, unit.index, low - offset, high - offset, ptype,
                    unit.text[low - offset:high - offset], "split-cells",
                    entity=value))
            offset += length + 1
        return pieces or None


def joined_value(text: str, start: int, end: int) -> str:
    """The part of ``text`` a joined match covered, for the mapping file."""
    return text[start:end]


class GazetteerCompanyDetector:
    """Curated organisations that carry no legal suffix (Trilegal, trusts)."""

    ptype = COMPANY
    name = "company/gazetteer"

    def __init__(self, entries: Sequence[str]):
        self.entries = sorted(set(entries), key=lambda e: (-len(e), e))
        self._patterns = [
            (entry, re.compile(r"(?<![A-Za-z])" + r"[ \u00a0]*".join(
                re.escape(tok) for tok in entry.split()) + r"(?![a-z])"))
            for entry in self.entries
        ]

    def find(self, ctx: ScanContext) -> List[Candidate]:
        found: List[Candidate] = []
        for unit in ctx.units:
            taken: List[Tuple[int, int]] = []
            for entry, pattern in self._patterns:
                for match in pattern.finditer(unit.text):
                    if any(match.start() < end and match.end() > start
                           for start, end in taken):
                        continue
                    value = match.group(0)
                    if _norm_key(value) in KEEP_ORGS:
                        continue
                    taken.append((match.start(), match.end()))
                    found.append(Candidate(unit.part, unit.index, match.start(),
                                           match.end(), COMPANY, value, "gazetteer"))
        return found


# --------------------------------------------------------------------------
# ADDRESSES
# --------------------------------------------------------------------------

#: Indian PIN: six digits with an optional internal space, tolerating the
#: letter-for-one typo present in the source ("Pune – 41l 005").
PIN = r"(?<![0-9A-Za-z])(?:\d{3}[ ]?\d{3}|\d{2}[ ]?[lI][ ]?\d{3})(?![0-9A-Za-z])"
PIN_RE = re.compile(PIN)
DASH_PIN_RE = re.compile(r"[ \u00a0]*[-\u2013\u2014][ \u00a0]*" + PIN)
TAIL_RE = re.compile(
    r"(?:\s*,?\s*\(?\s*(?:" + "|".join(ADDRESS_TAIL) + r")\s*\)?){0,2}\s*[;,]?\s*$",
    re.I,
)
#: The same state/country tail, but sitting mid paragraph because more of the
#: sentence follows: "… Pune – 410 501, Maharashtra, India and its Corporate
#: Office at …".  The tail is still the end of the address, so it is consumed;
#: a word character directly after it ("Indian", "Indo…") means it is not.
TAIL_MID_RE = re.compile(
    r"(?:\s*,?\s*\(?\s*(?:" + "|".join(ADDRESS_TAIL) + r")\s*\)?){1,2}\s*,?\s*"
    r"(?=[^\w]|$)",
    re.I,
)
#: Left boundary: a label, a sentence break, or a phrase that introduces an
#: address rather than being part of it.
LEFT_TRIGGER_RE = re.compile(
    r"(?i)\b(?:"
    r"registered\s+office|corporate\s+office|office|address|located\s+at|situated\s+at|"
    r"having\s+its|namely,?|namely|is|are|at|bounded\s+on"
    r")\b\s*[:.]?\s*"
)
#: Direction words sit *inside* an address far more often than in front of it
#: ("off 619, Bose Apartments", "opposite PYC basketball court").  They only
#: introduce one when nothing else of the address follows them.
WEAK_LEFT_TRIGGER_RE = re.compile(
    r"(?i)\b(?:off|near|next\s+to|opposite|behind|beside|above|below)\b\s*[:.]?\s*"
)
SENTENCE_BREAK_RE = re.compile(r"[.;:]\s")

#: Markers that can only appear in a postal address, as opposed to the weak
#: ones ("and", "at", "off", "no") which occur in ordinary prose.
STRONG_ADDRESS_MARKERS: Set[str] = {
    "road", "marg", "lane", "street", "nagar", "colony", "society", "soc",
    "apartment", "apartments", "bunglow", "bungalow", "villa", "layout",
    "avenue", "terrace", "tower", "building", "wing", "floor", "block",
    "plot", "village", "taluka", "district", "dist", "flat", "premises",
    "chambers", "estate", "heights", "residency", "mansion", "complex",
    "survey", "gat", "s.no", "square", "plaza", "gardens", "centre", "center",
    "farms", "farm", "township", "paradise", "enclave", "peth", "wadi",
}

#: A comma followed by one of these means the run is a clause, not an address.
CLAUSE_RE = re.compile(
    r"(?i),\s*(?:which|that|who|whom|whose|our|we|they|as\s+well|and\s+may|"
    r"and\s+the\s+company|result\s+of|in\s+case|if\s+the)"
)

#: Buildings and premises nouns: the first of them inside a run that starts
#: with a desk name is where the address really begins.
ADDRESS_BUILDINGS: Set[str] = {
    "bhavan", "house", "apartments", "apartment", "residency", "villa",
    "chambers", "complex", "plaza", "centre", "center", "mansion", "estate",
    "heights", "park", "campus", "tower", "building", "wings", "premises",
    "annex", "bldg", "bldgs", "square", "place", "court", "lodge", "hall",
}

#: "Chakan Unit No. 2 (Birdewadi)" is an address line without any of the
#: marker words: the unit and the village are named, nothing else.
UNIT_LINE_RE = re.compile(
    r"(?i)\b(?:unit|village|premises|shed|works|godown)\s*"
    r"(?:no\.?|number|#)?\s*[\w./-]*"
)

#: A departmental or desk name, not an address: "Corporation Finance
#: Department Division of Issues and Listing SEBI Bhavan, Plot No. C4 A".
ROLE_NOUN_RE = re.compile(
    r"(?i)\b(?:department|division|directorate|office|centre|center|section|"
    r"desk|team|cell)\b"
)


def _trim_role_prefix(text: str, start: int, end: int) -> Tuple[int, int]:
    """Keep a desk name out of an address span.

    A departmental label in front of a regulator's address is not an
    address, and swallowing it would delete the name of the department from
    the document.  Only a span that starts at the very beginning of a
    paragraph can be trimmed, so an address that a trigger has already
    delimited is never touched.
    """
    if start > 0:
        return start, end
    head = text[start:end]
    role = ROLE_NOUN_RE.search(head)
    if not role or role.end() > 60:
        return start, end
    for match in re.finditer(r"[A-Za-z][A-Za-z'.]*|\d+", head[role.end():]):
        word = match.group(0).lower()
        if (word in ADDRESS_BUILDINGS or word in STRONG_ADDRESS_MARKERS
                or word in ADDRESS_CITIES or match.group(0)[0].isdigit()):
            offset = role.end() + match.start()
            if offset <= 0:
                return start, end
            return start + offset, end
    return start, end


def _snap_to_token(text: str, start: int, end: int) -> Tuple[int, int]:
    """Move a span onto token boundaries: a boundary never cuts a word."""
    if start > 0 and (text[start - 1].isalnum() or text[start - 1] in "'-.&/"):
        cursor = start
        while cursor < len(text) and not text[cursor].isspace():
            cursor += 1
        while cursor < len(text) and text[cursor].isspace():
            cursor += 1
        start = cursor
    if (end < len(text) and end > 0 and text[end].isalnum()
            and (text[end - 1].isalnum() or text[end - 1] in "'-&/")):
        back = end
        while back > 0 and not text[back - 1].isspace():
            back -= 1
        end = back
    return _clean_span(text, start, end)


#: Tokens whose trailing period does not end a sentence.
ABBREVIATION_RE = re.compile(
    r"(?i)\b(?:no|nos|s|sr|ss|m/s|mr|mrs|ms|dr|shri|smt|st|rd|ln|floor|fl|opp|"
    r"near|sec|block|bl|dist|pin|c/o|care|of|co|pvt|ltd|inc|llp|extn|approx|"
    r"village|vill|taluka|dist)\.$"
)


def _truncate_at_sentence(text: str, start: int, end: int) -> Tuple[int, int]:
    """An address never spans two sentences."""
    for match in SENTENCE_BREAK_RE.finditer(text, start, end):
        if ABBREVIATION_RE.search(text[max(start, match.start() - 12):match.start() + 1]):
            continue
        return start, match.start() + 1
    return start, end



class AddressDetector:
    """Postal address detection.

    Primary rule: a PIN anchor, expanded left to the nearest safe boundary and
    right over the state/country tail.  Secondary rule: a loose address with no
    PIN, which needs at least two structural markers and one locality.
    Continuation paragraphs (an address broken over two or three paragraphs, as
    in the General Information section) are linked into one group.
    """

    ptype = ADDRESS
    name = "address"

    def find(self, ctx: ScanContext, hints: Sequence[Candidate] = ()) -> List[Candidate]:
        found: List[Candidate] = []
        for part in ctx.parts():
            units = ctx.units_of(part)
            boundaries: Dict[int, List[Tuple[int, int]]] = {}
            for hint in hints:
                if hint.part == part:
                    boundaries.setdefault(hint.paragraph, []).append(
                        (hint.start, hint.end)
                    )
            per_unit: List[List[Candidate]] = [[] for _ in units]
            for position, unit in enumerate(units):
                per_unit[position] = self._in_paragraph(
                    unit, boundaries.get(unit.index, [])
                )
            found.extend(self._link_continuations(units, per_unit))
        return found

    # -- within one paragraph -------------------------------------------
    def _in_paragraph(self, unit: Unit, boundaries: Sequence[Tuple[int, int]]) -> List[Candidate]:
        text = unit.text
        if len(text) > 600 or ADDRESS_NEGATIVE.search(text):
            if not PIN_RE.search(text):
                return []
        out: List[Candidate] = []
        occupied: List[Tuple[int, int]] = []
        for match in list(DASH_PIN_RE.finditer(text)) + list(PIN_RE.finditer(text)):
            start, end = match.span()
            if any(start < b_end and end > b_start for b_start, b_end in occupied):
                continue
            if not self._pin_is_address(text, match):
                continue
            left = self._left_boundary(text, match.start(), boundaries)
            right = self._right_boundary(text, match.end())
            if right <= left:
                continue
            left, right = _truncate_at_sentence(text, left, right)
            left, right = _snap_to_token(text, left, right)
            if right <= left:
                continue
            out.append(Candidate(unit.part, unit.index, left, right, ADDRESS,
                                 text[left:right], "pin", group="addr"))
            occupied.append((left, right))
        loose = list(self._loose_spans(unit, boundaries))
        loose += [(c.start, c.end)
                  for c in self._loose_in_sentences(unit, boundaries)]
        for start, end in loose:
            if any(start < b_end and end > b_start for b_start, b_end in occupied):
                continue
            if any(start < c.end and end > c.start for c in out):
                continue
            out.append(Candidate(unit.part, unit.index, start, end, ADDRESS,
                                 text[start:end], "loose", group="addr"))
        return out

    @staticmethod
    def _pin_is_address(text: str, match: "re.Match") -> bool:
        """A 6 digit run is a PIN only in an address context.

        Three contexts count: a dash before it ("801 - 804"), a
        state/country after it ("… 400 051, Maharashtra, India"), or a
        locality right before it ("Bandra Kurla Complex, Bandra East Mumbai
        400 051" - the tail of a wrapped address block).
        """
        before = text[max(0, match.start() - 3): match.start()]
        if re.search(r"[-\u2013\u2014][ \u00a0]*$", before) and not re.search(
            r"[A-Za-z0-9][-\u2013\u2014][ \u00a0]*$", before
        ):
            return True
        after = text[match.end(): match.end() + 30]
        if re.match(r"\s*,?\s*\(?\s*(?:" + "|".join(ADDRESS_TAIL) + r")\b", after, re.I):
            return True
        leading = text[max(0, match.start() - 40): match.start()].lower()
        for token in re.findall(r"[a-z]+", leading):
            if token in ADDRESS_CITIES or token in ADDRESS_TAIL:
                return True
        return False

    def _left_boundary(self, text: str, anchor_start: int,
                       boundaries: Sequence[Tuple[int, int]]) -> int:
        limit = 0
        for start, end in boundaries:
            if end <= anchor_start:
                limit = max(limit, end)
        for match in LEFT_TRIGGER_RE.finditer(text, 0, anchor_start):
            limit = max(limit, match.end())
        for match in WEAK_LEFT_TRIGGER_RE.finditer(text, 0, anchor_start):
            if "," not in text[match.end():anchor_start]:
                limit = max(limit, match.end())
        for match in SENTENCE_BREAK_RE.finditer(text, 0, anchor_start):
            # "S. no. 245/ 104, …" - the periods in the address' own label
            # are not sentence breaks.
            if ABBREVIATION_RE.search(text[max(0, match.start() - 12):match.start() + 1]):
                continue
            limit = max(limit, match.end())
        for match in re.finditer(r"(?i)\b(?:namely|being)\b", text[:anchor_start]):
            limit = max(limit, match.end())
        if limit >= anchor_start:
            return anchor_start
        start = limit
        while start < anchor_start and text[start] in " \u00a0,;:-":
            start += 1
        return start

    @staticmethod
    def _right_boundary(text: str, anchor_end: int) -> int:
        tail = TAIL_RE.match(text, anchor_end)
        if tail:
            return tail.end()
        # "… Mumbai 400083, (Maharashtra), India Telephone: +91 81081 14949"
        # and "… Pune – 410 501, Maharashtra, India and its Corporate Office
        # at …": the tail is the end of the address even though the sentence
        # continues.  Take the tail, and nothing beyond it.
        mid = TAIL_MID_RE.match(text, anchor_end)
        if mid:
            return mid.end()
        return anchor_end

    def _loose_spans(self, unit: Unit,
                      boundaries: Sequence[Tuple[int, int]] = ()) -> List[Tuple[int, int]]:
        """Address-like run with no PIN (e.g. "Unit no. 1601, B- wing BKC, ...").

        Requires a marker that cannot occur in prose plus a locality, and
        rejects a run that turns out to be a clause of a sentence.
        """
        text = unit.text
        if len(text) > 400 or PIN_RE.search(text):
            return []
        lowered = text.lower()
        if any(phrase in lowered for phrase in ADDRESS_STOP_PHRASES):
            return []
        markers = 0
        strong = False
        locality = False
        for token in re.split(r"(\s+)", text):
            bare = token.strip(" ,.;:()").lower()
            if not bare:
                continue
            if bare in STRONG_ADDRESS_MARKERS:
                strong = True
            if bare in ADDRESS_MARKERS:
                markers += 1
            elif bare in ADDRESS_CITIES or bare in ADDRESS_TAIL:
                locality = True
        if UNIT_LINE_RE.search(text):
            # "Chakan Unit No. 2 (Birdewadi)" names the place without a
            # single marker word.
            strong = True
            markers = max(markers, 2)
        if markers < 2 or not locality or not strong:
            return []
        start = 0
        limit = LEFT_TRIGGER_RE.search(text)
        if limit:
            start = limit.end()
        weak = WEAK_LEFT_TRIGGER_RE.search(text)
        if weak and "," not in text[weak.end():]:
            start = max(start, weak.end())
        # Never start inside a company name or a person name that is already
        # being redacted: "IndusInd Bank Limited 2401 Gen Thimmayya Road".
        for hint_start, hint_end in boundaries:
            if hint_end <= len(text) and hint_end > start and \
                    not text[hint_start:hint_end].strip().startswith(tuple("0123456789")):
                if hint_end < len(text) and text[hint_end] not in " ,.;:-–—\u00a0":
                    continue
                start = max(start, hint_end)
        end = len(text)
        stop = re.search(r"(?i)\b(?:telephone|tel|mobile|email|e-mail|website|"
                         r"contact\s*person|fax)\b", text)
        if stop:
            end = stop.start()
        start, end = _truncate_at_sentence(text, start, end)
        start, end = _snap_to_token(text, start, end)
        if end <= start or CLAUSE_RE.search(text[start:end]):
            return []
        if BARE_REGION_RE.match(text[start:end].strip()) and \
                text[start:end].strip() != text.strip():
            return []
        return [(start, end)]

    def _loose_in_sentences(self, unit: Unit,
                            boundaries: Sequence[Tuple[int, int]]) -> List[Candidate]:
        """Retry the loose rule on each sentence of a prose paragraph.

        "His contact details are as set forth below: Gat No. 11/3, 11/4, 11/5,
        Village Birdewadi" is a sentence of prose with an address after the
        colon.  The address rule refuses to span a sentence break, so the
        second sentence is offered on its own.
        """
        text = unit.text
        bounds: List[Tuple[int, int]] = []
        start = 0
        for match in SENTENCE_BREAK_RE.finditer(text):
            end = match.end()
            if ABBREVIATION_RE.search(text[max(0, match.start() - 12):match.start() + 1]):
                continue
            if text[start:end].strip():
                bounds.append((start, end))
            start = end
        if text[start:].strip():
            bounds.append((start, len(text)))
        if len(bounds) < 2:
            return []
        out: List[Candidate] = []
        for left, right in bounds[1:]:
            piece = Unit(unit.part, unit.index, text[left:right], unit.cell)
            shift = [(max(0, s - left), max(0, e - left))
                     for s, e in boundaries if left <= s < right or left < e <= right]
            for start_in, end_in in self._loose_spans(piece, shift):
                out.append(Candidate(
                    unit.part, unit.index, left + start_in, left + end_in,
                    ADDRESS, text[left + start_in:left + end_in], "loose",
                    group="addr"))
        return out

    # -- across paragraphs ----------------------------------------------
    def _link_continuations(self, units: Sequence[Unit],
                            per_unit: Sequence[List[Candidate]]) -> List[Candidate]:
        """Extend an address leftwards/rightwards over its own continuation lines.

        An address in this document is frequently split as
        "801-804, Wing A, Building No 3 Inspire BKC, G Block" /
        "Bandra Kurla Complex, Bandra East Mumbai 400 051" / "Maharashtra, India".
        """
        out: List[Candidate] = []
        for position, candidates in enumerate(per_unit):
            out.extend(candidates)
            if not candidates or candidates[0].source != "pin":
                continue
            start_para = position
            while start_para - 1 >= 0:
                previous = units[start_para - 1]
                include = self._continuation_span(previous)
                if include is None:
                    break
                start, end = include
                out.append(Candidate(previous.part, previous.index, start, end,
                                     ADDRESS, previous.text[start:end],
                                     "continuation", group="addr"))
                start_para -= 1
            end_para = position
            while end_para + 1 < len(units):
                nxt = units[end_para + 1]
                if self._is_address_tail_paragraph(nxt):
                    out.append(Candidate(nxt.part, nxt.index, 0, len(nxt.text),
                                         ADDRESS, nxt.text, "continuation",
                                         group="addr"))
                    end_para += 1
                else:
                    break
        return out

    @staticmethod
    def _continuation_span(unit: Unit) -> Optional[Tuple[int, int]]:
        text = unit.text
        if not text.strip() or len(text) > 300:
            return None
        if not text[0].isupper() and not text[0].isdigit() and text[0] not in "([\"'-":
            return None  # an address line never starts mid-sentence
        lowered = text.lower()
        if any(phrase in lowered for phrase in ADDRESS_STOP_PHRASES):
            return None
        strong = False
        for token in re.split(r"[\s,;]+", text):
            if token.strip("().:").lower() in STRONG_ADDRESS_MARKERS:
                strong = True
                break
        if not strong:
            if not UNIT_LINE_RE.search(text):
                # "Unit 1", "S." and ordinary prose all contain weak markers
                # such as "a", "s" or "and"; only a real address word counts.
                return None
        start = 0
        limit = LEFT_TRIGGER_RE.search(text)
        if limit:
            start = limit.end()
        weak = WEAK_LEFT_TRIGGER_RE.search(text)
        if weak and "," not in text[weak.end():]:
            start = max(start, weak.end())
        start, end = _truncate_at_sentence(text, start, len(text))
        start, end = _snap_to_token(text, start, end)
        if end <= start or CLAUSE_RE.search(text[start:end]):
            return None
        start, end = _trim_role_prefix(text, start, end)
        if end <= start:
            return None
        return (start, end)

    @staticmethod
    def _is_address_tail_paragraph(unit: Unit) -> bool:
        text = unit.text.strip()
        if not text or len(text) > 80:
            return False
        if re.fullmatch(r"(?i)" + PIN, text):
            return True
        lowered = text.lower()
        return bool(re.fullmatch(r"(?:,?\s*\(?(?:" + "|".join(ADDRESS_TAIL) +
                                 r")\)?,?\s*)+", lowered))


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------

class RedactionConfig:
    def __init__(self, use_gazetteer: bool = True,
                 person_gazetteer: Sequence[str] = (),
                 company_gazetteer: Sequence[str] = ()):
        self.use_gazetteer = use_gazetteer
        self.person_gazetteer = list(person_gazetteer)
        self.company_gazetteer = list(company_gazetteer)


def default_detectors(config: RedactionConfig) -> List:
    detectors = [
        EmailDetector(),
        UrlDetector(),
        PhoneDetector(),
        SsnDetector(),
        CreditCardDetector(),
        IpAddressDetector(),
        DateOfBirthDetector(),
        DinDetector(),
        LabelNameDetector(),
        ContactPersonNameDetector(),
        RoleAnchorNameDetector(),
        DesignationTableNameDetector(),
        AllCapsListNameDetector(),
        CompanyDetector(
            _entity_boundary_re(config) if config.use_gazetteer else None
        ),
    ]
    detectors.append(SplitEntityDetector(config))
    if config.use_gazetteer:
        detectors.append(GazetteerNameDetector(config.person_gazetteer))
        detectors.append(GazetteerCompanyDetector(config.company_gazetteer))
    return detectors


def _entity_boundary_re(config: "RedactionConfig"):
    """A matcher for known entity names, used as a name-growth boundary.

    Growing a legal-suffix name leftwards through a *different* known entity
    would swallow it: in the promoter list, "… Kanchenjunga Family Trust and
    Waterloo Industrial Park VI Private Limited", the leftwards walk must stop
    at the trust instead of turning the whole sentence into one company.
    """
    names = sorted(set(config.person_gazetteer) | set(config.company_gazetteer),
                   key=lambda e: (-len(e), e))
    if not names:
        return None
    body = "|".join(
        r"[ \u00a0]*".join(re.escape(tok) for tok in name.split())
        for name in names
    )
    return re.compile(r"(?i)(?<![A-Za-z])(?:" + body + r")(?![A-Za-z])")


def resolve(candidates: Iterable[Candidate]) -> List[Candidate]:
    """Keep the highest priority type on any overlap; drop exact duplicates."""
    ordered = sorted(
        candidates,
        # Within a type the longer span wins, so "Solar Energy Corporation of
        # India Limited" beats the "Solar Energy Corporation" that its own
        # suffix produced.
        key=lambda c: (-PRIORITY.get(c.ptype, 0), -len(c.text), c.start),
    )
    accepted: List[Candidate] = []
    for candidate in ordered:
        clash = False
        duplicate = False
        for other in accepted:
            if (other.part == candidate.part
                    and other.paragraph == candidate.paragraph
                    and candidate.start < other.end
                    and candidate.end > other.start):
                if (other.start, other.end) == (candidate.start, candidate.end):
                    duplicate = True
                clash = True
                break
        if clash or duplicate:
            continue
        accepted.append(candidate)
    accepted.sort(key=lambda c: (c.part, c.paragraph, c.start))
    return accepted


def scan(ctx: ScanContext, config: RedactionConfig) -> List[Candidate]:
    candidates: List[Candidate] = []
    for detector in default_detectors(config):
        candidates.extend(detector.find(ctx))
    settled = resolve(candidates)
    address_detector = AddressDetector()
    candidates.extend(address_detector.find(ctx, hints=settled))
    return resolve(candidates)
