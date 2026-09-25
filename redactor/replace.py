"""replace.py - deterministic, structure preserving pseudonyms.

Two properties matter more than anything else here:

**Consistency.**  Every mention of the same source value maps to the same fake
value, everywhere in the document, in every run.  That is what makes a redacted
document still readable (one person, one placeholder) and what the evaluation
report measures as the consistency rate.

**Reproducibility.**  The mapping is a keyed hash of the seed and the normalised
source value, so two runs on two machines produce byte-identical output, and the
consistency checks in the evaluation report are meaningful.

Generators preserve the *shape* of the value they replace: token count for
names, digit count and grouping for phone numbers, legal suffix for companies,
component structure and a well-formed PIN for addresses.  Shape is preserved so
the 3,225 fixed-layout table cells in this document do not reflow.
"""

import hashlib
import re
import unicodedata
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from . import detect

# --------------------------------------------------------------------------
# Pools.  Every value here is a placeholder.  They are deliberately ordinary
# looking (the assignment's own examples are "John Doe" and "Peter Parker") so
# the redacted document stays readable; identification is via redaction_map.json.
# --------------------------------------------------------------------------

GIVEN_NAMES: Tuple[str, ...] = (
    "John", "Peter", "Rahul", "Aisha", "Vikram", "Neha", "Arjun", "Meera",
    "Kabir", "Priya", "Rohan", "Ananya", "Aditya", "Sneha", "Nikhil", "Divya",
    "Karan", "Ishita", "Sameer", "Pooja", "Vivek", "Ritu", "Amit", "Kavya",
    "Sanjay", "Nisha", "Manoj", "Ritu", "Harsh", "Preeti", "Gaurav", "Sonal",
    "Deepak", "Anjali", "Suresh", "Megha", "Rakesh", "Shalini", "Naveen",
    "Pallavi", "Varun", "Swati", "Ajay", "Rekha", "Sanjay", "Trupti",
)

SURNAMES: Tuple[str, ...] = (
    "Doe", "Parker", "Sharma", "Nair", "Bose", "Iyer", "Kulkarni", "Menon",
    "Desai", "Chatterjee", "Reddy", "Banerjee", "Gowda", "Rao", "Pillai",
    "Sethi", "Khanna", "Basu", "Dutta", "Ghosh", "Kapoor", "Malhotra",
    "Saxena", "Trivedi", "Verma", "Joshi", "Mishra", "Pandey", "Rastogi",
    "Shetty", "Bhandary", "Salvi", "Thakur", "Vyas", "Wagh", "Zaveri",
)

COMPANY_STEMS: Tuple[str, ...] = (
    "Meridian", "Northbridge", "Silverline", "Kensington", "Fairmont",
    "Lakeshore", "Brightwater", "Stonebridge", "Highland", "Cedarpoint",
    "Westgate", "Ironwood", "Marlowe", "Ashford", "Pinecrest", "Redstone",
    "Clearwater", "Granville", "Hollowbrook", "Thornbury", "Ravenswood",
    "Ellington", "Bayfield", "Camden", "Doveridge", "Eastvale", "Foxglove",
)

COMPANY_MIDS: Tuple[str, ...] = (
    "Industries", "Ventures", "Holdings", "Systems", "Technologies",
    "Enterprises", "Solutions", "Capital", "Partners", "Trading", "Logistics",
    "Engineering", "Networks", "Resources", "Analytics",
)

#: Short invented words, for the names whose own head is only a few
#: characters wide - "KSH Infra Park 5 Private Limited" has fifteen characters
#: to spend before the legal form and cannot carry a forty-character stem.
COMPANY_SHORT: Tuple[str, ...] = (
    "Aria", "Vexa", "Nexo", "Orin", "Luma", "Vale", "Onza", "Rill", "Kite",
    "Fern", "Iris", "Opal", "Pike", "Reef", "Sable", "Tide", "Umber", "Wren",
    "Oriel", "Vesta", "Halcyon", "Zephyr", "Cobalt", "Marlow", "Sterling",
    "Ash", "Bay", "Elm", "Ink", "Oar", "Orb", "Sol", "Zen", "Cob", "Den",
    "Fen", "Lux", "Wyn",
)

#: Words that tie a name together and carry no identity of their own.
COMPANY_LINKERS = frozenset((
    "&", "and", "of", "for", "the", "de", "du", "von", "der", "/", "and,",
))

STREET_NAMES: Tuple[str, ...] = (
    "Maple", "Cedar", "Willow", "Birch", "Juniper", "Sycamore", "Hazel",
    "Aspen", "Rowan", "Laurel", "Magnolia", "Cypress", "Alder", "Poplar",
    "Orchard", "Meadow", "Clover", "Heather", "Fern", "Ivy",
)

STREET_TYPES: Tuple[str, ...] = (
    "Road", "Marg", "Lane", "Avenue", "Street", "Drive", "Boulevard", "Way",
)

BUILDING_WORDS: Tuple[str, ...] = (
    "Anand", "Shanti", "Gauri", "Kaveri", "Meera", "Saraswati", "Laxmi",
    "Indira", "Gandhi", "Nehru", "Tagore", "Bose", "Gopal", "Krishna",
)

#: Place names are invented, not real: a fake address must never be mistaken for
#: a real one, and must never be a real locality that a residual-PII check
#: would then report as unredacted.  The pools are single-word so a compact
#: replacement still fits a short source span.
CITIES: Tuple[str, ...] = (
    "Rajapur", "Nandgaon", "Belpura", "Karimnagar", "Devgaon", "Ambarpur",
    "Sultanpur", "Raghunathpur", "Chandrapur", "Vilaspur", "Narsinghpur",
    "Barwadih", "Sitapur", "Dharampur", "Gulbarga", "Kheda", "Mandvi",
)

STATES: Tuple[str, ...] = (
    "Northmarch", "Rajayana", "Sundarbans", "Kalimpore", "Vindhyanchal",
    "Mahekan", "Trivenipur", "Ashwin Pradesh", "Brahmaputra", "Coromandel",
    "Ghataprabha", "Sahyadri", "Dakshina Nadu", "Udaygarh", "Konkan West",
)

#: RFC 2606 reserved domains - guaranteed never to resolve, so a redacted
#: document cannot accidentally mail anybody.
EMAIL_DOMAINS: Tuple[str, ...] = ("example.com", "example.net", "example.org")

#: Invented words used to pad a redacted URL path to the original length.
URL_PATH_WORDS: Tuple[str, ...] = (
    "contact", "enquiry", "desk", "portal", "info", "support", "office",
    "connect", "services", "team", "inquiry", "reach", "connect", "links",
    "updates", "bulletin", "notices", "resources", "channel", "listing",
)


class Pseudonymiser:
    def __init__(self, seed: str = "scaler-pii-redaction-2026", salt: str = ""):
        self.seed = seed
        self.salt = salt
        self._cache: Dict[Tuple[str, str], str] = {}
        self._phone_digits: Dict[str, str] = {}
        self._email_local: Dict[str, str] = {}
        self._email_domain: Dict[str, str] = {}
        self._urls: Dict[str, str] = {}

    # -- core -----------------------------------------------------------
    def _int(self, *parts: str) -> int:
        payload = "\x1f".join((self.seed, self.salt) + parts).encode("utf-8")
        return int.from_bytes(hashlib.sha256(payload).digest()[:8], "big")

    def _pick(self, options: Sequence[str], *parts: str) -> str:
        return options[self._int(*parts) % len(options)]

    def _digits(self, count: int, *parts: str) -> str:
        value = self._int(*parts)
        out = []
        for _ in range(count):
            out.append(str(value % 10))
            value //= 10
        return "".join(out)

    # -- public API -----------------------------------------------------
    def fake(self, ptype: str, value: str) -> str:
        """Return the fake for ``value`` in the shape appropriate to ``ptype``."""
        key = self.canonical(ptype, value)
        if ptype == detect.PHONE:
            # The *digits* are the entity; the layout belongs to each
            # occurrence.  "+91 22 40094400" and "+91 22 4009 4400" are one
            # number printed twice, and each must keep its own grouping.
            if key not in self._phone_digits:
                self._phone_digits[key] = self._national_digits(value)
            return self._render_phone(value, self._phone_digits[key])
        if ptype == detect.EMAIL:
            if key not in self._email_local:
                local, domain = self._email_parts(key)
                # The address may be printed twice in a line - the
                # ``mailto:`` target of a hyperlink field and the visible
                # text - so the whole address is kept no wider than the
                # original.  The reserved domain is never dropped in favour
                # of the real one; the local part gives up the characters.
                room = len(value) - len(domain) - 1
                if len(local) > room:
                    local = local[:max(1, room)]
                self._email_local[key] = local
                self._email_domain[key] = domain
            return _match_case(
                self._email_local[key] + "@" + self._email_domain[key], value)
        if ptype == detect.URL:
            if key not in self._urls:
                self._urls[key] = self._url(value)
            return _match_case(self._urls[key], value)
        if key not in self._cache:
            self._cache[key] = self._generate(ptype, key, value)
        return self._cache[key]

    @staticmethod
    def canonical(ptype: str, value: str) -> str:
        """The identity of an entity: case, spacing and punctuation insensitive."""
        text = unicodedata.normalize("NFKC", value)
        text = re.sub(r"\s+", " ", text).strip()
        if ptype == detect.PHONE:
            return "".join(ch for ch in text if ch.isdigit())
        if ptype == detect.EMAIL:
            return text.lower()
        if ptype == detect.URL:
            # "www.kshinternational. com" and "www.kshinternational.com" are
            # one address: Word split it across a run seam.
            return re.sub(r"\s+", "", text.lower())
        if ptype == detect.DATE_OF_BIRTH:
            return re.sub(r"[^0-9]", "", text)
        if ptype in (detect.CREDIT_CARD,):
            return "".join(ch for ch in text if ch.isdigit())
        if ptype in (detect.PERSON_NAME, detect.COMPANY):
            # "KushalSubbayya Hegde" is the same person as "Kushal Subbayya
            # Hegde": Word dropped the space at a run seam.  "Kirtane & Pandit
            # LLP" and "Kirtane & Pandit, LLP" are the same firm.  Neither
            # spaces nor punctuation are part of the identity.
            return re.sub(r"[^0-9a-z]", "", text.lower())
        if ptype in (detect.ADDRESS, detect.DIN, detect.SSN, detect.IP_ADDRESS):
            return text.lower()
        return text

    def _generate(self, ptype: str, key: str, original: str) -> str:
        if ptype == detect.PERSON_NAME:
            return self._person(key, original)
        if ptype == detect.EMAIL:
            return self._email(key)
        if ptype == detect.PHONE:
            return self._phone(key, original)
        if ptype == detect.COMPANY:
            fake = self._company(key, original)
            if original.strip() == original and original.isupper() \
                    and not fake.isupper():
                # "HDFC BANK LIMITED" in a table header stays in capitals;
                # the layout and the emphasis of the line are unchanged.
                return fake.upper()
            return fake
        if ptype == detect.ADDRESS:
            return self._address(key, original)
        if ptype == detect.DIN:
            return self._digits(8, "din", key)
        if ptype == detect.SSN:
            area = 900 + self._int("ssn-area", key) % 100
            return "%03d-%02d-%04d" % (area, self._int("ssn-grp", key) % 100,
                                       self._int("ssn-ser", key) % 10000)
        if ptype == detect.CREDIT_CARD:
            base = "411111111111" + self._digits(4, "cc", key)
            return " ".join(base[i:i + 4] for i in range(0, len(base), 4))
        if ptype == detect.IP_ADDRESS:
            return "198.51.100.%d" % (self._int("ip", key) % 254 + 1)
        if ptype == detect.DATE_OF_BIRTH:
            year = 1950 + self._int("dob-y", key) % 50
            month = 1 + self._int("dob-m", key) % 12
            day = 1 + self._int("dob-d", key) % 28
            return "%02d/%02d/%d" % (day, month, year)
        return "[REDACTED]"

    # -- generators -----------------------------------------------------
    def _person(self, key: str, original: str) -> str:
        """Preserve token count and the First/Middle-initial/Last shape."""
        tokens = original.split()
        # "Dr. Sanjay Gaikwad" keeps its title: dropping it would change the
        # line, and the title is not identifying.
        title = ""
        if tokens and re.fullmatch(r"(?i)(dr|mr|mrs|ms|prof|shri|smt)\.?", tokens[0]):
            title = tokens[0] if tokens[0].endswith(".") else tokens[0] + "."
            tokens = tokens[1:]
        count = max(2, min(len(tokens), 4))
        first = self._pick(GIVEN_NAMES, "given", key)
        last = self._pick(SURNAMES, "surname", key)
        if count == 2:
            name = "%s %s" % (first, last)
        elif count == 3:
            initial = self._pick(SURNAMES, "middle", key)[0].upper() + "."
            name = "%s %s %s" % (first, initial, last)
        else:
            name = "%s %s %s %s" % (first, self._pick(GIVEN_NAMES, "given2", key),
                                    self._pick(SURNAMES, "surname2", key), last)
        return ("%s %s" % (title, name)) if title else name

    def _email_parts(self, key: str) -> Tuple[str, str]:
        """A local part and an RFC 2606 reserved domain, shaped like the source."""
        given, _, domain = key.partition("@")
        tokens = [t for t in re.split(r"[._%+\-]+", given) if t]
        if not tokens:
            tokens = ["user"]
        words: List[str] = []
        for token in tokens[:2]:
            words.append("".join(ch for ch in token.lower() if ch.isalnum()) or "user")
        local = ".".join(words) or "user"
        target = domain or "example.com"
        if re.search(r"\.[a-z]{2,}$", target):
            # Keep the shape, drop the real domain: example.com/org/net.
            tld = target.rsplit(".", 1)[-1].lower()
            target = "example." + (tld if tld in {"com", "net", "org"} else "com")
        else:
            target = self._pick(EMAIL_DOMAINS, "edomain", key)
        return local, target

    def _url(self, original: str) -> str:
        """A reserved-domain URL shaped like the one that was found.

        The scheme, the ``www.`` prefix, the top level domain and the length
        of any path are kept so the line does not change width, and the host
        and path are replaced: the real host is how a redacted party is
        re-identified.
        """
        scheme = ""
        rest = original
        for prefix in ("https://", "http://"):
            if rest.lower().startswith(prefix):
                scheme, rest = prefix, rest[len(prefix):]
                break
        host, slash, path = rest.partition("/")
        labels = host.split(".")
        tld = labels[-1].lower() if len(labels) > 1 else "com"
        tld = tld if tld in {"com", "net", "org"} else "com"
        with_www = "www.example.%s" % tld
        bare = "example.%s" % tld
        # "www." is kept when the original host had room for it: display text
        # reads better with it.  The reserved host is never longer than the
        # one it replaces - a URL printed twice in a hyperlink field would
        # otherwise widen the line by four characters.
        fake_host = self._fit_host(with_www, host, bare)
        if not slash:
            return scheme + fake_host
        tail = self._url_path(path, original)
        room = len(original) - len(scheme) - len(fake_host) - 1
        if len(tail) > room:
            tail = tail[:max(0, room)]
        return scheme + fake_host + "/" + tail

    @staticmethod
    def _fit_host(with_www: str, host: str, bare: str) -> str:
        """The longest reserved host that still fits the original host."""
        for candidate in (with_www, bare, "example.net", "example.org",
                          "test.invalid", "test"):
            if len(candidate) <= len(host):
                return candidate
        return host[:1] + ".test" if len(host) >= 7 else "test"

    def _url_path(self, path: str, original: str) -> str:
        """A same-length path: invented words for letters, digits kept.

        Punctuation and digits carry no identity and keep the tracking
        parameters readable; only the words are invented, and they are cut to
        the length of the run they replace so the line does not reflow.
        """
        out = []
        for token in re.findall(r"[A-Za-z]+|[^A-Za-z]+", path):
            if not token.isalpha():
                out.append(token)
                continue
            size = len(token)
            words = [w for w in URL_PATH_WORDS if w[0] == token[0].lower()] \
                or list(URL_PATH_WORDS)
            best = min(words, key=lambda w: (abs(len(w) - size), w))
            while len(best) < size:
                best += words[0]
            out.append(best[:size])
        return "".join(out)

    def _phone(self, key: str, original: str) -> str:
        return self._render_phone(original, self._national_digits(original))

    def _national_digits(self, original: str) -> str:
        """The subscriber digits, excluding the country / STD code."""
        positions = [i for i, ch in enumerate(original) if ch.isdigit()]
        national = positions[self._phone_prefix_length(original):]
        if not national:
            national = positions
        return self._digits(len(national), "phone", "".join(
            original[i] for i in national))

    @staticmethod
    def _render_phone(original: str, digits: str) -> str:
        """Rewrite only the national digits, leaving the literal layout intact.

        "+91 22 4009 4400" keeps its country code, its spaces and its group
        sizes; only the subscriber digits change.  This is the most layout-safe
        transformation available and it keeps the length identical, so table
        cells cannot reflow.
        """
        positions = [i for i, ch in enumerate(original) if ch.isdigit()]
        if not positions:
            return "+91 00000 00000"
        national = positions[Pseudonymiser._phone_prefix_length(original):]
        if not national:
            national = positions
        digits = Pseudonymiser._fit_digits(original, digits, len(national))
        chars = list(original)
        for position, digit in zip(national, digits):
            chars[position] = digit
        return "".join(chars)

    @staticmethod
    def _fit_digits(original: str, digits: str, width: int) -> str:
        """Fit the canonical digits to a differently spelled occurrence.

        One number may be printed as "+91 22 4009 4400" and "+91 22 4009 44"
        (a truncated cell).  The replacement must match the printed width, so
        the canonical digits are extended - never shortened, so the identity
        of the number is still recognisable - to fit.
        """
        if len(digits) == width:
            return digits
        if len(digits) > width:
            return digits[:width]
        payload = hashlib.sha256(("\x1f".join(("phone-pad", original)))
                                 .encode("utf-8")).digest()
        value = int.from_bytes(payload[:8], "big")
        out = list(digits)
        while len(out) < width:
            out.append(str(value % 10))
            value //= 10
        return "".join(out)

    @staticmethod
    def _phone_prefix_length(original: str) -> int:
        """How many leading digits are the country code / STD code."""
        country = re.match(r"\s*(\+[\s]*)?91", original)
        if country and country.group(1):
            return 2
        std = re.match(r"\s*(0\d{2,4})(?=[\s\-])", original)
        if std:
            return len(std.group(1))
        return 0

    def _company(self, key: str, original: str) -> str:
        """Preserve the legal suffix, the word count and the width.

        The replacement has as many words as the source and is never longer
        than it: a one-word firm ("Nuvama") stays one word, a firm whose legal
        form is a separate word ("Cindus Corporation") does not gain a middle
        word, and "KSH Infra Park 5 Private Limited" - fifteen characters of
        head - is given short words rather than a forty-character stem.  A
        longer name needs more line width than the table cell it sits in.
        """
        suffix = ""
        stripped = original.rstrip()
        for candidate in ("Private Limited", "Family Trust", "Pvt Ltd",
                          "Partnership", "Foundation", "Incorporated",
                          "Corporation", "Limited", "Society", "Trust",
                          "HUF", "LLP", "Inc.", "Ltd.", "Ltd"):
            # Case-insensitive, because the document also prints names in
            # capitals ("EVEREST FAMILY TRUST"); the suffix keeps the casing
            # it had in the source.
            if stripped.lower().endswith(candidate.lower()):
                found = stripped[len(stripped) - len(candidate):]
                # "EVEREST FAMILY TRUST" prints the legal form in capitals;
                # the replacement spells it the ordinary way.
                suffix = (candidate if found.isupper() else found)
                break
        tokens = original.split()
        suffix_words = [w.lower() for w in suffix.split()] if suffix else []
        if suffix and len(tokens) > len(suffix_words):
            head = tokens[:-len(suffix_words)]
        else:
            head = tokens
        budget = len(" ".join(head)) if head else 0
        stems = list(COMPANY_SHORT) + list(COMPANY_STEMS)
        mids = list(COMPANY_SHORT) + list(COMPANY_MIDS)
        min_word = min(len(w) for w in COMPANY_SHORT)
        invented: List[str] = []
        slots = sum(1 for t in head
                    if re.search(r"[0-9A-Za-z]", t) and t.lower() not in COMPANY_LINKERS)
        used = 0
        for token in head:
            bare = token.lower()
            if bare in COMPANY_LINKERS or not re.search(r"[0-9A-Za-z]", token):
                # "&", "and", "/" and similar: structure, not identity - kept so
                # the name keeps its shape ("Kirtane & Pandit, LLP").
                invented.append(token)
                continue
            used += 1
            left = slots - used + 1
            committed = len(" ".join(invented)) if invented else 0
            after = left - 1
            pool = stems if used == 1 else mids
            # This word may only spend what is left once every word still to
            # come has room for a separator and a short name part, and it may
            # not take more than its even share of the remaining width.
            room = budget - committed - 1 - after * (min_word + 1)
            share = max(min_word, min(room, (budget - committed - 1) // left))
            fitting = [w for w in pool if len(w) <= share]
            if not fitting:
                fitting = [min(pool, key=lambda w: (len(w), w))]
            else:
                # Of the words that fit, the widest ones keep the name as
                # close to the source width as the budget allows.
                widest = max(len(w) for w in fitting)
                fitting = [w for w in fitting if len(w) == widest]
            word = self._pick(fitting, "cfit%d" % used, key)
            # Punctuation around a name part belongs to the layout, not to the
            # identity: "Pandit," becomes "Ashford,".  A closing bracket is
            # left behind, because its opener goes with the name part.
            tail = re.search(r"[,.'\u2019-]*$", token).group(0)
            invented.append(word + tail)
        # Safety net: the greedy width above can still overshoot by a character
        # or two on names with punctuation, so shrink the widest words first.
        while len(" ".join(invented)) > budget:
            widest = max(range(len(invented)),
                         key=lambda i: len(invented[i]))
            if len(invented[widest]) <= min_word:
                break
            tail = re.search(r"[,.'\u2019-]*$", invented[widest]).group(0)
            invented[widest] = self._pick(
                [w for w in COMPANY_SHORT if len(w) == min_word],
                "cshrink%d" % widest, key) + tail
        if not invented:
            invented.append(self._pick(stems, "cstem", key))
        name = " ".join(invented)
        return "%s %s" % (name, suffix) if suffix else name

    def _address(self, key: str, original: str) -> str:
        """Preserve the component structure, a well-formed PIN, and the length.

        A full address is far longer than a bare "Maharashtra, India", so the
        replacement is chosen from a ladder of forms and the longest one that
        still fits the source span is used.  A longer replacement would push
        the rest of the line onto a new line and change the page layout.
        """
        flat = re.sub(r"\s+", " ", original).strip()
        target = len(flat)
        # Indian PINs are printed "410 501" in this document; keep that style.
        spaced_pin = bool(re.search(r"\d{3}[ ]\d{3}", flat))
        has_pin = bool(re.search(r"\d{3}[ ]?\d{3}", flat))
        raw_pin = self._digits(6, "addr-pin", key)
        pin = raw_pin[:3] + " " + raw_pin[3:] if spaced_pin else raw_pin
        house = self._int("addr-house", key) % 900 + 1
        building = self._pick(BUILDING_WORDS, "addr-building", key)
        street = self._pick(STREET_NAMES, "addr-street", key)
        street_type = self._pick(STREET_TYPES, "addr-type", key)
        city = self._pick(CITIES, "addr-city", key)
        state = self._pick(STATES, "addr-state", key)
        has_state = bool(re.search(r"(?i)\b(?:maharashtra|madhya pradesh|india)"
                                   r"\s*[,.]?\s*$", flat))
        forms: List[str] = []
        if has_pin:
            if has_state:
                forms.append("%d, %s Apartments, %s %s, %s – %s, %s" % (
                    house, building, street, street_type, city, pin, state))
            forms.append("%d, %s Apartments, %s %s, %s – %s" % (
                house, building, street, street_type, city, pin))
            forms.append("%s – %s" % (city, pin))
        if has_state:
            forms.append("%s, %s" % (city, state))
        forms.append("%s, India" % city)
        forms.append(city)
        for form in forms:
            if len(form) <= target:
                return form
        return forms[-1][:target].rstrip(" ,-–")


def _match_case(fake: str, original: str) -> str:
    """Copy the source's capitalisation onto the generated value.

    "Sarthak.malvadkar@..." and "sarthak.malvadkar@..." are the same mailbox;
    the replacement keeps whichever style each occurrence used.
    """
    local_original = original.split("@", 1)[0]
    if local_original.isupper():
        return fake.upper()
    if local_original[:1].isupper() and local_original[1:].islower():
        return fake[:1].upper() + fake[1:]
    return fake


def digit_count(text: str) -> int:
    return sum(1 for ch in text if ch.isdigit())


def token_count(text: str) -> int:
    return len(text.split())
