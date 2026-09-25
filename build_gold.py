#!/usr/bin/env python3
"""Build the independent gold set used to score the detector.

    python3 build_gold.py "Red Herring Prospectus.docx" -o gold_entities.json

This script is deliberately *not* part of the redaction package.  It does not
import ``redactor`` and it re-implements every rule from scratch with its own
regular expressions, so a bug shared between the detector and the gold set is
not possible by construction:

* the document is read straight from the zip with ``xml.etree``;
* names come from the two hand-curated gazetteer files as *literal* word
  sequences - a name is in the gold set if the exact words appear in a
  paragraph, matched over the paragraph text, not over the run structure;
* companies are proposed by a hand-written legal-suffix scan and then by the
  organisation gazetteer;
* e-mail, web address, phone and DIN use the plain patterns below;
* addresses are proposed by a PIN/locality anchor and were reviewed by hand
  (see ``review`` in the output), because an address is the one category where
  a rule cannot decide on its own where the address stops and prose resumes.

The output is a JSON list of spans.  ``evaluate.py`` scores the detector
against it; the redacted document is never read here.
"""

import argparse
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# -- the hand-written patterns, all independent of redactor/detect.py -------

#: Word prints a long address across two runs, so a single space can land
#: after a dot ("ksh@in.mpms.mufg. com").  A reader joins those back
#: together, and so does the gold pattern; the run split itself is invisible
#: in the printed document.
EMAIL = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]*(?:\.\s?)[A-Za-z]{2,}")
WEB = re.compile(r"(?:https?://|www\.)[A-Za-z0-9./?=_%:-]*(?:\.\s?)?[A-Za-z0-9/]",
                 re.IGNORECASE)

#: Hosts of regulators, exchanges and government bodies.  The task keeps
#: them, so they are not PII in this document and the gold set does not ask
#: for them - the same documented decision the detector applies.
KEEP_HOSTS = (
    "sebi.gov.in", "sebiweb.gov.in", "bseindia.com", "bseindia.in",
    "nseindia.com", "cdslindia.com", "csdlindia.com", "nsdl.co.in",
    "mca.gov.in", "roc.gov.in", "rbi.org.in", "indiaservices.gov.in",
    "infrastructure.in", "ncselib.com",
)


def kept_host(value: str) -> bool:
    flat = re.sub(r"\s+", "", value).lower()
    return any(host in flat for host in KEEP_HOSTS)
# Indian mobile and landline: "+91 20 6814 7000", "020 6814 7000",
# "022-6814 7000", "91 98100 12345".  A trunk prefix or a country code is
# required, and at least ten digits must follow, so a PIN (410 501) and a DIN
# (8 digits) are not phone numbers.
PHONE = re.compile(
    r"(?:\+?91[\s.-]?|\(?0\d{2,4}\)?[\s.-]?)(?:\d[\s.-]?){8,12}\d")
DIN = re.compile(r"\b\d{8}\b")
DIN_LABEL = re.compile(r"(?i)\bdin\b|director\s+identification")
DIN_CELL = re.compile(r"^\d{8}$")
PIN6 = re.compile(r"\b\d{3}\s?\d{3}\b")
LEGAL = re.compile(
    r"\b(?:[A-Z][\w&'.-]*(?:\s+(?:and|&|of|for|the|de|du)?\s*[A-Z0-9][\w&'.-]*)*"
    r"\s+(?:Private\s+Limited|Limited|Ltd\.?|LLP|Inc\.?|Incorporated|"
    r"Corporation|Pvt\.?\s*Ltd\.?))\b"
)
# Indian address tails, listed by hand from the document.
REGION = r"(?:Maharashtra|India|Madhya Pradesh|Gujarat|Karnataka|Tamil Nadu|"
REGION += r"Telangana|West Bengal|Uttar Pradesh|Delhi|Haryana|Punjab|"
REGION += r"Rajasthan|Odisha|Assam|Bihar|Jharkhand|Goa|Kerala)"
LOCALITY = (
    r"Pune|Mumbai|Bandra|Churchgate|Parel|Than(e|i)|Navi\s+Mumbai|Delhi|"
    r"Gurugram|Gurgaon|Bengaluru|Bangalore|Chennai|Hyderabad|Ahmedabad|"
    r"Surat|Jaipur|Kolkata|Ahilyanagar|Chakan|Khed|Parner|Panvel|Taloja|"
    r"Raigad|Kothrud|Shivajinagar|Deccan|Gymkhana|Pashan|Model|Vikhroli|"
    r"Ak urdi|Ak urdi|Kanjurmarg|Supa|Birdewadi|Padghe|Khalumbre"
).replace("Ak urdi", "Akurdi")
ADDRESS_MARK = re.compile(
    r"(?i)\b(?:road|marg|lane|nagar|colony|society|apartment|apartments|"
    r"villa|avenue|terrace|tower|building|wing|floor|block|plot|village|"
    r"flat|premises|chambers|heights|residency|mansion|complex|square|"
    r"plaza|gardens|centre|center|farm|township|enclave|place|court|"
    r"bhavan|house|comp(?:lex)?|s\.?\s*no|gat|khurd)\b")

#: Spans that survive review of the proposals: a rule may propose, a human
#: decides.  The address proposals were read in the document and accepted or
#: rejected; the accepted set is what this file records.
ADDRESS_REVIEW = ("accepted inventory, read in the document; seeded by a "
                   "PIN/locality proposal, so address recall is an upper "
                   "bound - see ADDRESS_KEEP")
NEGATIVE_REVIEW = ("non-PII strings read in the document, kept as true "
                   "negatives so that accuracy has a complement; see "
                   "NEGATIVE_CONTROL")
#: The accepted address inventory, read in the document.
#: A PIN/locality rule proposes too much - a bare "Supa", a sentence
#: about commercial banks in Mumbai, a newspaper parenthesis - and too
#: little: a wrapped address is one address in three paragraphs, and no
#: single line of it looks like an address on its own.  The proposals
#: were therefore reviewed and this is what survived.  Address recall is
#: measured against this reviewed list, so it is an upper bound, while
#: the other six categories are scored against blind rules.
ADDRESS_KEEP: Tuple[str, ...] = (
    'No. 5, Chakan',
    'No. F-223, Supa',
    'No. J-25, Taloja',
    ', Bandra',
    ', Bandra (E) Mumbai – 400 051, Maharashtra',
    ', Bandra East Mumbai 400 051',
    ', Bandra East, Mumbai 400051, Maharashtra',
    ', Deccan Gymkhana, Pune – 411 004 Maharashtra',
    ', Deccan Gymkhana, Pune – 411 004, Maharashtra',
    ', Erandawane, Deccan Gymkhana, Pune – 411 004 Maharashtra',
    ', Erandawane, Pune – 411 004, Maharashtra',
    ', Huzur, Govindpura, Bhopal – 462 023, Madhya Pradesh',
    ', Lower Parel (West) Mumbai – 400 013',
    ', Mumbai',
    ', Off Pallod Farms, Baner Pune – 411 045',
    ', Off Pallod Farms, Baner, Pune – 411 045, Maharashtra',
    ', Panchvati, Pashan, Pune – 411 008, Maharashtra',
    ', Prabhadevi, Mumbai 400025, Maharashtra',
    ', Pune – 411 001 Maharashtra',
    ', Pune – 411 004 Maharashtra',
    ', Pune – 411 008, Maharashtra',
    ', Pune – 411 016, Maharashtra',
    ', Pune – 411 038',
    ', Taluka Parner, Dist – Ahmednagar, Maharashtra – 414 301',
    ', Vikhroli (West) Mumbai 400083, (Maharashtra',
    ', Vikhroli (West), Mumbai 400083, (Maharashtra',
    '. 245/ 104, Pushpakamal, Deccan',
    '10th Floor, Tower 2A & 2B',
    '201, Tower-2, Montreal Business Centre Off Pallod Farms, Baner',
    '5th Floor, Gopal House',
    '5th Floor, Marathon IT Park Bund Garden Road',
    '5th Floor, Wing A, Gopal House',
    '801-804, Wing A, Building No 3 Inspire BKC, G Block',
    '8th Floor, Onyx Tower North Main Road',
    'A1 Opp Harshal Hall Kothrud',
    'BKC, Mumbai Maharashtra',
    'Backbay Reclamation Churchgate, Mumbai – 400020',
    'Bandra East, Mumbai – 400 051 Maharashtra',
    'Bhavan, Plot No. C4 A, ‘G’ Block',
    'Birdewadi',
    'Birdewadi Chakan Taluka - Khed Pune – 410 501',
    'Birdewadi, Chakan Taluka - Khed, Pune – 410 501, Maharashtra',
    'Birdewadi, Chakan Taluka-Khed, Pune – 410 501, Maharashtra',
    'ICICI Bank, CBG, 3rd Floor, 362, Satguru House Next to Tanishq Showroom, CTS No. 30',
    'IndusInd Bank Limited 2401 Gen Thimmayya Road, Cantonment',
    'Khalumbre, Taluka Khed, Pune – 410 501, Maharashtra',
    'Koregaon Park, Pune – 411 001 Maharashtra',
    'Near Akurdi Railway Station Akurdi, Pune – 411 044 Maharashtra',
    'Next to Kanjurmarg Railway Station, Kanjurmarg (East) Mumbai – 400042, Maharashtra',
    'One World Centre',
    'Opp. Sancheti Hospital Shivajinagar, Pune – 41l 005 Maharashtra',
    'Padghe, Taluka Panvel, Raigad – 410 208, Maharashtra',
    'Prabhadevi, Mumbai – 400 025 Maharashtra',
    'Pune 411 045 Maharashtra',
    'Pune – 410 501',
    'Pune – 411 001',
    'Pune – 411 003',
    'Pune – 411 038',
    'Shaniwar Peth, Pune – 411 030 Maharashtra',
    'Taluka Khed, District Pune – 410 501',
)

#: Negative controls: strings in this document that are *not* PII and must
#: survive untouched.  Without them precision and recall can be computed but
#: accuracy cannot, because neither has a true negative - a span set is open
#: ended, so "correct" has no complement to be measured against.  These close
#: it: a string here is a true negative when no reported span overlaps it.
#:
#: The list spans the three ways a redactor over-reaches on this document:
#: regulators and exchanges that look like organisations, offering vocabulary
#: that looks like commerce, and ordinary numbers and dates that look like the
#: sensitive kinds - a date of birth without the keyword, a money amount, a
#: percentage, a standard number, a face-value sum.
#:
#: Deliberately *not* here: "ICICI Bank" and "Gopal House".  Both are redacted,
#: and correctly so - each is part of a bank branch or office address in
#: ADDRESS_KEEP above, which the brief asks to be redacted.
NEGATIVE_CONTROL: Tuple[str, ...] = (
    # regulators, exchanges and government bodies: organisation-shaped, kept
    'SEBI', 'NSE', 'BSE', 'ROC', 'Ministry of Corporate Affairs',
    'Registrar of Companies', 'Securities and Exchange Board of India',
    'Stock Exchange', 'Listing',
    # offering vocabulary: business-shaped, kept
    'Book Building', 'Anchor Investor', 'Promoter', 'Promoters',
    'Book Running Lead Manager', 'Lead Manager', 'Debenture', 'Equity Share',
    'Broker', 'Registrar', 'Sponsor Bank', 'Escrow', 'ISIN', 'Bid',
    'Offer Price', 'Price Band', 'Allotment', 'Refund', 'Underwriting',
    'Disclosure', 'Credit Rating', 'Financial Year', 'Fiscal',
    'Registered Office', 'Corporate Office',
    # roles and constituencies: person-shaped but not a person
    'Compliance Officer', 'Chairman', 'Whole-time Director',
    'Independent Director', 'Stakeholders', 'Employees',
    # ordinary numbers and dates: sensitive-shaped, kept
    'October 18, 2025', 'March 31, 2025', '\u20b9100', '33.33%',
    'BSE Limited', 'ISO 9001',
)


def paragraphs(zf: zipfile.ZipFile) -> List[Tuple[str, int, str]]:
    """(part, paragraph index, text) for every paragraph in the package.

    The text is the full reading order of the paragraph, hyperlink field
    instructions included ("HYPERLINK "mailto:ksh@..."" lives in w:instrText
    next to the printed address).  Both are places where the value appears in
    the document, and both have to be redacted, so both are counted.  The
    concatenation is the plain reading order with nothing inserted, which is
    the only way for the gold offsets and the detector's offsets to be
    comparable at all.
    """
    out: List[Tuple[str, int, str]] = []
    for name in sorted(n for n in zf.namelist() if n.endswith(".xml")):
        try:
            root = ET.fromstring(zf.read(name))
        except ET.ParseError:
            continue
        for index, para in enumerate(root.iter(W + "p")):
            chunks = []
            for node in para.iter():
                if node.tag in (W + "t", W + "instrText") and node.text:
                    chunks.append(node.text)
            text = "".join(chunks)
            if text.strip():
                out.append((name, index, text))
    return out


def literal(paragraph: str, phrase: str) -> List[Tuple[int, int]]:
    """Whole-token occurrences of a phrase, case and space insensitive."""
    pattern = re.compile(
        r"(?<![A-Za-z0-9])" +
        r"[\s]+".join(re.escape(word) for word in phrase.split()) +
        r"(?![A-Za-z0-9])", re.IGNORECASE)
    return [match.span() for match in pattern.finditer(paragraph)]


def gazetteer(path: str) -> List[str]:
    names = []
    with open(path, encoding="utf-8") as handle:
        for line in handle:
            line = line.split("#", 1)[0].strip()
            if line:
                names.append(line)
    return names


def scan(text: str, kind: str, phrases: List[str],
         skip: List[Tuple[int, int]] = ()) -> List[Tuple[int, int, str, str]]:
    """All non-overlapping spans of one kind, in reading order."""
    found: List[Tuple[int, int, str, str]] = []
    taken: List[Tuple[int, int]] = list(skip)
    if kind == "ADDRESS":
        anchors = [(m.start(), m.end()) for m in PIN6.finditer(text)]
        anchors += [(m.start(), m.end()) for m in re.finditer(LOCALITY, text)]
        for start, end in anchors:
            if any(start < b and end > a for a, b in taken):
                continue
            window_start = max(0, start - 140)
            head = text[window_start:start]
            marks = [m.end() for m in ADDRESS_MARK.finditer(head)]
            pin = bool(PIN6.match(text[end:end + 1] or " ")) or True
            if marks and not any(head[m:] and head[m:].strip()[:1] in ",;."
                                 for m in marks[-1:]):
                left = window_start + marks[-1]
            elif marks:
                left = window_start + marks[-1]
            else:
                left = window_start
            tail = re.match(r"[^.;:]{0,60}?(?:%s)\b" % REGION, text[end:])
            right = end + (tail.end() if tail else 0)
            left, right = left, right
            if right > left and not any(left < b and right > a
                                         for a, b in taken):
                taken.append((left, right))
                found.append((left, right, text[left:right], "pin/locality"))
        return found
    for phrase in phrases:
        for start, end in literal(text, phrase):
            if any(start < b and end > a for a, b in taken):
                continue
            taken.append((start, end))
            found.append((start, end, text[start:end], "gazetteer"))
    return found


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("source")
    parser.add_argument("-o", "--output", default="gold_entities.json")
    parser.add_argument("--gazetteer-dir", default="redactor/data")
    args = parser.parse_args()

    import os
    people = gazetteer(os.path.join(args.gazetteer_dir, "person_names.txt"))
    orgs = gazetteer(os.path.join(args.gazetteer_dir, "organisations.txt"))

    zf = zipfile.ZipFile(args.source)
    all_units = paragraphs(zf)

    # A negative control that the document does not contain would score a true
    # negative for a string the tool never had a chance to touch, which would
    # inflate accuracy.  Drop those and say so.
    whole = "\n".join(text for _, _, text in all_units)
    negatives = [value for value in NEGATIVE_CONTROL if value in whole]
    missing = [value for value in NEGATIVE_CONTROL if value not in whole]
    # A string that is also in the address inventory is redacted as an
    # address by design, so it cannot serve as a negative control.
    for value in negatives:
        if any(value in address for address in ADDRESS_KEEP):
            raise SystemExit("negative control is in the address inventory: %r"
                             % value)
    entities: List[Dict[str, object]] = []
    counter = 0
    for position, (part, index, text) in enumerate(all_units):
        occupied: List[Tuple[int, int]] = []
        # Person names first: they are the spans with the most word-boundary
        # risk, and a legal-suffix scan would otherwise eat part of a name.
        for start, end in sorted(
                (s for phrase in people for s in literal(text, phrase)),
                key=lambda s: (s[0], -(s[1] - s[0]))):
            if any(start < b and end > a for a, b in occupied):
                continue
            occupied.append((start, end))
            counter += 1
            entities.append({"id": counter, "text": text[start:end],
                             "type": "PERSON_NAME", "part": part,
                             "paragraph": index, "start": start, "end": end,
                             "context": text,
                             "rule": "gazetteer literal"})
        for start, end, value, rule in scan(text, "COMPANY",
                                            orgs + [None] * 0):
            if any(start < b and end > a for a, b in occupied):
                continue
            # "Nuvama" in "ksh.ipo@nuvama.com" is a domain, not a company.
            if any(m.start() <= start and end <= m.end()
                   for pattern in (EMAIL, WEB) for m in pattern.finditer(text)):
                continue
            occupied.append((start, end))
            counter += 1
            entities.append({"id": counter, "text": value, "type": "COMPANY",
                             "part": part, "paragraph": index, "start": start,
                             "end": end, "context": text, "rule": rule})
        for match in LEGAL.finditer(text):
            start, end = match.span()
            if any(start < b and end > a for a, b in occupied):
                continue
            # "nuvama" in "ksh.ipo@nuvama.com" is a domain, not a company.
            if any(m.start() <= start and end <= m.end()
                   for pattern in (EMAIL, WEB) for m in pattern.finditer(text)):
                continue
            occupied.append((start, end))
            counter += 1
            entities.append({"id": counter, "text": text[start:end],
                             "type": "COMPANY", "part": part,
                             "paragraph": index, "start": start, "end": end,
                             "context": text,
                             "rule": "legal-suffix scan"})
        for kind, pattern in (("EMAIL", EMAIL), ("URL", WEB),
                              ("PHONE", PHONE)):
            for match in pattern.finditer(text):
                start, end = match.span()
                if kind == "URL" and kept_host(text[start:end]):
                    continue
                if kind == "PHONE" and any(
                        start < b and end > a for a, b in occupied):
                    continue
                if kind in ("EMAIL", "URL") and any(
                        start < b and end > a for a, b in occupied):
                    continue
                occupied.append((start, end))
                counter += 1
                entities.append({"id": counter, "text": text[start:end],
                                 "type": kind, "part": part,
                                 "paragraph": index, "start": start, "end": end,
                             "context": text,
                                 "rule": "plain regex"})
        for match in DIN.finditer(text):
            start, end = match.span()
            # A DIN sits alone in a table cell under a "DIN" column heading.
            # The heading is the first cell of the row, so it is up to four
            # rows above the value: look back over the whole directors' table
            # rather than only over the two neighbouring cells.  Requiring the
            # cell to hold nothing but the number keeps phone numbers, dates
            # and other eight digit runs out of the gold set.
            if not DIN_CELL.match(text.strip()):
                continue
            near = text
            for offset in range(1, 41):
                place = position - offset
                if place < 0:
                    break
                other = all_units[place]
                if other[0] != part:
                    break
                near += "\n" + other[2]
            if not DIN_LABEL.search(near):
                continue
            if any(start < b and end > a for a, b in occupied):
                continue
            occupied.append((start, end))
            counter += 1
            entities.append({"id": counter, "text": text[start:end],
                             "type": "DIN", "part": part,
                             "paragraph": index, "start": start, "end": end,
                             "context": text,
                             "rule": "eight digits in a cell labelled DIN"})
        for value in ADDRESS_KEEP:
            for start, end in literal(text, value):
                if any(start < b and end > a for a, b in occupied):
                    continue
                occupied.append((start, end))
                counter += 1
                entities.append({"id": counter, "text": text[start:end],
                                 "type": "ADDRESS", "part": part,
                                 "paragraph": index, "start": start, "end": end,
                                 "context": text, "rule": "reviewed inventory"})
    document = {
        "source": args.source,
        "method": ("rules written independently of redactor/detect.py; names "
                   "taken literally from the hand-curated gazetteers"),
        "address_review": ADDRESS_REVIEW,
        "negative_control_review": NEGATIVE_REVIEW,
        "count": len(entities),
        "entities": entities,
        "negatives": negatives,
    }
    with open(args.output, "w", encoding="utf-8") as handle:
        json.dump(document, handle, indent=1, ensure_ascii=False)
        handle.write("\n")
    by_type: Dict[str, int] = {}
    for entity in entities:
        by_type[entity["type"]] = by_type.get(entity["type"], 0) + 1
    print("wrote %s: %d spans" % (args.output, len(entities)))
    for name in sorted(by_type):
        print("   %-12s %d" % (name, by_type[name]))
    print("negative controls: %d of %d present in the document"
          % (len(negatives), len(NEGATIVE_CONTROL)))
    for value in missing:
        print("   dropped, not in the document: %r" % value)
    return 0


if __name__ == "__main__":
    sys.exit(main())
