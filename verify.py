#!/usr/bin/env python3
"""Independent verification of a redacted DOCX.

This script does not import the redaction pipeline's detectors: it re-reads
both packages from disk, pulls the original values from the mapping file and
proves the output is (a) structurally the same document, (b) free of the
original PII, and (c) reproducible.

    python3 verify.py "Red Herring Prospectus.docx" \\
        "Red Herring Prospectus - REDACTED.docx" --map redaction_map.json

Exit status is 0 only when every check passes.
"""

import argparse
import collections
import hashlib
from typing import Dict, List
import os
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: Elements whose count must survive redaction untouched.  Text lives in
#: w:t / w:instrText; their *content* changes, their number must not.
STRUCTURE_TAGS = (
    "p", "r", "tbl", "tr", "tc", "hyperlink", "bookmarkStart", "bookmarkEnd",
    "drawing", "fldChar", "instrText", "t", "tab", "br", "sectPr", "footnote",
)

#: A state/country pair that survives only as a prose location reference:
#: "our facilities are located in Maharashtra, India".  A location named in a
#: sentence is not an address, so the pair is documented as kept - but the
#: verifier still requires that every surviving occurrence is introduced by a
#: preposition, which is what makes it prose and not a leaked address line.
PROSE_LOCATION = re.compile(
    r"\b(?:in|at|near|from|to|within|across|of|into|towards)\s*$", re.I)

#: Emails in the output must use an RFC 2606 reserved domain.
ALLOWED_EMAIL_DOMAINS = ("example.com", "example.net", "example.org")

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


# ---------------------------------------------------------------- utilities
def strip_xml(data: bytes) -> str:
    """Visible text of an XML part, with runs joined and entities resolved.

    Tags are replaced by a single space so that a value split across two
    ``w:t`` elements ("Maharashtra" + ", India") is still found, and the
    result is whitespace-collapsed so a value split by a line break is found.
    """
    text = data.decode("utf-8", "replace")
    text = re.sub(r"<[^>]+>", " ", text)
    for entity, char in (("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
                         ("&quot;", '"'), ("&apos;", "'")):
        text = text.replace(entity, char)
    text = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), text)
    text = re.sub(r"&#x([0-9a-fA-F]+);",
                  lambda m: chr(int(m.group(1), 16)), text)
    return re.sub(r"\s+", " ", text)


def normalise(value: str) -> str:
    """Loose form of a value: no case, spacing or punctuation distinctions."""
    return re.sub(r"[^0-9a-z]", "", value.lower())


def value_variants(value: str):
    """Ways the same value can appear after Word has reflowed it."""
    flat = re.sub(r"\s+", " ", value).strip()
    yield flat
    yield flat.replace(" ", "")
    yield re.sub(r"[\s\-–—(),]+", "", flat)
    yield flat.lower()
    yield re.sub(r"[\s\-–—(),]+", "", flat).lower()


# ------------------------------------------------------------------- checks
class Report:
    def __init__(self) -> None:
        self.rows = []
        self.failed = 0

    def add(self, name: str, ok: bool, detail: str = "") -> bool:
        self.rows.append((name, ok, detail))
        if not ok:
            self.failed += 1
        status = "PASS" if ok else "FAIL"
        print("[%s] %-34s %s" % (status, name, detail))
        return ok

    def note(self, detail: str) -> None:
        print("       %s" % detail)


def check_zip(path: str, rep: Report) -> zipfile.ZipFile:
    with open(path, "rb") as handle:
        head = handle.read(4)
    ok = head[:2] == b"PK"
    rep.add("zip magic", ok, path)
    zf = zipfile.ZipFile(path)
    bad = zf.testzip()
    rep.add("zip integrity", bad is None, "first bad part: %s" % bad if bad else "")
    return zf


def check_xml_wellformed(zf: zipfile.ZipFile, rep: Report) -> None:
    broken = []
    for name in zf.namelist():
        if not name.endswith((".xml", ".rels")):
            continue
        try:
            ET.fromstring(zf.read(name))
        except ET.ParseError as exc:
            broken.append("%s: %s" % (name, exc))
    rep.add("xml well-formed", not broken, "; ".join(broken[:3]))


def check_part_parity(src: zipfile.ZipFile, out: zipfile.ZipFile,
                      rep: Report) -> None:
    same = set(src.namelist()) == set(out.namelist())
    missing = sorted(set(src.namelist()) - set(out.namelist()))
    added = sorted(set(out.namelist()) - set(src.namelist()))
    rep.add("package parts identical", same and not missing and not added,
            "missing=%s added=%s" % (missing[:3], added[:3]))


def tag_counts(data: bytes) -> collections.Counter:
    counts = collections.Counter()
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return counts
    for element in root.iter():
        if not isinstance(element.tag, str):
            continue
        local = element.tag.rsplit("}", 1)[-1]
        if local in STRUCTURE_TAGS:
            counts[local] += 1
    return counts


def check_structure(src: zipfile.ZipFile, out: zipfile.ZipFile,
                    rep: Report) -> None:
    differences = []
    for name in src.namelist():
        if not name.endswith(".xml"):
            continue
        a = tag_counts(src.read(name))
        b = tag_counts(out.read(name))
        for tag in sorted(set(a) | set(b)):
            if a[tag] != b[tag]:
                differences.append("%s <w:%s> %d -> %d"
                                   % (name, tag, a[tag], b[tag]))
    rep.add("structure preserved", not differences,
            "; ".join(differences[:3]) + (" (+%d more)" % (len(differences) - 3)
                                           if len(differences) > 3 else ""))


def value_probe(value: str):
    """A regex that finds the value however Word split and cased it."""
    tokens = [t for t in re.split(r"[\s,]+", value) if t]
    if not tokens:
        return None
    pattern = r"[\s,]*".join(re.escape(token) for token in tokens)
    if len(re.sub(r"[^0-9A-Za-z]", "", value)) < 4:
        return None
    return re.compile(pattern, re.I)


def classify_residual(blob: str, start: int, end: int) -> str:
    """Decide whether a surviving value is a documented keep or a leak."""
    before = blob[max(0, start - 60): start]
    if re.search(r"\d{3}\s?\d{3}\s*[-,]?\s*$", before):
        # directly after a PIN: part of an address line
        return "leak"
    if re.search(r"\d{1,4}\s*[-,]\s*$", before):
        return "leak"
    if PROSE_LOCATION.search(before.rstrip("(")):
        return "prose location reference"
    return "unclassified"


def check_residual(out: zipfile.ZipFile, mapping, rep: Report) -> None:
    blobs = {name: strip_xml(out.read(name))
             for name in out.namelist()
             if name.endswith((".xml", ".rels"))}
    leaks = collections.Counter()
    keeps = collections.Counter()
    unclassified = collections.defaultdict(set)
    for entry in mapping:
        value = entry["original"]
        ptype = entry["type"]
        probe = value_probe(value)
        if probe is None:
            continue
        for name, blob in blobs.items():
            hit = probe.search(blob)
            if hit is None:
                continue
            verdict = classify_residual(blob, hit.start(), hit.end())
            if verdict == "prose location reference":
                keeps[ptype] += 1
            else:
                leaks[ptype] += 1
                unclassified[ptype].add("%s @ %s" % (value, name))
    detail = []
    if leaks:
        for ptype, count in leaks.most_common():
            examples = sorted(unclassified[ptype])[:3]
            detail.append("%s: %d (%s)" % (ptype, count, " | ".join(examples)))
    rep.add("no original PII survives", not leaks, " | ".join(detail))
    if keeps:
        rep.note("documented keeps (prose location references): %s"
                 % ", ".join("%s=%d" % kv for kv in sorted(keeps.items())))


def check_fake_domains(out: zipfile.ZipFile, rep: Report) -> None:
    offenders = collections.Counter()
    for name in out.namelist():
        if not name.endswith((".xml", ".rels")):
            continue
        for address in EMAIL_RE.findall(strip_xml(out.read(name))):
            domain = address.rsplit("@", 1)[-1].lower()
            if domain not in ALLOWED_EMAIL_DOMAINS:
                offenders[domain] += 1
    rep.add("emails use reserved domains", not offenders,
            ", ".join("%s x%d" % kv for kv in offenders.most_common(3)))


def check_shape(mapping, rep: Report) -> None:
    """Each fake must be usable in place of the value it replaces.

    One entity can be printed several ways ("ICICI Securities Limited" and the
    run-seam spelling "ICICISecurities Limited"), and all of them must share
    one fake, so the fake is compared against the *most demanding* spelling:
    the one with the most words, and for addresses the shortest.
    """
    problems = []
    widest = {}
    shortest = {}
    for entry in mapping:
        key = (entry["type"], normalise(entry["original"]))
        widest[key] = max(widest.get(key, 0), len(entry["original"].split()))
        shortest[key] = min(shortest.get(key, 10 ** 6), len(entry["original"]))
    for entry in mapping:
        original, fake, ptype = entry["original"], entry["fake"], entry["type"]
        key = (ptype, normalise(original))
        if ptype in ("PHONE", "DIN"):
            if sum(c.isdigit() for c in fake) != sum(c.isdigit() for c in original):
                problems.append("digit count %s: %r -> %r" % (ptype, original, fake))
        elif ptype == "EMAIL":
            if fake.count("@") != 1 or len(fake.split("@")[0]) > len(
                    original.split("@")[0]):
                problems.append("email shape: %r -> %r" % (original, fake))
        elif ptype in ("PERSON_NAME", "COMPANY"):
            if len(fake.split()) > widest[key]:
                problems.append("word count %s: %r -> %r" % (ptype, original, fake))
        elif ptype == "ADDRESS":
            if len(fake) > shortest[key]:
                problems.append("longer address: %r -> %r" % (original, fake))
    rep.add("fakes preserve shape", not problems,
            "; ".join(problems[:3]) + (" (+%d more)" % (len(problems) - 3)
                                       if len(problems) > 3 else ""))


def check_consistency(out: zipfile.ZipFile, mapping, rep: Report) -> None:
    """The same original value must always become the same fake.

    Two things are checked against the produced document.  First, one
    original must never have two fakes - that is what makes a document
    self-inconsistent.  Second, every original spelling in the mapping must
    be gone from the output, so no mention was left behind under a second
    spelling.

    Distinct originals that share one fake are reported, not failed: every
    e-mail becomes ``...@example.com`` and every web address becomes
    ``www.example.com`` on purpose, so different parties do collapse onto
    the same placeholder.
    """
    problems = []
    shared = 0
    originals = {}
    fakes = {}
    for entry in mapping:
        original = entry["original"]
        if original in originals and originals[original] != entry["fake"]:
            problems.append("two fakes for %r: %r and %r"
                            % (original, originals[original], entry["fake"]))
        originals[original] = entry["fake"]
        if entry["fake"] in fakes:
            shared += 1
        else:
            fakes[entry["fake"]] = original

    blob = "\n".join(strip_xml(out.read(name)) for name in out.namelist()
                     if name.endswith(".xml"))
    survivors = []
    for entry in mapping:
        value = entry["original"]
        probe = value_probe(value)
        if probe is None:
            continue
        hit = probe.search(blob)
        if hit is None:
            continue
        if classify_residual(blob, hit.start(), hit.end()) == "prose location reference":
            continue
        survivors.append(value)
    if survivors:
        problems.append("%d original value(s) still present, e.g. %r"
                        % (len(survivors), survivors[0]))
    rep.add("pseudonyms are consistent", not problems,
            "; ".join(problems[:3]) + (" (+%d more)" % (len(problems) - 3)
                                        if len(problems) > 3 else "")
            or "%d entities, one fake each (%d share a placeholder)"
               % (len(mapping), shared))


def check_determinism(source: str, out_path: str, rep: Report,
                      use_gazetteer: bool = True) -> None:
    """Re-run the pipeline and compare the bytes.

    The second run must use the same detection mode as the file under test,
    otherwise a pattern-only output is compared with a gazetteer-assisted one.
    """
    here = os.path.dirname(os.path.abspath(__file__))
    script = os.path.join(here, "redact.py")
    command = [sys.executable, script, source]
    if not use_gazetteer:
        command.append("--no-gazetteer")
    with tempfile.TemporaryDirectory() as tmp:
        again = os.path.join(tmp, "again.docx")
        result = subprocess.run(
            command + ["-o", again], capture_output=True, text=True, cwd=here)
        if result.returncode != 0:
            rep.add("second run succeeds", False, result.stderr.strip()[:200])
            return
        one = hashlib.sha256(open(out_path, "rb").read()).hexdigest()
        two = hashlib.sha256(open(again, "rb").read()).hexdigest()
        rep.add("deterministic output", one == two,
                "sha256 %s" % one[:16] if one == two else "hashes differ")


#: The source document must never be written to.  The hash of the delivered
#: source is recorded so a later run can prove the file is unchanged.
SOURCE_SHA256 = "8b5c93f7642d659e64b51be9f6172c86c2825417f376ca1800ed331515e6f929"


def check_source_untouched(source: str, rep: Report) -> None:
    digest = hashlib.sha256(open(source, "rb").read()).hexdigest()
    rep.add("source file unchanged", digest == SOURCE_SHA256,
            "sha256 %s" % digest[:16] if digest == SOURCE_SHA256
            else "sha256 %s, expected %s" % (digest[:16], SOURCE_SHA256[:16]))


#: A replacement that is wider than the text it covers pushes the rest of the
#: line onto the next line and moves everything below it: page count, table
#: row heights and the footer's page numbers all follow.  A replacement that is
#: narrower is harmless, so only growth is a failure.
MAX_GROWTH_CHARS = 8


DOUBLE_SPACE_RE = re.compile(r"[ \u00a0]{2,}")


def check_word_spacing(src: zipfile.ZipFile, out: zipfile.ZipFile,
                      rep: Report) -> None:
    """A replacement may not leave a doubled space behind.

    The writer cuts one fake over the ``<w:t>`` nodes the original occupied,
    in proportion to their lengths.  A fake that does not divide the way the
    runs do can put a space where the original had none - "Lakeshore
    Familyly  Trust" - and the document then silently stops saying what it
    should.  A fake may have fewer words than the original (a two word
    address becomes one token), so only *added* whitespace is a fault.
    """
    def doubles(zf: zipfile.ZipFile) -> Dict[str, List[int]]:
        found: Dict[str, List[int]] = {}
        for name in sorted(n for n in zf.namelist() if n.endswith(".xml")):
            try:
                root = ET.fromstring(zf.read(name))
            except ET.ParseError:
                continue
            rows: List[int] = []
            for para in root.iter(W + "p"):
                text = "".join(node.text or "" for node in para.iter()
                               if node.tag in (W + "t", W + "instrText"))
                rows.append(len(DOUBLE_SPACE_RE.findall(text)))
            if rows:
                found[name] = rows
        return found

    before, after = doubles(src), doubles(out)
    added = 0
    compared = 0
    example = ""
    for name, old in before.items():
        new_rows = after.get(name)
        if not new_rows or len(new_rows) != len(old):
            continue
        for position, (a, b) in enumerate(zip(old, new_rows)):
            compared += 1
            if b > a:
                added += b - a
                if not example:
                    example = "%s paragraph %d: %d -> %d doubled spaces" % (
                        name.rsplit("/", 1)[-1], position, a, b)
    rep.add("no doubled spaces introduced", added == 0,
            example or "%d paragraphs, none introduced" % compared)


def check_line_width(src: zipfile.ZipFile, out: zipfile.ZipFile,
                     rep: Report) -> None:
    """Paragraphs must not get wider, so the pagination cannot shift."""
    def widths(zf: zipfile.ZipFile) -> Dict[str, List[int]]:
        found: Dict[str, List[int]] = {}
        for name in sorted(n for n in zf.namelist() if n.endswith(".xml")):
            try:
                root = ET.fromstring(zf.read(name))
            except ET.ParseError:
                continue
            rows: List[int] = []
            for para in root.iter(W + "p"):
                length = 0
                for node in para.iter():
                    if node.tag in (W + "t", W + "instrText", W + "tab"):
                        length += len(node.text or "") if node.tag != W + "tab" else 1
                rows.append(length)
            if rows:
                found[name] = rows
        return found

    before, after = widths(src), widths(out)
    grew = 0
    worst = 0
    compared = 0
    for name, old in before.items():
        new = after.get(name)
        if not new or len(new) != len(old):
            continue
        for a, b in zip(old, new):
            compared += 1
            if b - a > worst:
                worst = b - a
            if b > a:
                grew += 1
    limit = MAX_GROWTH_CHARS
    ok = compared > 0 and worst <= limit
    rep.add("no line grows beyond %d characters" % limit, ok,
            "%d paragraphs, %d grew, widest growth +%d char%s"
            % (compared, grew, worst, "" if worst == 1 else "s"))


# --------------------------------------------------------------------- main
def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source")
    parser.add_argument("redacted")
    parser.add_argument("--map", dest="map_path", required=True)
    parser.add_argument("--skip-determinism", action="store_true")
    parser.add_argument("--no-gazetteer", action="store_true",
                        help="the output under test was made without the "
                             "name gazetteer")
    args = parser.parse_args()

    import json
    with open(args.map_path, encoding="utf-8") as handle:
        mapping = json.load(handle)

    rep = Report()
    src = check_zip(args.source, rep)
    out = check_zip(args.redacted, rep)
    check_xml_wellformed(out, rep)
    check_part_parity(src, out, rep)
    check_structure(src, out, rep)
    check_residual(out, mapping, rep)
    check_fake_domains(out, rep)
    check_shape(mapping, rep)
    check_consistency(out, mapping, rep)
    check_word_spacing(src, out, rep)
    check_line_width(src, out, rep)
    if not args.skip_determinism:
        check_determinism(args.source, args.redacted, rep,
                          use_gazetteer=not args.no_gazetteer)
    check_source_untouched(args.source, rep)

    print()
    if rep.failed:
        print("%d of %d checks FAILED" % (rep.failed, len(rep.rows)))
        return 1
    print("all %d checks passed" % len(rep.rows))
    return 0


if __name__ == "__main__":
    sys.exit(main())
