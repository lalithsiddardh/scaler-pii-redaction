#!/usr/bin/env python3
"""redact.py - replace personally identifiable information in a .docx.

    python3 redact.py "Red Herring Prospectus.docx" -o "Red Herring Prospectus - REDACTED.docx"

The same source value always maps to the same fake value, in every paragraph,
table cell, header, footer and hyperlink field code.  Formatting, images,
numbering, styles and table layout are preserved byte for byte; only the text
characters that carried PII are rewritten.

Options
    --no-gazetteer   pattern-only detection (no curated name/organisation
                     lists).  Used to report gazetteer-free precision/recall.
    --seed TEXT      change the deterministic mapping seed.
    --emit-map PATH  write original -> fake pairs (re-identifies the data;
                     keep it out of any released document).
    --summary PATH   write a JSON summary of what was redacted.
    --dry-run        detect and report, write nothing.
"""

import argparse
import collections
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from redactor import detect, load_gazetteers
from redactor.docx_io import Package
from redactor.replace import Pseudonymiser


def build_units(package: Package) -> detect.ScanContext:
    cells = package.cell_of()
    units = []
    for document in package.text_parts():
        owner = cells.get(document.name, [])
        for paragraph in document.paragraphs:
            if paragraph.text.strip():
                units.append(detect.Unit(document.name, paragraph.index,
                                         paragraph.text,
                                         owner[paragraph.index]
                                         if paragraph.index < len(owner) else -1))
    return detect.ScanContext(units)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("input", help="source .docx")
    parser.add_argument("-o", "--output", help="redacted .docx")
    parser.add_argument("--no-gazetteer", action="store_true",
                        help="pattern-only detection")
    parser.add_argument("--seed", default="scaler-pii-redaction-2026")
    parser.add_argument("--emit-map", help="write original -> fake JSON")
    parser.add_argument("--summary", help="write a JSON summary")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv)

    people, organisations = load_gazetteers()
    config = detect.RedactionConfig(
        use_gazetteer=not args.no_gazetteer,
        person_gazetteer=people,
        company_gazetteer=organisations,
    )

    package = Package(args.input)
    context = build_units(package)
    candidates = detect.scan(context, config)
    print("scanned %d text blocks in %d parts; %d spans"
          % (len(context.units), len(context.parts()), len(candidates)))

    by_type = collections.Counter(c.ptype for c in candidates)
    for ptype, count in sorted(by_type.items(), key=lambda kv: -kv[1]):
        print("  %-16s %d" % (ptype, count))

    if args.dry_run:
        return 0

    pseudonymiser = Pseudonymiser(seed=args.seed)
    replacements = list(_replacements(candidates, pseudonymiser))
    edits = [{
        "part": candidate.part,
        "paragraph": candidate.paragraph,
        "start": candidate.start,
        "end": candidate.end,
        "text": text,
    } for candidate, text in replacements]
    mapping = collections.OrderedDict()
    for candidate, text in replacements:
        value = candidate.entity or candidate.text
        key = "%s|%s" % (candidate.ptype, pseudonymiser.canonical(
            candidate.ptype, value))
        mapping.setdefault(key, {
            "original": value,
            "fake": pseudonymiser.fake(candidate.ptype, value),
            "type": candidate.ptype,
        })

    _assert_no_collision(package, list(mapping.values()))
    touched = package.apply(edits)
    package.scrub_metadata()

    output = args.output or _default_output(args.input)
    package.save(output)
    print("rewrote %d text elements across %d parts" % (touched, len(package.documents)))
    print("wrote %s" % output)

    if args.emit_map:
        with open(args.emit_map, "w", encoding="utf-8") as handle:
            json.dump(list(mapping.values()), handle, indent=2, ensure_ascii=False)
        print("wrote %s (%d unique entities) - this file re-identifies the data"
              % (args.emit_map, len(mapping)))
    if args.summary:
        with open(args.summary, "w", encoding="utf-8") as handle:
            json.dump({
                "input": os.path.basename(args.input),
                "output": os.path.basename(output),
                "seed": args.seed,
                "gazetteer": not args.no_gazetteer,
                "spans_by_type": dict(by_type),
                "unique_entities": len(mapping),
                "elements_rewritten": touched,
            }, handle, indent=2)
        print("wrote %s" % args.summary)
    return 0


def _replacements(candidates, pseudonymiser):
    """Pair every candidate with the text that replaces it.

    A name that a cell break tore apart is one entity spread over several
    candidates: it gets one pseudonym, cut into a slice per piece in the
    shape of the piece, so the cell keeps its width and no original text is
    left behind.
    """
    groups = collections.OrderedDict()
    for candidate in candidates:
        # A name that a cell break tore apart is one entity spread over
        # several candidates: it gets one pseudonym, cut into a slice per
        # piece in the shape of the piece.  The pieces are the candidates
        # that carry a joined value - from a split cell or from a name that
        # runs across a paragraph break.
        joined = candidate.entity or ""
        if candidate.source == "split-cells" or (joined and joined != candidate.text):
            key = (candidate.ptype, pseudonymiser.canonical(
                candidate.ptype, joined))
            groups.setdefault(key, []).append(candidate)
        else:
            yield candidate, pseudonymiser.fake(candidate.ptype, candidate.text)
    for key, pieces in groups.items():
        pieces.sort(key=lambda c: (c.part, c.paragraph, c.start))
        full = pseudonymiser.fake(pieces[0].ptype, pieces[0].entity)
        for piece, text in zip(pieces, _spread(full, [p.text for p in pieces])):
            yield piece, text


def _spread(value: str, shapes):
    """Cut one pseudonym into one slice per piece, matching the shapes."""
    count = len(shapes)
    if count == 1:
        return [value]
    if len(value) < count:
        return [value] + [""] * (count - 1)
    total = sum(len(shape) for shape in shapes) or 1
    slices = []
    used = 0
    for index, shape in enumerate(shapes):
        left = count - index - 1
        if left == 0:
            slices.append(value[used:])
            break
        take = int(round(len(value) * len(shape) / total))
        take = max(1, min(take, len(value) - used - left))
        slices.append(value[used:used + take])
        used += take
    return slices


def _assert_no_collision(package, entries) -> None:
    """A fake must never coincide with text that stays in the document.

    Otherwise the redaction would be ambiguous (and could be reversed by a
    reader comparing two mentions of the same entity).  Whole pseudonyms are
    compared, not the slices a split name is cut into, and a match has to
    land on a word boundary.
    """
    survivors = set()
    for document in package.text_parts():
        for paragraph in document.paragraphs:
            survivors.add(paragraph.text)
    joined = "\n".join(survivors).lower()
    for entry in entries:
        fake = entry["fake"]
        if len(fake) < 4:
            continue
        if re.search(r"(?<![0-9a-z])%s(?![0-9a-z])" % re.escape(fake.lower()),
                     joined):
            raise SystemExit(
                "refusing to write: generated value %r also occurs in text that "
                "is not being redacted" % fake)


def _default_output(path: str) -> str:
    base, extension = os.path.splitext(path)
    return base + " - REDACTED" + extension


if __name__ == "__main__":
    raise SystemExit(main())
