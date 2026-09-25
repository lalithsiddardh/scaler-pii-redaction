#!/usr/bin/env python3
"""Score the detector against the independent gold set.

    python3 evaluate.py "Red Herring Prospectus.docx" \\
        "Red Herring Prospectus - REDACTED.docx" \\
        --gold gold_entities.json --map redaction_map.json

Three families of numbers are reported:

* **Detection** - per type, precision, recall and F1 of the spans the
  pipeline reports against the spans in the gold set, twice: once with an
  exact character match and once with a 50 % overlap ("any overlap covering
  half of either span counts").  A name that a cell break tore apart is
  credited to the gold span it covers.
* **Consistency** - one original to one fake, and the shape of the fake
  against the shape of the original.
* **Residual** - every gold value re-searched in the redacted package.

The gold set is produced by ``build_gold.py``, which does not import the
redaction package, so a mistake in the detector's rules cannot be mirrored in
the gold.  The evaluator does import the pipeline: it is the system under
test.
"""

import argparse
import collections
import difflib
import json
import re
import sys
import zipfile
from typing import Dict, List, Sequence, Tuple

sys.path.insert(0, ".")

from redactor import detect, load_gazetteers  # noqa: E402
from redactor.docx_io import Package  # noqa: E402
from redact import build_units  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
OVERLAP = 0.5


# --------------------------------------------------------------------- score
def overlap(a_start: int, a_end: int, b_start: int, b_end: int) -> float:
    """Intersection over the shorter of the two spans."""
    shared = max(0, min(a_end, b_end) - max(a_start, b_start))
    shortest = min(a_end - a_start, b_end - b_start)
    return shared / shortest if shortest else 0.0


def similarity(a: str, b: str) -> float:
    """How much of one span the other covers, 0.0 to 1.0."""
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b or b in a:
        return 0.9
    matcher = difflib.SequenceMatcher(None, a, b)
    return matcher.quick_ratio()


def match_spans(gold: Sequence[dict], found: Sequence[dict], ptype: str):
    """Pair every gold span with the reported span that covers it.

    A pair is scored by the text the two spans cover, not by paragraph
    numbers: the gold set walks the XML tree and the pipeline walks the raw
    element stream, and the two disagree about how many paragraphs a table
    cell or a text box contains.  The text of a span is the same fact either
    way, and a wrapped address is one gold line inside one longer reported
    span, which the containment case scores as a partial credit.
    """
    truth = [g for g in gold if g["type"] == ptype]
    hits = [f for f in found if f["ptype"] == ptype]
    by_part: Dict[str, List[int]] = {}
    for index, hit in enumerate(hits):
        by_part.setdefault(hit["part"], []).append(index)
    used = set()
    pairs: List[Tuple[float, int, int]] = []
    for gold_index, one in enumerate(truth):
        best = None
        for hit_index in by_part.get(one["part"], ()):
            if hit_index in used:
                continue
            two = hits[hit_index]
            share = similarity(one["text"], two["text"])
            if share >= OVERLAP and (best is None or share > best[0]):
                best = (share, hit_index)
        if best is not None:
            used.add(best[1])
            pairs.append((best[0], gold_index, best[1]))
    return truth, hits, pairs, used


def score(gold: Sequence[dict], found: Sequence[dict], ptype: str):
    """(tp, fp, fn, exact, gold count, disagreements) for one type."""
    truth, hits, pairs, used = match_spans(gold, found, ptype)
    matched_gold = {gold_index for _, gold_index, _ in pairs}
    exact = 0
    for _, gold_index, hit_index in pairs:
        one, two = truth[gold_index], hits[hit_index]
        if (one["text"] == two["text"] and one["part"] == two["part"]
                and one["context"] == two["context"]
                and one["start"] == two["start"] and one["end"] == two["end"]):
            exact += 1
    missed = [truth[i] for i in range(len(truth)) if i not in matched_gold]
    extra = [hits[i] for i in range(len(hits)) if i not in used]
    disagreements = ([{"kind": "missed", "text": m["text"]} for m in missed]
                     + [{"kind": "extra", "text": e["text"]} for e in extra])
    return (len(matched_gold), len(extra), len(missed), exact, len(truth),
            disagreements)


def f1(tp: int, fp: int, fn: int) -> float:
    if not tp:
        return 0.0
    precision = tp / (tp + fp)
    recall = tp / (tp + fn)
    return 2 * precision * recall / (precision + recall)


def precision(tp: int, fp: int) -> float:
    """Of the spans reported, the share that matched a gold span."""
    if not (tp + fp):
        return 1.0
    return tp / (tp + fp)


def recall(tp: int, fn: int) -> float:
    """Of the gold spans, the share that was reported."""
    if not (tp + fn):
        return 1.0
    return tp / (tp + fn)


def negatives_table(negatives: Sequence[str], found: Sequence[dict]) -> str:
    """Score the strings the document contains that are *not* PII.

    A true negative is a string that no reported span overlaps.  This is the
    only source of a true negative in the whole evaluation, and without it
    there is no accuracy figure to report, because a span set is open ended:
    "correct" has no complement to be measured against.
    """
    if not negatives:
        return ("no negative controls in the gold set; accuracy is not "
                "defined for this run")
    kept, touched = [], []
    for value in negatives:
        clash = [hit for hit in found if _covers_text(hit, value)]
        (touched if clash else kept).append((value, clash))
    lines = ["%-9s %s" % ("true neg", "%d of %d negative controls left alone"
                          % (len(kept), len(negatives)))]
    for value, clash in touched:
        lines.append("   redacted: %-28r %s" % (value, ", ".join(
            sorted({hit["ptype"] for hit in clash}))))
    return "\n".join(lines)


def _covers_text(hit: dict, value: str) -> bool:
    """True when a reported span overlaps ``value`` inside its paragraph."""
    start = hit["context"].find(value)
    while start >= 0:
        if hit["start"] < start + len(value) and hit["end"] > start:
            return True
        start = hit["context"].find(value, start + 1)
    return False


def _covers_any(found: Sequence[dict], value: str) -> bool:
    return any(_covers_text(hit, value) for hit in found)


def accuracy(tp: int, fp: int, fn: int, tn: int) -> float:
    """(TP + TN) / everything the evaluator was asked about.

    Defined over the union of the gold spans, the reported spans and the
    negative controls, so the denominator is a closed set and the figure can
    be compared between runs.
    """
    denominator = tp + fp + fn + tn
    if not denominator:
        return 1.0
    return (tp + tn) / denominator


def detection_table(gold, found, show=6) -> str:
    types = sorted({g["type"] for g in gold})
    lines = ["%-12s %6s %6s %6s %6s %6s %8s %8s %7s" %
             ("type", "gold", "found", "TP", "FP", "FN", "precision",
              "recall", "F1@50%")]
    total = [0, 0, 0, 0, 0, 0]
    for ptype in types:
        tp, fp, fn, exact, truth, _ = score(gold, found, ptype)
        found_n = tp + fp
        total = [total[0] + tp, total[1] + fp, total[2] + fn,
                 total[3] + exact, total[4] + truth, total[5] + found_n]
        lines.append("%-12s %6d %6d %6d %6d %6d %8.3f %8.3f %7.3f" %
                     (ptype, truth, found_n, tp, fp, fn, precision(tp, fp),
                      recall(tp, fn), f1(tp, fp, fn)))
    tp, fp, fn, exact, truth, found_n = total
    lines.append("%-12s %6d %6d %6d %6d %6d %8.3f %8.3f %7.3f" %
                 ("ALL", truth, found_n, tp, fp, fn, precision(tp, fp),
                  recall(tp, fn), f1(tp, fp, fn)))
    exactness = exact / truth if truth else 0.0
    return "\n".join(lines) + "\nexact character match: %d/%d (%.1f%%)" % (
        exact, truth, 100.0 * exactness) + "\ntotals: TP %d  FP %d  FN %d" % (
        tp, fp, fn)


def coverage_table(gold, found) -> str:
    """How much of the gold text was redacted, span by span.

    The one-to-one pairing above scores a detection against the single best
    reported span.  The privacy question is different: if the gold line "Pune
    – 411 001" was redacted as part of a longer reported address, the text is
    gone either way, so it is covered.  A gold span is covered when some
    reported span of the same type in the same part contains it or reads
    like it.
    """
    lines = ["%-12s %6s %8s %9s" % ("type", "gold", "covered", "recall")]
    covered_all = gold_all = 0
    for ptype in sorted({g["type"] for g in gold}):
        truth = [g for g in gold if g["type"] == ptype]
        hits = [f for f in found if f["ptype"] == ptype]
        by_part: Dict[str, List[str]] = {}
        for hit in hits:
            by_part.setdefault(hit["part"], []).append(hit["text"])
        covered = 0
        for one in truth:
            if any(one["text"] in text or similarity(one["text"], text) >= 0.9
                   for text in by_part.get(one["part"], ())):
                covered += 1
        covered_all += covered
        gold_all += len(truth)
        lines.append("%-12s %6d %8d %9.3f" %
                     (ptype, len(truth), covered,
                      covered / len(truth) if truth else 0.0))
    lines.append("%-12s %6d %8d %9.3f" %
                 ("ALL", gold_all, covered_all,
                  covered_all / gold_all if gold_all else 0.0))
    return "\n".join(lines)


def disagreements(gold, found, limit=12) -> str:
    rows: List[str] = []
    for ptype in sorted({g["type"] for g in gold}):
        _, _, _, _, _, items = score(gold, found, ptype)
        for item in items:
            rows.append("  %-11s %-7s %r" % (ptype, item["kind"],
                                             item["text"][:70]))
    if not rows:
        return "  (none)"
    return "\n".join(rows[:limit]) + (
        "\n  ... and %d more" % (len(rows) - limit) if len(rows) > limit else "")


# --------------------------------------------------------------- consistency
def consistency(mapping: Sequence[dict], out_path: str) -> List[str]:
    lines = []
    by_value = collections.defaultdict(set)
    by_fake = collections.defaultdict(set)
    for entry in mapping:
        by_value[entry["original"]].add(entry["fake"])
        by_fake[(entry["type"], entry["fake"])].add(entry["original"])
    split = [v for v, f in by_value.items() if len(f) > 1]
    lines.append("one original -> one fake: %s (%d entities, %d split)"
                 % ("yes" if not split else "NO", len(by_value), len(split)))
    shared = collections.Counter(
        ptype for (ptype, fake), originals in by_fake.items() if len(originals) > 1)
    lines.append("distinct originals -> distinct fakes: %s%s"
                 % ("yes" if not shared else "placeholder reuse in %s"
                    % dict(shared),
                    "" if not shared else
                    " (by design: one reserved domain per category)"))

    zf = zipfile.ZipFile(out_path)
    text = []
    for name in zf.namelist():
        if not name.endswith(".xml"):
            continue
        body = zf.read(name).decode("utf-8", "replace")
        text.append(re.sub(r"<[^>]+>", "", body))
    blob = "\n".join(text)
    survivors = []
    for entry in mapping:
        value = entry["original"]
        probe = re.compile(r"(?<![A-Za-z0-9])" +
                           re.escape(re.sub(r"\s+", " ", value).strip()) +
                           r"(?![A-Za-z0-9])", re.IGNORECASE)
        if probe.search(blob):
            survivors.append(value)
    lines.append("original values surviving in the output: %d%s"
                 % (len(survivors),
                    "" if not survivors else " e.g. %r" % survivors[0]))
    return lines


# ----------------------------------------------------------------------- main
def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("source")
    parser.add_argument("redacted")
    parser.add_argument("--gold", default="gold_entities.json")
    parser.add_argument("--map", dest="map_path", default="redaction_map.json")
    parser.add_argument("--no-gazetteer", action="store_true")
    parser.add_argument("--show", type=int, default=12)
    args = parser.parse_args()

    with open(args.gold, encoding="utf-8") as handle:
        gold = json.load(handle)
    negatives = gold.get("negatives", []) if isinstance(gold, dict) else []
    gold = gold["entities"] if isinstance(gold, dict) else gold
    with open(args.map_path, encoding="utf-8") as handle:
        mapping = json.load(handle)

    people, organisations = load_gazetteers()
    config = detect.RedactionConfig(
        use_gazetteer=not args.no_gazetteer,
        person_gazetteer=people,
        company_gazetteer=organisations,
    )
    package = Package(args.source)
    context = build_units(package)
    found = []
    for candidate in detect.scan(context, config):
        unit = next((u for u in context.units_of(candidate.part)
                     if u.index == candidate.paragraph), None)
        text = unit.text if unit else candidate.text
        found.append({"part": candidate.part, "paragraph": candidate.paragraph,
                      "start": candidate.start, "end": candidate.end,
                      "ptype": candidate.ptype, "text": text[candidate.start:
                                                             candidate.end],
                      "context": text, "source": candidate.source})

    mode = "pattern-only (--no-gazetteer)" if args.no_gazetteer \
        else "gazetteer-assisted"
    print("=" * 66)
    print("detection against the gold set - %s" % mode)
    print("=" * 66)
    print("gold spans: %d, reported spans: %d" % (len(gold), len(found)))
    print()
    print(detection_table(gold, found))
    print()
    print("negative controls (non-PII strings in the document)")
    print(negatives_table(negatives, found))
    print()
    tp = fp = fn = 0
    for ptype in sorted({g["type"] for g in gold}):
        one, two, three, _, _, _ = score(gold, found, ptype)
        tp, fp, fn = tp + one, fp + two, fn + three
    tn = sum(1 for value in negatives if not _covers_any(found, value))
    print("-" * 66)
    print("accuracy, precision, recall")
    print("-" * 66)
    print("accuracy  = (TP + TN) / (TP + FP + FN + TN)")
    print("          = (%d + %d) / (%d + %d + %d + %d) = %.3f"
          % (tp, tn, tp, fp, fn, tn, accuracy(tp, fp, fn, tn)))
    print("precision = TP / (TP + FP) = %d / %d = %.3f"
          % (tp, tp + fp, precision(tp, fp)))
    print("recall    = TP / (TP + FN) = %d / %d = %.3f"
          % (tp, tp + fn, recall(tp, fn)))
    print("F1        = %.3f" % f1(tp, fp, fn))
    print()
    print("gold text redacted (covered by any reported span)")
    print(coverage_table(gold, found))
    print()
    print("disagreements (first %d):" % args.show)
    print(disagreements(gold, found, args.show))
    print()
    print("-" * 66)
    print("consistency and residual")
    print("-" * 66)
    for line in consistency(mapping, args.redacted):
        print(line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
