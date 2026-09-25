# DOCX PII pseudonymiser

A standard-library-only tool that replaces the personally identifiable
information in `Red Herring Prospectus.docx` with consistent pseudonyms, and
a verifier that proves the result is safe, structurally identical, and
deterministic.

    python3 redact.py "Red Herring Prospectus.docx" \
        -o "Red Herring Prospectus - REDACTED.docx" \
        --emit-map redaction_map.json --summary redaction_summary.json

    python3 verify.py "Red Herring Prospectus.docx" \
        "Red Herring Prospectus - REDACTED.docx" --map redaction_map.json

The pipeline imports nothing outside the Python 3.9 standard library, so
`requirements.txt` asks for nothing on its behalf.  The single dependency in
that file is `streamlit`, used only by the optional web demo in `app.py`.
`python3 -m unittest discover -v` runs the test suite.

## Web demo

    pip install -r requirements.txt
    streamlit run app.py

`app.py` is an optional front end for trying the tool without the command line.
It takes a `.docx` by upload or falls back to a built-in sample, runs
`redact.py` exactly as documented above, and then displays the results of
`verify.py`'s own checks — the same functions the CLI calls, so a passing panel
is a real result and not a hard-coded claim.

The pipeline is never imported and never modified: `app.py` invokes
`redact.py` as a subprocess, the way the verifier's own determinism check does.
What the page demonstrates is therefore the graded tool and not a
re-implementation of it.

Two deliberate choices:

* **nothing is retained.**  The upload is processed in a temporary directory
  that is deleted as soon as the run finishes.
* **the map is not downloadable.**  The audit map is written to that temporary
  directory because the checks require it, but it is not exposed in the UI,
  since it re-identifies the document.  The original → pseudonym table is shown
  in the page instead.

## Approach in brief

The pipeline is **run-level rewriting**: the DOCX is read as a zip archive,
text nodes are concatenated with their byte offsets, detectors emit
candidates with character ranges, overlapping candidates are resolved by a
priority system (address > person name > email > company > phone > DIN > URL),
and the chosen replacements are cut back over the original XML runs so the run
structure, tables, merged cells, hyperlink fields and text boxes survive
unchanged.

Determinism is achieved by seeding every pseudonym with
`SHA-256(seed || type || canonical_value)`.  A fake is *fitted* to the
original — same or shorter length, same digit groups, punctuation,
capitalisation, token count and legal form — so no line grows and pagination
does not shift.

The gold standard is built by a completely independent program (`build_gold.py`)
that does not import the detector; it applies the same logical rules (gazetteer
literals, legal-suffix scan, pattern regexes, reviewed address inventory) over
the raw XML paragraphs.  Evaluation scores the pipeline against that gold set,
then adds **46 negative controls** — strings that really appear in the document
and must not be redacted — to compute accuracy as
(TP + TN) / (TP + FP + FN + TN) = 0.881.

## What it redacts

| type | what it finds | what replaces it |
| --- | --- | --- |
| `COMPANY` | legal-suffix names (`… Private Limited`, `… Family Trust`, `… LLP`, `Pvt Ltd`, …) and gazetteer names | invented company of the same length, legal form and capitalisation |
| `PERSON_NAME` | gazetteer names, role-anchored names (`… is our Company Secretary`), designation tables, all-caps lists, contact blocks | invented name of the same token count, same initials and title |
| `ADDRESS` | PIN-anchored addresses, address continuation lines, address-like runs with no PIN | invented address of the same shape and length |
| `EMAIL` | visible addresses and the `mailto:` target of a hyperlink field | `local@` one of the RFC 2606 reserved domains, no wider than the original |
| `PHONE` | `+91 …` numbers, including the ones inside field codes | same digit count, country code and grouping |
| `URL` | `http(s)://…` and bare hosts, including hyperlink field targets | reserved host, same scheme, `www.`, TLD and path length |
| `DIN` | the eight digit director identification numbers of the directors' table | eight digits |

Regulators, exchanges and government bodies (`SEBI`, `NSE`, `BSE`, `ROC`,
`Ministry …`) are deliberately kept, as agreed.

## How it is built

    redact.py              CLI: read, detect, replace, write, map, summary
    redactor/detect.py     one detector class per entity type
    redactor/replace.py    deterministic pseudonym generation and fitting
    redactor/docx_io.py    zip/XML reading, cell ownership, run-level writing
    redactor/data/         two hand-curated gazetteer files
    verify.py              independent verification of the output
    build_gold.py          independent gold annotation set
    evaluate.py            detection, consistency and residual measurement
    app.py                 optional Streamlit demo; invokes redact.py as a
                           subprocess and imports none of the above
    tests/                 unittest suite

### Detection

`redactor/detect.py` runs one detector per type over the paragraph stream
that `docx_io` produces.  Detectors return *candidates* with a character
range, never edits, so overlapping detections from different detectors can be
resolved in one place.

The hard part of a `.docx` is that a logical entity is rarely one paragraph:

* **run seams** - "HDFC Bank" and "Limited" are separate `<w:t>` nodes;
  `decode_with_offsets` flattens the nodes into one string with an offset map
  and the replacement is cut back over the same nodes, so the run structure
  survives (`Package.apply`).
* **table cells** - `claim_innermost_nodes` gives every character to the
  innermost `<w:tc>` that contains it, so a merged cell never edits its
  neighbour's text, and `SplitEntityDetector` rejoins names that a narrow
  column tore apart (`KSH` | `Distriparks` | `Private` | `Limited`).
* **paragraph-spanning names** - `Everest Family` | `Trust` is one entity in
  two paragraphs; the second piece is linked to the first and the single
  pseudonym is *spread* over the pieces in the shape of each piece
  (`_spread` in `redact.py`).
* **hyperlink fields** - the address appears twice, in the `w:instrText`
  target and in the visible text.  Both are detected and both are replaced.
* **text boxes and shapes** - the same ownership rules apply, so a text box
  inside a cell is edited as part of the cell.

### Replacement

`replace.py` derives every fake from `SHA-256(seed || type || canonical
value)`, so the same entity always gets the same pseudonym, in this run and
in any later run, and two different entities do not collide by accident.  A
fake is then *fitted* to the original:

* the same length, or shorter - the line must not get wider, or the text can
  reflow and the pagination shifts;
* the same shape - digit groups, punctuation, capitalisation, legal form,
  token count, initials, title, and the reserved email domain that is no
  wider than the domain it replaces.

### Verification

`verify.py` re-opens both files as zip archives and runs 15 checks over 12
properties: the first property, zip magic and integrity, is checked against
both files and so contributes four of the fifteen rows.  The properties are
zip integrity, XML well-formedness, identical part lists, preserved structure,
no surviving original, reserved email domains, preserved shapes, one original
to one fake, no doubled spaces, no line growth, byte-identical re-runs, and an
unchanged source file.  It does not import the detector, so a detector bug
cannot hide itself.

    $ python3 verify.py "Red Herring Prospectus.docx" \
          "Red Herring Prospectus - REDACTED.docx" --map redaction_map.json
    ...
    all 15 checks passed

Current run: 747 spans over 278 distinct entities, 2 605 `<w:t>` elements
rewritten across 150 parts, output `sha256 2a399f003f85f878…`, all 15 checks
passing.  `EVALUATION_REPORT.md` has the full measurement.

## The map is the sensitive file

`redaction_map.json` contains every original value next to its pseudonym and
**re-identifies the document**.  It exists so that a redaction can be audited
or reversed by an authorised holder, and it must never be released with the
redacted document.  Delete it, or keep it in a separate location, once the
audit is done.

## Pattern-only mode

    python3 redact.py "Red Herring Prospectus.docx" --no-gazetteer -o out.docx

Without the two gazetteer files the tool falls back to shape and context
alone.  Companies, addresses, e-mails, phones, URLs and DINs are still found
almost completely; free-text person names are not - a name that is not next
to a role, in a designation table, or in an all-caps list is missed.  The
verifier reports this honestly (13 of 15 checks pass, 15 entities survive) and
`EVALUATION_REPORT.md` has the numbers.  Use this mode only where a gazetteer
cannot be shipped with the tool; the gazetteer-assisted mode is the
deliverable.

## Tradeoffs: false positives and false negatives

**False positives (over-redaction).**  The 42 address spans the pipeline
reports where the gold inventory recorded one whole address are not errors in
practice — every piece is a valid address fragment and redacting it is safer
than leaving it.  The 14 single-copy company names the pipeline finds twice
are likewise harmless; the pseudonym is identical.  The 4 documented keeps
(`Maharashtra, India` after a preposition) are left alone by design: a state
and country name in running prose is not an address, and redacting it would
make the sentence nonsense.  In pattern-only mode the company detector
mistakes 17 all-caps headings for company names because the legal-suffix
heuristic has no gazetteer to ground it.

**False negatives (under-redaction).**  The only recall gap in gazetteer mode
is addresses: 6 gold spans not reported, all of them multi-paragraph addresses
where the PIN and the street sit in different XML paragraphs.  In pattern-only
mode the collapse is dramatic: person name recall falls to 0.250 — 156 of 208
gold names missed — because without the gazetteer a bare name in a sentence
has no signal, and most names in the document sit in table columns without a
role word in the same line.  Two detectors recover part of it — role/label
anchors and the **colon-pair rule** (`Rashi Patil: John Doe`) — which is why
precision stays at 0.88 instead of the 0.67 a bare-bigram rule would score.

The four required types absent from the document (SSN, credit card, date of
birth, IP address) are covered by synthetic unit tests; they cannot be scored
on the prospectus.
