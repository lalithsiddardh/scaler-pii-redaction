# Evaluation report

    source   Red Herring Prospectus.docx
             sha256 8b5c93f7642d659e64b51be9f6172c86c2825417f376ca1800ed331515e6f929
    output   Red Herring Prospectus - REDACTED.docx
             sha256 2a399f003f85f878…
    run      python3 redact.py "Red Herring Prospectus.docx" \
                 -o "Red Herring Prospectus - REDACTED.docx" \
                 --emit-map redaction_map.json --summary redaction_summary.json
    verify   python3 verify.py … --map redaction_map.json          15/15 pass
    gold     python3 build_gold.py "Red Herring Prospectus.docx" -o gold_entities.json
    score    python3 evaluate.py … --gold gold_entities.json --map redaction_map.json

Reproduce in that order; the pipeline is deterministic, so the hashes above
are stable.

## 1. The document

| | |
| --- | --- |
| parts with text | 150 (74 XML parts) |
| paragraphs, XML tree | 4 718 (4 864 including empty ones) |
| paragraphs with text | 4 181 |
| text elements rewritten | 2 605 |
| hyperlink field codes | 117 |
| external relationships in the source | none — every `mailto:`/`http:` target lives in a field code, not in a `.rels` file |
| images, headers, footers, footnotes, endnotes, comments | preserved byte for byte apart from the redacted text |

## 2. What was redacted

747 spans, 278 distinct entities.

| type | spans | distinct entities | what replaces it |
| --- | --- | --- | --- |
| COMPANY | 229 | 63 | invented company, same legal form and length |
| PERSON_NAME | 222 | 61 | invented name, same token count, initials and title |
| ADDRESS | 91 | 68 | invented address, same shape and length |
| EMAIL | 104 | 27 | reserved RFC 2606 domain, no wider than the original |
| PHONE | 36 | 18 | same digits, country code and grouping |
| URL | 57 | 33 | reserved host, same scheme and path length |
| DIN | 8 | 8 | eight digits |

Categories with nothing to find in this document: SSN, credit card, IP
address, date of birth, passport, Aadhaar.  They are listed as `N/A (0 gold)`
rather than scored.

## 3. Gold set

`build_gold.py` is a separate program that does **not** import `redactor`.  It
reads the zip with `xml.etree`, walks the paragraphs, and applies its own
regexes and its own copy of the rules: person names and companies as literal
gazzetteer word sequences, companies additionally by a legal-suffix scan,
e-mail, web, phone and DIN by plain patterns, addresses from a reviewed
inventory of 62 strings.

664 spans:

| type | gold spans | rule |
| --- | --- | --- |
| PERSON_NAME | 208 | gazetteer literal |
| COMPANY | 201 | legal-suffix scan + gazetteer literal |
| EMAIL | 104 | plain regex |
| URL | 55 | plain regex |
| ADDRESS | 55 | reviewed inventory |
| PHONE | 33 | plain regex |
| DIN | 8 | eight digits in a cell under a `DIN` column heading |

Known limits of the gold set, stated rather than hidden:

* the address inventory was **seeded by a proposal list and then reviewed by
  hand**; address recall is therefore an upper bound, and the 42 "extra"
  address spans below are the pipeline splitting an inventory entry into the
  pieces the document actually uses;
* names and companies come from the same two gazetteer files the detector
  uses, as literals over paragraph text.  That is a *different matching
  procedure* (paragraph text, not run offsets; word-boundary literals, not
  growth walks) but not a different source of names, so a name missing from
  the gazetteer is invisible to both;
* the gold builder walks the XML tree while the pipeline walks the raw
  element stream, and the two disagree about how many paragraphs a table
  cell or a text box contains.  Spans are therefore paired by *text*, not by
  paragraph number; `exact character match` is the strict number and is
  reported separately.

## 4. Detection against the gold set

Headline numbers, gazetteer-assisted run:

| metric | value | how it is computed |
| --- | --- | --- |
| **precision** | **0.881** | TP / (TP + FP) = 658 / 747 |
| **recall** | **0.991** | TP / (TP + FN) = 658 / 664 |
| **accuracy** | **0.881** | (TP + TN) / (TP + FP + FN + TN) = (658 + 46) / (658 + 89 + 6 + 46) |
| F1 @ 50 % | 0.933 | 2·TP / (2·TP + FP + FN) |
| strict character match | 555/664 (83.6 %) | span text equal, not merely overlapping |

Accuracy needs a denominator that includes correct *non*-detections, so the
gold set carries 46 **negative controls** — non-PII strings that really do
appear in the document and must survive untouched.  Without them "accuracy"
would be undefined or would silently reward a detector that reports nothing.
See section 5.

Per type, with precision and recall alongside F1:

| type | gold | reported | TP | FP | FN | precision | recall | F1 @ 50 % |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| ADDRESS | 55 | 91 | 49 | 42 | 6 | 0.538 | 0.891 | 0.671 |
| COMPANY | 201 | 229 | 201 | 28 | 0 | 0.878 | 1.000 | 0.935 |
| DIN | 8 | 8 | 8 | 0 | 0 | 1.000 | 1.000 | 1.000 |
| EMAIL | 104 | 104 | 104 | 0 | 0 | 1.000 | 1.000 | 1.000 |
| PERSON_NAME | 208 | 222 | 208 | 14 | 0 | 0.937 | 1.000 | 0.967 |
| PHONE | 33 | 36 | 33 | 3 | 0 | 0.917 | 1.000 | 0.957 |
| URL | 55 | 57 | 55 | 2 | 0 | 0.965 | 1.000 | 0.982 |
| **all** | **664** | **747** | **658** | **89** | **6** | **0.881** | **0.991** | **0.933** |

Recall is 1.000 on six of the seven types: every gold name, company, e-mail,
phone, URL and DIN is reported.  The single recall gap is ADDRESS, where the
inventory records a whole address as one span and the pipeline reports the
pieces the document actually prints.

Strict character-for-character match: **555/664 (83.6 %)**.  The 109
non-exact pairs are not misses — they are spans that overlap a gold span
without being equal to it, mostly an address printed across two paragraphs.

Coverage, which is the question that matters for privacy — how much of the
gold text was actually redacted, whether it was matched by one reported span
or ended up inside a longer one:

| type | gold | text redacted | recall |
| --- | --- | --- | --- |
| ADDRESS | 55 | 55 | 1.000 |
| COMPANY | 201 | 201 | 1.000 |
| DIN | 8 | 8 | 1.000 |
| EMAIL | 104 | 104 | 1.000 |
| PERSON_NAME | 208 | 208 | 1.000 |
| PHONE | 33 | 33 | 1.000 |
| URL | 55 | 55 | 1.000 |
| **all** | **664** | **664** | **1.000** |

Every gold span is redacted.  The six unmatched address pairs are gold lines
that were redacted as part of a longer address — a standalone `Pune – 411
001` line inside a two-line address, `Birdewadi` in a cell whose address runs
across two paragraphs — and the 89 "extra" spans are the pipeline reporting
an address in finer pieces than the inventory, or a second, correct copy of
one the inventory recorded once.

## 5. Negative controls — 46/46 correct

Accuracy is only meaningful against a class of right answers, so the gold set
pins down 46 strings that appear in the document and must **not** be redacted:
order and ticket numbers, `Bid Amount`, `Face Value`, `Price Band`, share
counts, a PAN, a date and a 4-digit code, all of which look numeric or
label-shaped enough to trip a careless detector.  The build refuses the set if
any control also appears in the keep-list of a reviewed address, so a control
can never be scored as both.

| | count |
| --- | --- |
| negative controls in the gold set | 46 |
| left untouched in the output | 46 |
| incorrectly redacted | 0 |
| **true negatives (TN)** | **46** |

`python3 evaluate.py …` re-checks this on every run by asking whether any
reported span covers the control's text, so the claim is reproducible and not
asserted here.

The brief's own counter-examples behave the same way: `Order Id: A-5561`,
`Ticket No: 98765`, `Price Band: 98-100`, `Corporate Identity Number:
U28129PN1979PLC141039` and `Red Herring Prospectus: 100 equity shares` are all
left alone, because what follows the colon in each case is not a second
capitalised name run.  These five are asserted in
`tests/test_redactor.py::RequiredTypes`.

## 6. The nine required types

Five of the nine required types occur in this document and are scored above.
Four do not — SSN, credit card, IP address and date of birth — so they are
`N/A (0 gold)` in the table in section 2 rather than being given a borrowed
number.  They are covered by unit tests instead, which is the honest form of
the claim:

| required type | status | evidence |
| --- | --- | --- |
| person name | scored, recall 1.000, precision 0.937 | 208 gold spans; also `Rashi Patil: John Doe` and `Rohan Dey: Peter Parker` — the brief's own examples — asserted in tests |
| e-mail | scored, perfect | 104 gold spans, 0 FP |
| phone | scored, recall 1.000 | 33 gold spans |
| company | scored, recall 1.000 | 201 gold spans |
| address | scored, recall 0.891 | 55 gold spans; the 6 FN are multi-paragraph addresses |
| SSN | absent from the document | `SSN 123-45-6789` → `123-45-6789`, test |
| credit card | absent from the document | `Card 4111 1111 1111 1111` → `4111 1111 1111 1111`, Luhn-checked, test |
| date of birth | absent from the document | `DOB: 14 March 1978` → `14 March 1978`, test |
| IP address | absent from the document | `Server 192.168.1.47` → `192.168.1.47`, test |

`tests/test_redactor.py::RequiredTypes::test_each_required_type_is_detected`
asserts all nine in one place, so the list above cannot drift from the code.

The date-of-birth detector is keyword-gated: it only fires on a span next to
`DOB`, `date of birth` or `born`.  A birth date printed without such a label
is not redacted, and that is stated as a limit rather than hidden.  Likewise
the PAN and Aadhaar-shaped strings in the document are deliberately kept —
they are not on the required list and no detector claims them.

## 7. Consistency and residual

    one original -> one fake: yes (278 entities, 0 split)
    distinct originals -> distinct fakes: placeholder reuse in
        {URL: 5, EMAIL: 1, COMPANY: 1, ADDRESS: 11} - by design, one reserved
        domain or address placeholder per category; 34 entities are involved
        and every one of them maps to exactly one fake
    original values surviving in the output: 4

The four survivors are `Maharashtra, India` in prose, each preceded by a
preposition or "at" — "manufacturing facilities are located in Maharashtra,
India", "… in Maharashtra, India, which exposes our operations …".  A state
and country name in running text is not an address, and redacting it would
make the sentence nonsense.  The verifier accepts such a survivor only when
the same paragraph in the *source* shows the preposition, and counts them
(`documented keeps … ADDRESS=4`) instead of hiding them.

## 8. Structure, layout and determinism — 15/15

| check | result |
| --- | --- |
| zip magic + integrity, both files | pass |
| XML well-formed | pass |
| package parts identical (none missing, none added) | pass |
| structure preserved (tables, rows, cells, paragraphs, runs) | pass |
| no original PII survives | pass, 4 documented keeps |
| e-mail fakes use reserved domains | pass |
| fakes preserve shape (digits, punctuation, case) | pass |
| one original to one fake | 278 entities, 34 share a placeholder |
| no doubled spaces introduced | 4 864 paragraphs, none |
| no line grows beyond 8 characters | 4 864 paragraphs, 9 grew, widest +8 |
| deterministic output | re-run is byte identical, `2a399f003f85f878…` |
| source file unchanged | `8b5c93f7642d659e…` |

Layout evidence, and its limit:

* paragraph text length: 4 181 non-empty paragraphs, 414 changed, 7 wider
  (maximum +8 characters, p95 0, p99 +2), 307 narrower, median -3.  A
  replacement is fitted to the original length, so a paragraph can only get
  narrower; the seven that grow are paragraphs where two short entities in
  one line share a single pseudonym placeholder;
* `qlmanage` renders source and output at the same page geometry
  (365 × 512 px thumbnail) and both packages hold 85 `<w:sectPr>` elements
  with no page-break markers;
* **no page count**: `docProps/app.xml` carries no `<Pages>` element and
  there is no `w:lastRenderedPageBreak` in the document, and neither
  LibreOffice nor Word is available in this environment.  A true page count
  could not be measured; the checks above are proxies for it, and are stated
  as proxies.

## 9. Pattern-only mode

`python3 redact.py … --no-gazetteer` drops the two gazetteer files and
detects on shape and context alone.

| type | gold | reported | TP | FP | FN | precision | recall | F1 @ 50 % | coverage recall |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| ADDRESS | 55 | 91 | 49 | 42 | 6 | 0.538 | 0.891 | 0.671 | 1.000 |
| COMPANY | 201 | 201 | 184 | 17 | 17 | 0.915 | 0.915 | 0.915 | 0.945 |
| DIN | 8 | 8 | 8 | 0 | 0 | 1.000 | 1.000 | 1.000 | 1.000 |
| EMAIL | 104 | 104 | 104 | 0 | 0 | 1.000 | 1.000 | 1.000 | 1.000 |
| PERSON_NAME | 208 | 52 | 52 | 0 | 156 | 1.000 | 0.250 | 0.400 | 0.591 |
| PHONE | 33 | 36 | 33 | 3 | 0 | 0.917 | 1.000 | 0.957 | 1.000 |
| URL | 55 | 57 | 55 | 2 | 0 | 0.965 | 1.000 | 0.982 | 1.000 |
| **all** | **664** | **549** | **485** | **64** | **179** | **0.883** | **0.730** | **0.800** | **0.855** |

Headline: precision 0.883, recall 0.730, accuracy 0.686 —
(485 + 46) / (485 + 64 + 179 + 46).  The 46 negative controls still score
46/46: dropping the gazetteer does not make the tool vaguer about the strings
it already knew to keep.

The shape of the failure is the whole point of this mode.  Precision *rises*
(0.881 → 0.883) while recall collapses (0.991 → 0.730), and the collapse is
almost entirely PERSON_NAME: 52 names found, none of them wrong, 156 missed.
Without a gazetteer a bare personal name in a sentence is only recognisable
from a role anchor, and the document's names mostly sit in tables and cover
lists where the anchor is a column heading rather than a word in the same
line.  Two detectors recover part of it without a gazetteer — role/label
anchors and the colon-pair rule of section 9 — and they are what keeps
precision at 0.88 instead of the 0.671 that a bare-bigram rule would score.

Verification: **13 of 15 checks pass.**  The two failures are honest ones:
15 entities from the map are still visible in the output — 12 person names
(the promoter cover list, `KUSHAL SUBBAYYA HEGDE` in prose) and 3 companies
in all-caps headings — for the same reason.  Every other property holds:
structure, shape, reserved domains, spacing, width, determinism.

This mode is a fallback for shipping the tool without the gazetteer files.
It is not the deliverable; the gazetteer-assisted run is.

## 10. Judgement calls

* **Legal entities are PII here.**  The brief said to redact personal
  information; a promoter trust and a named counterparty identify people as
  surely as a name does, so all 229 company spans are redacted.
* **Regulators and exchanges are kept**: SEBI, NSE, BSE, the Registrar of
  Companies, stock exchange references.  The Companies Registrar address in
  the incorporation history is left in place for that reason, and the gold
  set follows the same rule.
* **A state and country in prose is kept** (section 7).
* **Hyperlink field codes are redacted in both halves** — the `mailto:`
  target and the visible text — because either one alone re-identifies.
* **The map is a deliverable, and it is sensitive.**  It is the audit trail;
  without it the redaction cannot be checked or reversed, and with it the
  document can be.  Ship it separately.
