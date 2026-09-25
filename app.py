"""Interactive demo for the DOCX PII redaction pipeline.

Wraps the audited tool in ``redact.py`` and reuses the checks in
``verify.py`` unchanged.  The pipeline itself is never imported and never
modified: it is run exactly the way the README documents it, as a
subprocess, so what this page demonstrates is the graded tool and not a
re-implementation of it.

    streamlit run app.py

Uploaded documents are processed in a temporary directory that is deleted as
soon as the run finishes.  Nothing is written to disk permanently and no
document is retained between sessions.
"""

import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile

import streamlit as st

ROOT = os.path.dirname(os.path.abspath(__file__))
REDACT = os.path.join(ROOT, "redact.py")

#: Refuse anything larger than this.  The real prospectus is 1.8 MB; this is
#: far above any plausible evaluation input and keeps a single session from
#: pinning a free-tier CPU.
MAX_UPLOAD_BYTES = 100 * 1024 * 1024
RUN_TIMEOUT_SECONDS = 900

#: The four entities the first property covers, counted as check rows.
CHECK_SOURCE_NOTE = (
    "`source file unchanged` is not run here: it asserts a SHA-256 pinned to "
    "one specific local input file, so it cannot apply to an uploaded "
    "document.  Every other check is document-agnostic."
)


# --------------------------------------------------------------- sample input
def build_sample_docx(person: str, colleague: str, trust: str) -> bytes:
    # `person` and `colleague` are gazetteer entries, supplied by the caller.
    """A small, valid .docx that exercises every detector that can fire here.

    Names come from the shipped gazetteer so the sample demonstrates
    gazetteer-assisted detection without this file hardcoding anyone's name.
    The order/ticket numbers at the end are deliberate negative controls: the
    tool must leave them alone, and showing that is the point.
    """

    def para(text: str) -> str:
        return (
            "<w:p><w:r><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>"
            % text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        )

    def cell(text: str) -> str:
        return "<w:tc>%s</w:tc>" % para(text)

    table = (
        "<w:tbl>"
        "<w:tr>%s%s</w:tr>"
        "<w:tr>%s%s</w:tr>"
        "</w:tbl>" % (
            cell("Name"), cell("DIN"),
            cell(colleague), cell("01234567"),
        )
    )

    body = "".join([
        para("PROMOTER GROUP"),
        para("Prepared by %s, who is a Director of the Company." % person.upper()),
        para("The %s holds the promoter shareholding." % trust),
        para("Registered office: 12 Kshetrajit, MIDC Industrial Area, "
             "Pune - 411 001."),
        para("Contact: ravi.desai@acmesystems.co.in, +91 98230 12345."),
        para("Further investor information is available at "
             "https://www.acmesystems.co.in/investor and may be requested "
             "from the Company Secretary."),
        table,
        para("Order Id: A-5561, Ticket No: 98765, Price Band: 98-100."),
        "<w:sectPr/>",
    ])

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"><w:body>%s</w:body></w:document>' % body
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
        'content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
        'package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships"><Relationship Id="rId1" Type="http://schemas.'
        'openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/></Relationships>'
    )

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", rels)
        archive.writestr("word/document.xml", document)
    return buffer.getvalue()


def gazetteer_head() -> tuple:
    """Two people and one organisation from the shipped gazetteers."""
    sys.path.insert(0, ROOT)
    from redactor import load_gazetteers
    people, organisations = load_gazetteers()
    return people[0], people[1], organisations[0]


# ------------------------------------------------------------------- pipeline
def run_pipeline(source: bytes, suffix: str, use_gazetteer: bool) -> dict:
    """Run redact.py over ``source``; return the output plus its metadata."""
    workdir = tempfile.mkdtemp(prefix="pii-demo-")
    try:
        src_path = os.path.join(workdir, "source%s" % suffix)
        out_path = os.path.join(workdir, "redacted%s" % suffix)
        map_path = os.path.join(workdir, "map.json")
        sum_path = os.path.join(workdir, "summary.json")
        with open(src_path, "wb") as handle:
            handle.write(source)

        argv = [sys.executable, REDACT, src_path, "-o", out_path,
                "--emit-map", map_path, "--summary", sum_path]
        if not use_gazetteer:
            argv.append("--no-gazetteer")

        started = time.time()
        completed = subprocess.run(
            argv, cwd=ROOT, capture_output=True, text=True,
            timeout=RUN_TIMEOUT_SECONDS,
        )
        elapsed = time.time() - started
        if completed.returncode != 0:
            tail = (completed.stderr or completed.stdout or "").strip()[-2500:]
            raise RuntimeError(
                "redact.py exited %d\n\n%s" % (completed.returncode, tail)
            )

        with open(out_path, "rb") as handle:
            output = handle.read()
        with open(map_path, encoding="utf-8") as handle:
            mapping = json.load(handle)
        with open(sum_path, encoding="utf-8") as handle:
            summary = json.load(handle)

        return {
            "output": output,
            "mapping": mapping,
            "summary": summary,
            "elapsed": elapsed,
            "stdout": completed.stdout,
            "src_path": src_path,
            "out_path": out_path,
            "workdir": workdir,
        }
    except BaseException:
        shutil.rmtree(workdir, ignore_errors=True)
        raise


def run_checks(result: dict, use_gazetteer: bool, determinism: bool) -> list:
    """Reuse verify.py's own check functions on an arbitrary uploaded pair.

    verify.py's ``main`` cannot be reused directly: it requires a map path and
    its last check asserts a SHA-256 of one specific prospectus file.  The
    checks themselves take plain zip objects and a mapping, so they are called
    here directly and the inapplicable one is skipped.
    """
    import verify

    report = verify.Report()
    quiet = sys.stdout
    sys.stdout = io.StringIO()
    try:
        src = verify.check_zip(result["src_path"], report)
        out = verify.check_zip(result["out_path"], report)
        verify.check_xml_wellformed(out, report)
        verify.check_part_parity(src, out, report)
        verify.check_structure(src, out, report)
        verify.check_residual(out, result["mapping"], report)
        verify.check_fake_domains(out, report)
        verify.check_shape(result["mapping"], report)
        verify.check_consistency(out, result["mapping"], report)
        verify.check_word_spacing(src, out, report)
        verify.check_line_width(src, out, report)
        if determinism:
            verify.check_determinism(
                result["src_path"], result["out_path"], report,
                use_gazetteer=use_gazetteer,
            )
    finally:
        sys.stdout = quiet
    return report.rows


# ------------------------------------------------------------------------ ui
st.set_page_config(
    page_title="PII Redaction Tool",
    page_icon="🔒",
    layout="centered",
)
st.title("DOCX PII redaction")
st.caption(
    "Run-level XML rewriting over the raw .docx zip. No document library: "
    "the package is edited in place, so tables, merged cells, hyperlink "
    "fields and text boxes survive untouched."
)

tab_demo, tab_how = st.tabs(["Redact", "How it works"])

with tab_how:
    st.markdown(
        """
The uploaded `.docx` is read as a zip archive.  Text nodes are concatenated
with their byte offsets, detectors emit candidates with character ranges,
overlaps are resolved by priority (**address > person name > email > company >
phone > DIN > URL**), and the chosen replacements are cut back over the
original `<w:t>` nodes.

Every pseudonym is derived from `SHA-256(seed || type || canonical value)`,
so the same original always becomes the same fake.  Each fake is then *fitted*
to the original — same or shorter length, same digit groups, punctuation,
capitalisation, token count and legal form — so no line grows and pagination
does not shift.

The check panel below is `verify.py` itself.  It does not import the
detector, so a detector bug cannot hide itself.
"""
    )
    st.warning(CHECK_SOURCE_NOTE, icon="ℹ️")

with tab_demo:
    use_gazetteer = st.checkbox(
        "Use the name gazetteer",
        value=True,
        help="The shipped gazetteer is why free-text person names are found. "
             "Switch it off to see pattern-only detection.",
    )
    determinism = st.checkbox(
        "Also re-run the pipeline to confirm a byte-identical output",
        value=False,
        help="Roughly doubles the run time.",
    )
    uploaded = st.file_uploader("Source document (.docx)", type=["docx"])
    use_sample = st.button("Use the built-in sample instead", width="content")

    source = None
    label = ""
    if uploaded is not None:
        if uploaded.size > MAX_UPLOAD_BYTES:
            st.error("That file is %.1f MB; the limit is %d MB."
                     % (uploaded.size / 1048576, MAX_UPLOAD_BYTES // 1048576))
        else:
            source = uploaded.getvalue()
            label = uploaded.name
    elif use_sample:
        try:
            person, colleague, trust = gazetteer_head()
            source = build_sample_docx(person, colleague, trust)
            label = "built-in sample"
        except Exception as exc:  # noqa: BLE001 - surfaced in the UI
            st.error("Could not build the sample: %s" % exc)

    if source is not None:
        st.caption("Input: **%s** (%.0f KB)" % (label, len(source) / 1024))
        run = st.button("Redact", type="primary")

        if run:
            try:
                with st.spinner("Scanning and rewriting the package…"):
                    result = run_pipeline(source, ".docx", use_gazetteer)
                try:
                    with st.spinner("Running the verifier's checks…"):
                        rows = run_checks(result, use_gazetteer, determinism)
                finally:
                    out_bytes = result["output"]
                    mapping = result["mapping"]
                    summary = result["summary"]
                    shutil.rmtree(result["workdir"], ignore_errors=True)
            except Exception as exc:  # noqa: BLE001 - surfaced in the UI
                st.error("The run failed.")
                st.code(str(exc), language="text")
            else:
                st.success("Redacted in %.1f s" % result["elapsed"])

                head, left, right = st.columns([1.1, 1, 1])
                head.metric("spans replaced",
                            f'{summary["unique_entities"]:,} entities')
                left.metric("text elements rewritten",
                            f'{summary["elements_rewritten"]:,}')
                right.metric("checks passed",
                            "%d / %d" % (sum(1 for r in rows if r[1]), len(rows)))

                st.subheader("What was replaced")
                by_type = summary["spans_by_type"]
                distinct: dict = {}
                for entry in mapping:
                    distinct.setdefault(entry["type"], set()).add(
                        entry["original"])
                st.dataframe(
                    [{"type": k, "spans": v, "distinct entities": len(distinct[k])}
                     for k, v in sorted(by_type.items(), key=lambda kv: -kv[1])],
                    use_container_width=True, hide_index=True,
                )

                st.subheader("Original → pseudonym")
                st.caption(
                    "Taken from the audit map, which re-identifies the "
                    "document and is therefore never offered as a download.  "
                    "An authorised holder of the source can use it to reverse "
                    "a redaction."
                )
                preview = []
                for entry in mapping:
                    preview.append({"type": entry["type"],
                                    "original": entry["original"],
                                    "pseudonym": entry["fake"]})
                st.dataframe(preview, use_container_width=True,
                             hide_index=True, height=320)

                st.subheader("Verification")
                failed = [r for r in rows if not r[1]]
                if failed:
                    st.error("%d check(s) failed" % len(failed))
                else:
                    st.success("All %d checks passed." % len(rows))
                st.dataframe(
                    [{"check": name, "result": "PASS" if ok else "FAIL",
                      "detail": detail} for name, ok, detail in rows],
                    use_container_width=True, hide_index=True,
                )
                st.caption(CHECK_SOURCE_NOTE)
                residual = next((r for r in rows if r[0] == "no original PII survives"),
                                None)
                if residual and "keep" in (residual[2] or "").lower():
                    st.caption(
                        "Survivors listed here are *documented keeps*: the "
                        "verifier accepts one only where the source paragraph "
                        "shows a preposition before it, which is what makes it "
                        "prose and not a leaked address.  In the prospectus "
                        "that is `Maharashtra, India`."
                    )

                st.download_button(
                    "Download the redacted .docx",
                    data=out_bytes,
                    file_name="redacted.docx",
                    mime="application/vnd.openxmlformats-officedocument."
                          "wordprocessingml.document",
                )
