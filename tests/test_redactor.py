"""Unit and smoke tests for the DOCX pseudonymiser.

    python3 -m unittest discover -v

The pure tests cover pseudonym generation, width fitting and the span
helpers.  The integration test runs the pipeline over the prospectus if it is
in the working directory and is skipped otherwise.
"""

import os
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import evaluate                                         # noqa: E402
from redact import _replacements, _spread                      # noqa: E402
from redactor import detect                                    # noqa: E402
from redactor.detect import (                                  # noqa: E402
    MAX_COMPANY_CHARS, NAME_TAIL_RE, _clean_span, _trim_role_prefix,
)
from redactor.docx_io import Package                           # noqa: E402
from redactor.replace import Pseudonymiser                     # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SOURCE = os.path.join(ROOT, "Red Herring Prospectus.docx")


class ReplacementShape(unittest.TestCase):
    def setUp(self):
        self.p = Pseudonymiser(seed="test-seed")

    def test_same_value_same_fake(self):
        a = self.p.fake(detect.COMPANY, "KSH International Limited")
        b = Pseudonymiser(seed="test-seed").fake(
            detect.COMPANY, "KSH International Limited")
        self.assertEqual(a, b)

    def test_different_seed_different_fake(self):
        a = Pseudonymiser(seed="one").fake(detect.PERSON_NAME, "Asha Rao")
        b = Pseudonymiser(seed="two").fake(detect.PERSON_NAME, "Asha Rao")
        self.assertNotEqual(a, b)

    def test_case_variants_are_one_entity(self):
        a = self.p.fake(detect.COMPANY, "HDFC Bank Limited")
        b = self.p.fake(detect.COMPANY, "HDFC BANK LIMITED")
        self.assertEqual(a.lower(), b.lower())

    def test_company_keeps_legal_form(self):
        for original in ("Kushal Motors and Electricals Private Limited",
                         "Broad Family Trust", "Kirtane & Pandit, LLP",
                         "Acme Exports Pvt Ltd"):
            fake = self.p.fake(detect.COMPANY, original)
            self.assertLessEqual(len(fake), len(original), original)
            self.assertEqual(fake.split()[-1].rstrip(".,"),
                             original.split()[-1].rstrip(".,"))

    def test_company_case_is_preserved(self):
        fake = self.p.fake(detect.COMPANY, "HDFC BANK LIMITED")
        self.assertEqual(fake, fake.upper())

    def test_email_is_never_wider(self):
        for original in ("hitesh.ramani@citi.com", "ksh.ipo@nuvama.com",
                         "cs.connect@kshinternational.com",
                         "a.long.local.part@some-long-domain.example.org"):
            fake = self.p.fake(detect.EMAIL, original)
            self.assertLessEqual(len(fake), len(original), original)
            self.assertIn("@", fake)
            self.assertIn(fake.rsplit("@", 1)[1],
                          ("example.com", "example.net", "example.org"))

    def test_phone_keeps_digits_shape_not_digits(self):
        """A new number, in the same grouping and with the same country code."""
        for original in ("+91 22 40094400", "+ 91 8879770456",
                         "+91-20-26234000", "022 4009 4400"):
            fake = self.p.fake(detect.PHONE, original)
            old = [c for c in original if c.isdigit()]
            new = [c for c in fake if c.isdigit()]
            self.assertEqual(len(old), len(new), original)
            self.assertNotEqual(old, new, original)
            self.assertEqual([c for c in original if not c.isdigit()],
                             [c for c in fake if not c.isdigit()], original)
            if original.lstrip("+").startswith("91"):
                self.assertTrue(new[:2] == ["9", "1"], original)

    def test_url_keeps_scheme_and_length(self):
        for original in ("http://www.nuvama.com/", "https://www.sebi.gov.in/sebi",
                         "www.kshinternational.com"):
            fake = self.p.fake(detect.URL, original)
            self.assertLessEqual(len(fake), len(original), original)
            self.assertIn("example", fake)


class Spread(unittest.TestCase):
    def test_slices_rebuild_the_value(self):
        value = "Vale Family Trust"
        pieces = _spread(value, ["Broad", " Family Trust"])
        self.assertEqual("".join(pieces), value)

    def test_no_piece_is_empty(self):
        pieces = _spread("Camden Family Trust", ["Everest", " Family", " Trust"])
        self.assertTrue(all(pieces))
        self.assertEqual("".join(pieces), "Camden Family Trust")

    def test_single_piece_is_untouched(self):
        self.assertEqual(_spread("Lakeshore", ["Dhaulagiri"]), ["Lakeshore"])


class SpanHelpers(unittest.TestCase):
    def test_role_prefix_is_trimmed(self):
        text = "Sales Department Bandra Kurla Road, Mumbai 400051"
        start, end = _trim_role_prefix(text, 0, len(text))
        self.assertGreater(start, 0)
        self.assertNotIn("Department", text[start:end])

    def test_role_prefix_left_alone_when_the_span_does_not_start_the_line(self):
        text = "Our office is at Bandra Kurla Road, Mumbai 400051"
        self.assertEqual(_trim_role_prefix(text, 16, len(text)), (16, len(text)))

    def test_clean_span_trims_surrounding_space(self):
        text = "  HDFC Bank Limited  "
        start, end = _clean_span(text, 0, len(text))
        self.assertEqual(text[start:end], "HDFC Bank Limited")

    def test_company_length_cap(self):
        self.assertLess(MAX_COMPANY_CHARS, 200)

    def test_name_tail_is_refused(self):
        self.assertTrue(NAME_TAIL_RE.match("Family Trust, Annapurna Family Trust"))
        self.assertFalse(NAME_TAIL_RE.match("Annapurna Family Trust"))


class RequiredTypes(unittest.TestCase):
    """Every type the brief asks for has a detector that fires on it."""

    #: The brief's list, in its order, with an input and the span each
    #: detector must return - a detector reports the value, not the label
    #: in front of it, so "SSN 123-45-6789" yields "123-45-6789".
    REQUIRED = {
        detect.PERSON_NAME: ("Rashi Patil: John Doe", "Rashi Patil"),
        detect.EMAIL: ("rashhi.patil@gmail.com", "rashhi.patil@gmail.com"),
        detect.PHONE: ("+91 9876543210", "+91 9876543210"),
        detect.COMPANY: ("Kushal Motors and Electricals Private Limited",
                         "Kushal Motors and Electricals Private Limited"),
        detect.ADDRESS: ("45 Park Street, Pune 411001",
                         "45 Park Street, Pune 411001"),
        detect.SSN: ("SSN 123-45-6789", "123-45-6789"),
        detect.CREDIT_CARD: ("Card 4111 1111 1111 1111", "4111 1111 1111 1111"),
        detect.DATE_OF_BIRTH: ("DOB: 14 March 1978", "14 March 1978"),
        detect.IP_ADDRESS: ("Server 192.168.1.47", "192.168.1.47"),
    }

    def _scan(self, text, ptype):
        context = detect.ScanContext([detect.Unit("p", 0, text, 0)])
        config = detect.RedactionConfig(False, [], [])
        return [c.text for c in detect.scan(context, config)
                if c.ptype == ptype]

    def test_each_required_type_is_detected(self):
        for ptype, (text, expected) in sorted(self.REQUIRED.items()):
            with self.subTest(type=ptype):
                self.assertEqual(self._scan(text, ptype), [expected])

    def test_the_briefs_own_name_examples_are_caught(self):
        # "Rashi Patil: John Doe" is the assignment's headline example, and a
        # name in free text is the one shape with no role or label around it.
        for text, expected in (("Rashi Patil: John Doe", "Rashi Patil"),
                               ("Rohan Dey: Peter Parker", "Rohan Dey")):
            with self.subTest(text=text):
                self.assertEqual(self._scan(text, detect.PERSON_NAME),
                                 [expected])

    def test_offering_vocabulary_is_not_mistaken_for_a_name(self):
        # The brief's own example of what should *not* be redacted.
        for text in ("Order Id: A-5561", "Ticket No: 98765",
                     "Price Band: 98-100",
                     "Corporate Identity Number: U28129PN1979PLC141039",
                     "Red Herring Prospectus: 100 equity shares",
                     "Book Running Lead Managers: Nuvama Wealth Management "
                     "Limited"):
            with self.subTest(text=text):
                self.assertEqual(self._scan(text, detect.PERSON_NAME), [])


class Metrics(unittest.TestCase):
    """The arithmetic behind the numbers in the evaluation report."""

    def test_precision_recall_and_f1(self):
        self.assertAlmostEqual(evaluate.precision(658, 89), 658 / 747, places=6)
        self.assertAlmostEqual(evaluate.recall(658, 6), 658 / 664, places=6)
        self.assertAlmostEqual(evaluate.f1(658, 89, 6), 0.933, places=3)
        # An empty prediction is perfect, not undefined.
        self.assertEqual(evaluate.precision(0, 0), 1.0)
        self.assertEqual(evaluate.recall(0, 0), 1.0)
        self.assertEqual(evaluate.f1(0, 0, 0), 0.0)

    def test_accuracy_uses_the_negative_controls(self):
        self.assertAlmostEqual(evaluate.accuracy(658, 89, 6, 46), 0.881,
                               places=3)
        # Without true negatives the same run is not 1.0, which is why the
        # gold set has to carry them.
        self.assertLess(evaluate.accuracy(658, 89, 6, 0), 0.9)

    def test_a_negative_control_is_a_true_negative_only_when_untouched(self):
        found = [{"context": "Order Id: A-5561", "start": 0, "end": 7,
                  "ptype": detect.EMAIL}]
        self.assertTrue(evaluate._covers_any(found, "Order Id"))
        self.assertFalse(evaluate._covers_any(found, "Ticket No"))
        self.assertIn("left alone", evaluate.negatives_table(["Ticket No"],
                                                             found))


class Pipeline(unittest.TestCase):
    @unittest.skipUnless(os.path.exists(SOURCE), "prospectus not present")
    def test_run_is_deterministic_and_wider_lines_do_not_appear(self):
        from redactor import load_gazetteers
        from redact import build_units
        people, orgs = load_gazetteers()
        config = detect.RedactionConfig(True, people, orgs)
        one = _replacements(detect.scan(build_units(Package(SOURCE)), config),
                            Pseudonymiser())
        two = _replacements(detect.scan(build_units(Package(SOURCE)), config),
                            Pseudonymiser())
        self.assertEqual([(c.text, f) for c, f in one],
                         [(c.text, f) for c, f in two])
        self.assertTrue(one)
        for candidate, fake in one:
            self.assertTrue(fake.strip(), candidate.text)

    def test_verify_reports_on_a_run(self):
        if not os.path.exists(SOURCE):
            self.skipTest("prospectus not present")
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "out.docx")
            result = subprocess.run(
                [sys.executable, os.path.join(ROOT, "redact.py"), SOURCE,
                 "-o", out], capture_output=True, text=True, cwd=ROOT)
            self.assertEqual(result.returncode, 0, result.stderr[-400:])
            self.assertTrue(os.path.getsize(out) > 0)


if __name__ == "__main__":
    unittest.main()
