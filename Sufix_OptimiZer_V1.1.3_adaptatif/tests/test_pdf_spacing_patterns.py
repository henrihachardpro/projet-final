import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sufix_core import extract_spacing_from_page


class PDFSpacingPatternTests(unittest.TestCase):
    def test_patterns_observed_in_real_sufix_reports(self):
        samples = {
            "Espace entre les supports\n1,5m": 1.5,
            "Espace entre les supports\n1m": 1.0,
            "Espace entre les supports\n2m": 2.0,
            "Espace entre les supports\n2,5m": 2.5,
            "Espace entre les supports\n3m": 3.0,
            "Espace entre les supports\n5,5m": 5.5,
            "Espace entre les supports\n2,2m": 2.2,
            "Espace entre les supports\n0,675m": 0.675,
            "Espace entre les supports\n1,3m": 1.3,
        }
        for text, expected in samples.items():
            with self.subTest(text=text):
                self.assertAlmostEqual(
                    extract_spacing_from_page(text),
                    expected,
                )

    def test_absent_spacing_returns_none(self):
        text = "NIVEAU 1\nSUPPORT15\nDépartement\n01\nLongueur du caisson\n1,5m"
        self.assertIsNone(extract_spacing_from_page(text))


if __name__ == "__main__":
    unittest.main()
