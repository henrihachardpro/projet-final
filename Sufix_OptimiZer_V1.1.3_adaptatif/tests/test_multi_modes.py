import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from sufix_core import (
    build_multi_mode_comparison_rows,
    optimization_mode_label,
    optimization_mode_file_suffix,
)


class MultiModeTests(unittest.TestCase):
    def test_mode_labels(self):
        self.assertEqual(
            optimization_mode_label("matiere"),
            "Économie matière",
        )
        self.assertEqual(
            optimization_mode_file_suffix("fabrication"),
            "FABRICATION",
        )

    def test_comparison(self):
        results = {
            "matiere": {
                "kpis": {
                    "Rendement matière (%)": 95.0,
                    "Plans de découpe distincts": 8,
                    "Barres optimisées": 10,
                    "Longueur optimisée (m)": 30,
                }
            },
            "fabrication": {
                "kpis": {
                    "Rendement matière (%)": 92.0,
                    "Plans de découpe distincts": 4,
                    "Barres optimisées": 11,
                    "Longueur optimisée (m)": 33,
                }
            },
        }
        rows = build_multi_mode_comparison_rows(results)
        self.assertEqual(len(rows), 2)
        material = next(row for row in rows if row["Mode"] == "Économie matière")
        fabrication = next(row for row in rows if row["Mode"] == "Simplicité fabrication")
        self.assertIn("Meilleur rendement", material["Points forts"])
        self.assertIn("Moins de plans", fabrication["Points forts"])


if __name__ == "__main__":
    unittest.main()
