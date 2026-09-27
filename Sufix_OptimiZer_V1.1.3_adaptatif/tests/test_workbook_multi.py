import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sufix_core import *

ROOT = Path(__file__).resolve().parents[1]


class MultiWorkbookTests(unittest.TestCase):
    def test_three_mode_sheets(self):
        db = load_article_database(ROOT / "base_article_supportage.xlsx")
        options, mapping = load_profile_database(ROOT / "sufix_profiles.xlsx")

        rows = [{
            "Niveau": "Niveau 1",
            "Support": "Support 1",
            "Code article": "60512060PEM",
            "Libellé": "x",
            "Quantité": 20.0,
            "Longueur utile": None,
            "Longueur proposée": None,
            "Chute calculée": None,
        }]
        rows, _ = apply_article_database(rows, db)
        enriched = add_quantities(
            rows,
            {"Support 1": 20},
            {"Support 1": 1.5},
        )
        cable, warnings = build_cable_tray_rows(enriched, db)
        total = build_total_for_optimization(enriched)
        initial = build_initial_offer_rows(
            total,
            db,
            mapping,
            options,
            cable,
        )

        base_kpis = {
            "Barres initiales": 0,
            "Barres optimisées": 0,
            "Longueur initiale (m)": 0,
            "Longueur optimisée (m)": 0,
            "Gain longueur (m)": 0,
            "Gain longueur (%)": 0,
            "Longueur utilisée (m)": 0,
            "Chute optimisée (m)": 0,
            "Rendement matière (%)": 100,
            "Plans de découpe distincts": 0,
            "Codes à chiffrer initiaux": 11,
            "Codes à chiffrer optimisés": 11,
            "Gain codes à chiffrer": 0,
        }
        mode_results = {
            key: {
                "cut_summary": [],
                "cut_detail": [],
                "kpis": dict(base_kpis),
            }
            for key in ("matiere", "equilibre", "fabrication")
        }
        controls = build_control_rows(
            enriched,
            db,
            [],
            [],
            cable,
            warnings,
            [],
            "V1.0.9",
            "V1.1.2",
        )

        wb, initial_rows, optimized = create_output_workbook_multi(
            enriched,
            total,
            [],
            [],
            initial,
            cable,
            controls,
            mode_results,
            {},
            options,
            db,
            Path("x.csv"),
            None,
            True,
            False,
            "V1.0.9",
            "V1.1.2",
        )

        expected = {
            "Synthèse optimisation",
            "Opti globale - Matière",
            "Plan découpe - Matière",
            "Opti globale - Équilibré",
            "Plan découpe - Équilibré",
            "Opti globale - Fabrication",
            "Plan découpe - Fabrication",
        }
        self.assertTrue(expected.issubset(set(wb.sheetnames)))
        self.assertEqual(set(optimized), {"matiere", "equilibre", "fabrication"})
        self.assertEqual(
            wb["Synthèse optimisation"]["A2"].value,
            "Économie matière",
        )


if __name__ == "__main__":
    unittest.main()
