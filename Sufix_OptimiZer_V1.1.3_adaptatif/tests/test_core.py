import sys, unittest, tempfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from sufix_core import *
class CoreTests(unittest.TestCase):
    def test_spacing(self):
        self.assertAlmostEqual(extract_spacing_from_page('Espacement entre supports 1,5 m'),1.5)
        self.assertAlmostEqual(extract_spacing_from_page('Espacement entre support\n2 m'),2.0)
        self.assertAlmostEqual(extract_spacing_from_page('Espace entre les supports\n1,5m'),1.5)
        self.assertAlmostEqual(extract_spacing_from_page('Espace entre les supports\n0,675m'),0.675)
        self.assertAlmostEqual(extract_spacing_from_page('Espace entre les supports\n5,5m'),5.5)
    def test_cable_category(self):
        self.assertTrue(is_cable_tray_category('Chemin de câbles fil rapide'))
        self.assertTrue(is_cable_tray_category('Couvercle de chemin de câble'))
        self.assertFalse(is_cable_tray_category('Rail'))
    def test_cable_math(self):
        db={'60512060PEM':{'Category':'Chemin de câbles fil rapide','Packaging':3,'Unité de vente':'ML','Libellé':'Fil'}}
        rows=[{'Support':'Support 1','Code article':'60512060PEM','Libellé':'Fil','Nombre de supports':20,'Espacement entre supports (m)':1.5,'Quantité pour 1 support':1}]
        result,w=build_cable_tray_rows(rows,db); self.assertFalse(w); self.assertAlmostEqual(result[0]['Métré nécessaire (m)'],31.5); self.assertEqual(result[0]['Nombre de barres'],11); self.assertEqual(result[0]['Métré à chiffrer (m)'],33)
    def test_modes(self):
        for mode in ['matiere','equilibre','fabrication']:
            bins=optimize_cuts_multi([1600,1000,800],[1000,2000,3000],True,mode); self.assertTrue(bins)
    def test_csv_spacing_alias(self):
        self.assertEqual(canonicalize_header('Espacement entre support'),'Espacement entre supports (m)')
        self.assertEqual(canonicalize_header('Espace entre les supports'),'Espacement entre supports (m)')
if __name__=='__main__': unittest.main()
