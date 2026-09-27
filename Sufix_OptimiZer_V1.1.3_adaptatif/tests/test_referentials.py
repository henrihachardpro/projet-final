import sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from sufix_core import *
ROOT=Path(__file__).resolve().parents[1]
class RefTests(unittest.TestCase):
    def test_article_db(self):
        db=load_article_database(ROOT/'base_article_supportage.xlsx'); self.assertIn('60512060PEM',db); self.assertTrue(is_cable_tray_category(db['60512060PEM']['Category']))
    def test_profiles(self):
        options,mapping=load_profile_database(ROOT/'sufix_profiles.xlsx'); self.assertTrue(options); self.assertTrue(mapping)
if __name__=='__main__': unittest.main()
