import sys, unittest, tempfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from sufix_core import *
ROOT=Path(__file__).resolve().parents[1]
class WorkbookTests(unittest.TestCase):
    def test_cable_tray_workbook(self):
        db=load_article_database(ROOT/'base_article_supportage.xlsx')
        options,mapping=load_profile_database(ROOT/'sufix_profiles.xlsx')
        rows=[{'Niveau':'Niveau 1','Support':'Support 1','Code article':'60512060PEM','Libellé':'x','Quantité':20.0,'Longueur utile':None,'Longueur proposée':None,'Chute calculée':None}]
        rows,_=apply_article_database(rows,db); enriched=add_quantities(rows,{'Support 1':20},{'Support 1':1.5}); cable,w=build_cable_tray_rows(enriched,db); total=build_total_for_optimization(enriched); initial=build_initial_offer_rows(total,db,mapping,options,cable)
        kpis={'Barres initiales':0,'Barres optimisées':0,'Longueur initiale (m)':0,'Longueur optimisée (m)':0,'Gain longueur (m)':0,'Gain longueur (%)':0,'Longueur utilisée (m)':0,'Chute optimisée (m)':0,'Rendement matière (%)':100,'Plans de découpe distincts':0,'Codes à chiffrer initiaux':11,'Codes à chiffrer optimisés':11,'Gain codes à chiffrer':0}
        controls=build_control_rows(enriched,db,[],[],cable,w,[],'V1.0.9','V1.1.0')
        wb,ini,opt=create_output_workbook(enriched,total,[],[],[],[],initial,cable,controls,kpis,{},options,db,Path('x.csv'),None,True,False,'equilibre','V1.0.9','V1.1.0')
        self.assertEqual(wb['Etude Globale']['J2'].value,1.5)
        self.assertIn('Espacement : 1.5 m',wb['Nomenclature par support']['A1'].value)
        self.assertEqual(wb['Opti Chemin de câbles']['F2'].value,11)
        self.assertEqual(wb['Offre initiale']['I2'].value,11)
if __name__=='__main__':unittest.main()
