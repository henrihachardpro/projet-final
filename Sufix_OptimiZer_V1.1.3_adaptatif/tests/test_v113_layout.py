import unittest
from pathlib import Path
from sufix_core import *

class V113LayoutTests(unittest.TestCase):
    def test_print_layout_helper(self):
        wb=Workbook(); ws=wb.active; ws.title='Visible'; ws.append(['A','B','C','D','E','F']); ws.append([1,2,3,4,5,6])
        configure_print_layout(ws)
        self.assertEqual(ws.page_setup.fitToWidth, 1)
        self.assertEqual(ws.page_setup.orientation, 'landscape')
        self.assertTrue(ws.print_area)

    def test_export_wording_mapping_is_stable(self):
        self.assertEqual(optimization_mode_file_suffix('equilibre'),'EQUILIBRE')
        self.assertEqual(optimization_mode_file_suffix('fabrication'),'FABRICATION')
        self.assertEqual(optimization_mode_file_suffix('matiere'),'MATIERE')

if __name__=='__main__': unittest.main()
