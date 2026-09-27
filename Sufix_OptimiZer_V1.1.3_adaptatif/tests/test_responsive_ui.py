import sys
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sufix_ui import (
    WELCOME_VARIANTS,
    choose_welcome_variant,
    welcome_control_geometry,
)

class ResponsiveUITests(unittest.TestCase):
    def test_small_1366x768_like_canvas(self):
        name,w,h=choose_welcome_variant(1366,728)
        self.assertLessEqual(w,1354)
        self.assertLessEqual(h,716)
        self.assertIn(name, {'welcome_background_1080.png','welcome_background_960.png'})

    def test_large_screen_uses_larger_asset(self):
        name,w,h=choose_welcome_variant(1920,1080)
        self.assertEqual(name,'welcome_background.png')
        self.assertEqual((w,h),(1536,1024))

    def test_minimal_window_still_returns_safe_asset(self):
        name,w,h=choose_welcome_variant(800,600)
        self.assertEqual(name,'welcome_background_720.png')

    def test_all_responsive_assets_are_present(self):
        project_dir = Path(__file__).resolve().parents[1]
        for name, _width, _height in WELCOME_VARIANTS:
            self.assertTrue((project_dir / name).is_file(), name)

    def test_controls_stay_inside_every_welcome_asset(self):
        for _name, width, height in WELCOME_VARIANTS:
            for x, y, control_width, control_height in (
                welcome_control_geometry(0, 0, width, height).values()
            ):
                self.assertGreaterEqual(x, 0)
                self.assertGreaterEqual(y, 0)
                self.assertLessEqual(x + control_width, width)
                self.assertLessEqual(y + control_height, height)

if __name__=='__main__':
    unittest.main()
