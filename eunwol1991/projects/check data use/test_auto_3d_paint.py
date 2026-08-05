from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image, ImageDraw


def load_module():
    path = Path(__file__).with_name("auto_3d_paint.py")
    spec = importlib.util.spec_from_file_location("auto_3d_paint", path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class Auto3dPaintTests(unittest.TestCase):
    def test_build_paint_plans_offsets_strokes_from_center(self):
        module = load_module()

        plans = module.build_paint_plans((1000, 500))
        torso = plans["torso"]

        self.assertEqual(torso.sample, (855, 500))
        self.assertGreaterEqual(len(torso.strokes), 5)
        self.assertEqual(torso.strokes[0][0], (978, 460))
        self.assertEqual(torso.strokes[0][-1], (1022, 460))

    def test_build_paint_plans_covers_long_limbs_with_multiple_strokes(self):
        module = load_module()

        plans = module.build_paint_plans((1000, 500))
        left_arm = plans["left_arm"]
        right_leg = plans["right_leg"]

        self.assertGreaterEqual(len(left_arm.strokes), 3)
        self.assertGreaterEqual(len(right_leg.strokes), 3)
        self.assertLess(left_arm.strokes[0][0][0], left_arm.strokes[0][-1][0])
        self.assertLess(right_leg.strokes[0][0][1], right_leg.strokes[0][-1][1])

    def test_detect_light_model_finds_bbox_in_synthetic_image(self):
        module = load_module()
        image = Image.new("RGB", (80, 60), (30, 35, 40))
        draw = ImageDraw.Draw(image)
        draw.rectangle((22, 12, 58, 42), fill=(238, 242, 246))

        detection = module.detect_light_model(image, threshold=220, scanline_step=8)

        self.assertEqual(detection.bbox, (22, 12, 59, 43))
        self.assertGreaterEqual(len(detection.strokes), 3)

    def test_detect_light_model_ignores_left_ui_by_default(self):
        module = load_module()
        image = Image.new("RGB", (100, 60), (30, 35, 40))
        draw = ImageDraw.Draw(image)
        draw.rectangle((2, 5, 18, 55), fill=(250, 250, 250))
        draw.rectangle((35, 14, 65, 44), fill=(238, 242, 246))

        detection = module.detect_light_model(image, threshold=220, scanline_step=8)

        self.assertEqual(detection.bbox, (35, 14, 66, 45))

    def test_scanline_strokes_follow_mask_width(self):
        module = load_module()
        mask = Image.new("1", (40, 30), 0)
        draw = ImageDraw.Draw(mask)
        draw.rectangle((10, 6, 30, 22), fill=1)

        strokes = module.scanline_strokes_from_mask(mask, step=5, min_width=6)

        self.assertEqual(strokes[0], ((10, 10), (30, 10)))
        self.assertEqual(strokes[-1], ((10, 20), (30, 20)))
        self.assertEqual(len(strokes), 3)

    def test_save_detection_preview_writes_overlay_image(self):
        module = load_module()
        image = Image.new("RGB", (80, 60), (30, 35, 40))
        draw = ImageDraw.Draw(image)
        draw.rectangle((22, 12, 58, 42), fill=(238, 242, 246))
        detection = module.detect_light_model(image, threshold=220, scanline_step=8)

        with tempfile.TemporaryDirectory() as temp_dir:
            preview_path = Path(temp_dir) / "preview.png"

            module.save_detection_preview(image, detection, preview_path)

            self.assertTrue(preview_path.exists())
            self.assertGreater(preview_path.stat().st_size, 0)

    def test_detect_selected_region_prefers_light_component_near_center(self):
        module = load_module()
        image = Image.new("RGB", (120, 80), (30, 35, 40))
        draw = ImageDraw.Draw(image)
        draw.rectangle((0, 0, 28, 79), fill=(242, 242, 242))
        draw.ellipse((48, 24, 82, 58), fill=(238, 242, 246))

        detection = module.detect_selected_region(image, origin=(300, 200), threshold=220, scanline_step=8)

        self.assertEqual(detection.bbox, (348, 224, 383, 259))
        self.assertTrue(all(stroke[0][0] >= 348 for stroke in detection.strokes))

    def test_offset_detection_moves_crop_coordinates_to_screen_coordinates(self):
        module = load_module()
        detection = module.DetectionResult(bbox=(10, 20, 30, 40), strokes=(((10, 20), (30, 20)),))

        shifted = module.offset_detection(detection, (300, 200))

        self.assertEqual(shifted.bbox, (310, 220, 330, 240))
        self.assertEqual(shifted.strokes, (((310, 220), (330, 220)),))


if __name__ == "__main__":
    unittest.main()
