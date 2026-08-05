"""
一键 3D 取色上色（自动识别屏幕版）

首次使用：
1. 进入游戏绘画模式，把模型放在画面中。
2. 按 F6 后按住鼠标左键拖出包含模型的矩形范围。
3. 松开左键后程序只在该范围内识别模型，并生成预览图。
4. 预览图确认没抓错后，按 F8 自动扫笔上色。

热键：
- F6：左键拖拽框选模型范围并生成预览
- F8：按识别到的区域自动扫笔上色
- F9：显示当前识别结果
- ESC：退出

默认假设：
- 游戏内 3D 吸管快捷键为 Space（长按）
- 普通鼠标左键用于涂色
- 框选范围内模型尽量位于中间，模型仍是浅色/白色

安装：
    py -m pip install pyautogui pynput
"""

from __future__ import annotations

import json
import sys
import threading
import time
from collections import deque
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Tuple

from PIL import Image, ImageDraw

try:
    import pyautogui
    from pynput import keyboard, mouse
except ImportError as exc:
    print("缺少依赖，请运行：py -m pip install pyautogui pynput")
    raise SystemExit(1) from exc

Point = Tuple[int, int]
CONFIG_PATH = Path(__file__).with_name("auto_3d_paint_onekey_config.json")
PREVIEW_PATH = Path(__file__).with_name("auto_3d_paint_preview.png")


@dataclass(frozen=True)
class Pair:
    # 相对于模型中心的目标点
    target: Point
    # 相对于模型中心的背景取样点
    sample: Point


Stroke = Tuple[Point, ...]


@dataclass(frozen=True)
class PaintPlan:
    sample: Point
    strokes: Tuple[Stroke, ...]


@dataclass(frozen=True)
class DetectionResult:
    bbox: Tuple[int, int, int, int] | None
    strokes: Tuple[Stroke, ...]


# 依照截图中的平躺人形建立六个部位。
# 需要微调时，只改这里的相对坐标即可。
PAIRS: Dict[str, Pair] = {
    "head": Pair(target=(0, -88), sample=(0, -155)),
    "torso": Pair(target=(0, 0), sample=(-145, 0)),
    "left_arm": Pair(target=(-92, -36), sample=(-165, -52)),
    "right_arm": Pair(target=(92, -36), sample=(165, -52)),
    "left_leg": Pair(target=(-48, 102), sample=(-105, 165)),
    "right_leg": Pair(target=(48, 102), sample=(105, 165)),
}

PART_STROKES: Dict[str, Tuple[Stroke, ...]] = {
    "head": (
        ((-18, -96), (0, -108), (18, -96), (18, -80), (0, -70), (-18, -80), (-18, -96)),
        ((-14, -88), (14, -88)),
        ((0, -104), (0, -72)),
    ),
    "torso": (
        ((-22, -40), (22, -40)),
        ((-35, -22), (35, -22)),
        ((-42, -4), (42, -4)),
        ((-36, 16), (36, 16)),
        ((-24, 36), (24, 36)),
    ),
    "left_arm": (
        ((-132, -58), (-74, -32)),
        ((-128, -44), (-62, -20)),
        ((-122, -30), (-56, -8)),
    ),
    "right_arm": (
        ((132, -58), (74, -32)),
        ((128, -44), (62, -20)),
        ((122, -30), (56, -8)),
    ),
    "left_leg": (
        ((-62, 50), (-100, 132)),
        ((-48, 56), (-84, 140)),
        ((-34, 60), (-70, 146)),
    ),
    "right_leg": (
        ((62, 50), (100, 132)),
        ((48, 56), (84, 140)),
        ((34, 60), (70, 146)),
    ),
}

PIPETTE_KEY = "space"
START_DELAY = 1.2
MOVE_DURATION = 0.08
PIPETTE_HOLD_DELAY = 0.14
AFTER_SAMPLE_DELAY = 0.13
AFTER_PAINT_DELAY = 0.10
STROKE_DURATION = 0.16
LIGHT_THRESHOLD = 220
SCANLINE_STEP = 8
MIN_STROKE_WIDTH = 12
ROI_LEFT_RATIO = 0.24
ROI_RIGHT_RATIO = 0.96
ROI_TOP_RATIO = 0.04
ROI_BOTTOM_RATIO = 0.94

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.025

model_center: Point | None = None
current_detection: DetectionResult | None = None
selected_region: Tuple[int, int, int, int] | None = None
selection_active = False
running = True
busy = False
lock = threading.Lock()


def add(a: Point, b: Point) -> Point:
    return a[0] + b[0], a[1] + b[1]


def offset_stroke(center: Point, stroke: Stroke) -> Stroke:
    return tuple(add(center, point) for point in stroke)


def build_paint_plans(center: Point) -> Dict[str, PaintPlan]:
    return {
        name: PaintPlan(
            sample=add(center, pair.sample),
            strokes=tuple(offset_stroke(center, stroke) for stroke in PART_STROKES[name]),
        )
        for name, pair in PAIRS.items()
    }


def build_light_mask(image: Image.Image, threshold: int = LIGHT_THRESHOLD) -> Image.Image:
    gray = image.convert("L")
    threshold_table = [255 if value >= threshold else 0 for value in range(256)]
    return gray.point(threshold_table, mode="1")


def default_detection_roi(size: tuple[int, int]) -> tuple[int, int, int, int]:
    width, height = size
    return (
        int(width * ROI_LEFT_RATIO),
        int(height * ROI_TOP_RATIO),
        int(width * ROI_RIGHT_RATIO),
        int(height * ROI_BOTTOM_RATIO),
    )


def apply_roi(mask: Image.Image, roi: tuple[int, int, int, int]) -> Image.Image:
    clipped = Image.new("1", mask.size, 0)
    clipped.paste(mask.crop(roi), roi)
    return clipped


def scanline_strokes_from_mask(
    mask: Image.Image,
    step: int = SCANLINE_STEP,
    min_width: int = MIN_STROKE_WIDTH,
) -> Tuple[Stroke, ...]:
    pixels = mask.convert("1").load()
    assert pixels is not None
    width, height = mask.size
    strokes: list[Stroke] = []

    for y in range(0, height, step):
        x = 0
        while x < width:
            while x < width and not pixels[x, y]:
                x += 1
            start = x
            while x < width and pixels[x, y]:
                x += 1
            end = x - 1
            if end - start + 1 >= min_width:
                strokes.append(((start, y), (end, y)))

    return tuple(strokes)


def detect_light_model(
    image: Image.Image,
    threshold: int = LIGHT_THRESHOLD,
    scanline_step: int = SCANLINE_STEP,
    roi: tuple[int, int, int, int] | None = None,
) -> DetectionResult:
    mask = build_light_mask(image, threshold)
    mask = apply_roi(mask, roi if roi is not None else default_detection_roi(image.size))
    bbox = mask.getbbox()
    if bbox is None:
        return DetectionResult(bbox=None, strokes=())
    strokes = scanline_strokes_from_mask(mask, step=scanline_step)
    return DetectionResult(bbox=bbox, strokes=strokes)


def component_mask_near_center(mask: Image.Image) -> Image.Image:
    mask = mask.convert("1")
    pixels = mask.load()
    assert pixels is not None
    width, height = mask.size
    visited: set[Point] = set()
    center = (width / 2, height / 2)
    best_pixels: list[Point] = []
    best_score = float("inf")

    for y in range(height):
        for x in range(width):
            if not pixels[x, y] or (x, y) in visited:
                continue
            queue: deque[Point] = deque([(x, y)])
            visited.add((x, y))
            component: list[Point] = []
            while queue:
                px, py = queue.popleft()
                component.append((px, py))
                for nx, ny in ((px - 1, py), (px + 1, py), (px, py - 1), (px, py + 1)):
                    if 0 <= nx < width and 0 <= ny < height and pixels[nx, ny] and (nx, ny) not in visited:
                        visited.add((nx, ny))
                        queue.append((nx, ny))
            avg_x = sum(point[0] for point in component) / len(component)
            avg_y = sum(point[1] for point in component) / len(component)
            distance = abs(avg_x - center[0]) + abs(avg_y - center[1])
            score = distance - len(component) * 0.01
            if score < best_score:
                best_score = score
                best_pixels = component

    result = Image.new("1", mask.size, 0)
    if best_pixels:
        result_pixels = result.load()
        assert result_pixels is not None
        for point in best_pixels:
            result_pixels[point] = 1
    return result


def offset_detection(detection: DetectionResult, origin: Point) -> DetectionResult:
    if detection.bbox is None:
        bbox = None
    else:
        left, top, right, bottom = detection.bbox
        bbox = (left + origin[0], top + origin[1], right + origin[0], bottom + origin[1])
    strokes = tuple(offset_stroke(origin, stroke) for stroke in detection.strokes)
    return DetectionResult(bbox=bbox, strokes=strokes)


def detect_selected_region(
    image: Image.Image,
    origin: Point,
    threshold: int = LIGHT_THRESHOLD,
    scanline_step: int = SCANLINE_STEP,
) -> DetectionResult:
    mask = component_mask_near_center(build_light_mask(image, threshold))
    bbox = mask.getbbox()
    if bbox is None:
        return DetectionResult(bbox=None, strokes=())
    detection = DetectionResult(
        bbox=bbox,
        strokes=scanline_strokes_from_mask(mask, step=scanline_step),
    )
    return offset_detection(detection, origin)


def normalize_region(start: Point, end: Point) -> Tuple[int, int, int, int]:
    left = min(start[0], end[0])
    top = min(start[1], end[1])
    right = max(start[0], end[0])
    bottom = max(start[1], end[1])
    return left, top, right, bottom


def save_detection_preview(
    image: Image.Image,
    detection: DetectionResult,
    path: Path = PREVIEW_PATH,
) -> None:
    preview = image.convert("RGB")
    draw = ImageDraw.Draw(preview)
    if detection.bbox is not None:
        draw.rectangle(detection.bbox, outline=(255, 0, 0), width=3)
    for stroke in detection.strokes:
        draw.line(stroke, fill=(0, 255, 0), width=2)
    preview.save(path)


def capture_detection() -> DetectionResult:
    image = pyautogui.screenshot()
    detection = detect_light_model(image)
    save_detection_preview(image, detection)
    return detection


def capture_selected_region(region: Tuple[int, int, int, int]) -> DetectionResult:
    left, top, right, bottom = region
    image = pyautogui.screenshot(region=(left, top, right - left, bottom - top))
    detection = detect_selected_region(image, origin=(left, top))
    screen = pyautogui.screenshot()
    save_detection_preview(screen, detection)
    return detection


def select_region_with_mouse() -> None:
    global current_detection, selected_region, selection_active
    start: Point | None = None

    print("[F6] 请按住鼠标左键拖出包含模型的矩形范围，松开后自动识别。")

    def on_click(x: int, y: int, button: mouse.Button, pressed: bool):
        nonlocal start
        global current_detection, selected_region, selection_active
        if button != mouse.Button.left:
            return None
        if pressed:
            start = (int(x), int(y))
            print(f"[框选] 起点：{start}")
            return None
        if start is None:
            return None

        region = normalize_region(start, (int(x), int(y)))
        selected_region = region
        selection_active = False
        print(f"[框选] 范围：{region}")
        current_detection = capture_selected_region(region)
        print(f"[识别] 模型区域：{current_detection.bbox}")
        print(f"[识别] 扫笔数量：{len(current_detection.strokes)}")
        print(f"[识别] 预览图：{PREVIEW_PATH}")
        return False

    selection_active = True
    with mouse.Listener(on_click=on_click) as listener:
        listener.join()


def save_center(center: Point) -> None:
    _ = CONFIG_PATH.write_text(
        json.dumps({"model_center": list(center)}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_center() -> Point | None:
    if not CONFIG_PATH.exists():
        return None
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        x, y = data["model_center"]
        return int(x), int(y)
    except Exception:
        return None


def paint_stroke(stroke: Stroke) -> None:
    start, *rest = stroke
    pyautogui.moveTo(*start, duration=MOVE_DURATION)
    if not rest:
        pyautogui.click()
        return

    pyautogui.mouseDown()
    try:
        for point in rest:
            pyautogui.moveTo(*point, duration=STROKE_DURATION)
    finally:
        pyautogui.mouseUp()


def paint_strokes(strokes: Tuple[Stroke, ...]) -> None:
    for stroke in strokes:
        paint_stroke(stroke)
        time.sleep(0.03)


def sample_then_paint_stroke(stroke: Stroke) -> None:
    sample = stroke[0]
    pyautogui.moveTo(*sample, duration=MOVE_DURATION)
    pyautogui.keyDown(PIPETTE_KEY)
    time.sleep(PIPETTE_HOLD_DELAY)
    pyautogui.click()
    time.sleep(0.04)
    pyautogui.keyUp(PIPETTE_KEY)
    time.sleep(AFTER_SAMPLE_DELAY)
    paint_stroke(stroke)
    time.sleep(0.03)


def sample_then_paint_strokes(strokes: Tuple[Stroke, ...]) -> None:
    for stroke in strokes:
        sample_then_paint_stroke(stroke)


def sample_then_paint(plan: PaintPlan) -> None:
    # 使用游戏自带 3D 吸管取样
    sample = plan.sample
    pyautogui.moveTo(*sample, duration=MOVE_DURATION)
    pyautogui.keyDown(PIPETTE_KEY)
    time.sleep(PIPETTE_HOLD_DELAY)
    pyautogui.click()
    time.sleep(0.06)
    pyautogui.keyUp(PIPETTE_KEY)
    time.sleep(AFTER_SAMPLE_DELAY)

    # 在对应模型部位上色
    paint_strokes(plan.strokes)
    time.sleep(AFTER_PAINT_DELAY)


def run_onekey() -> None:
    global busy
    with lock:
        if busy:
            print("[跳过] 正在执行。")
            return
        if current_detection is None or not current_detection.strokes:
            print("[未识别] 先按 F6 截屏识别模型区域。")
            return
        busy = True
        detection = current_detection

    try:
        print(f"[开始] {START_DELAY:.1f} 秒后执行，请保持游戏在前台。")
        time.sleep(START_DELAY)
        print(f"  bbox={detection.bbox} strokes={len(detection.strokes)}")
        sample_then_paint_strokes(detection.strokes)
        print("[完成] 已按识别区域自动扫笔上色。")

    except pyautogui.FailSafeException:
        print("[中止] 鼠标移到左上角，已触发安全停止。")
    except Exception as exc:
        print(f"[错误] {exc}")
    finally:
        try:
            pyautogui.keyUp(PIPETTE_KEY)
        except Exception:
            pass
        with lock:
            busy = False


def print_points() -> None:
    if current_detection is None:
        print("[未识别] 按 F6 后左键拖拽框选模型范围。")
        return
    print(f"框选范围：{selected_region}")
    point_count = sum(len(stroke) for stroke in current_detection.strokes)
    print(f"识别区域：{current_detection.bbox}")
    print(f"扫笔数量：{len(current_detection.strokes)} strokes/{point_count} points")
    print(f"预览图：{PREVIEW_PATH}")


def on_press(key: keyboard.Key | keyboard.KeyCode | None) -> None:
    global current_detection, model_center, running
    if key is None:
        return

    if key == keyboard.Key.f6:
        if selection_active:
            print("[框选] 已在等待鼠标拖拽。")
            return
        threading.Thread(target=select_region_with_mouse, daemon=True).start()

    elif key == keyboard.Key.f8:
        threading.Thread(target=run_onekey, daemon=True).start()

    elif key == keyboard.Key.f9:
        print_points()

    elif key == keyboard.Key.esc:
        running = False
        print("[退出]")


def main() -> int:
    global model_center
    model_center = load_center()

    print("=" * 64)
    print("框选锁定模型区域上色")
    print("F6 左键拖拽框选+生成预览 | F8 自动扫笔上色 | F9 显示识别结果 | ESC 退出")
    print("框选时把模型放在范围中间，程序只在框内锁定最像模型的浅色区域。")
    print("紧急停止：把鼠标快速移到屏幕左上角。")
    print("=" * 64)

    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    while running:
        time.sleep(0.1)
    listener.stop()
    listener.join()
    return 0


if __name__ == "__main__":
    sys.exit(main())
