import matplotlib
matplotlib.rcParams["font.family"] = "Microsoft YaHei"  # 中文字体
matplotlib.rcParams["axes.unicode_minus"] = False       # 负号正常显示

import matplotlib.pyplot as plt
from dataclasses import dataclass, asdict
from typing import Dict, Optional
from matplotlib.backend_bases import cursors
import tkinter as tk
from tkinter import simpledialog, colorchooser, filedialog
import json
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

DARK_BG = "#353535"
LIGHT_BG = "#F0F0F0"
dark_mode = True

COLOR_MAIN = "#00BFFF"
COLOR_LOFT = "#FF5252"
COLOR_DOOR = "#FFA07A"
COLOR_DESK = "#00FFC6"


@dataclass
class Object3D:
    name: str
    l: float
    w: float
    h: float
    x: float
    y: float
    z: float
    color: str


selected_name: Optional[str] = None


class RectInteractor:
    """Handle move/rotate/resize for rectangles in 2D view."""

    EDGE_TOL = 30  # px distance considered near an edge

    class Item:
        def __init__(self, patch, obj: Object3D):
            self.patch = patch
            self.obj = obj

    def __init__(self, fig):
        self.fig = fig
        self.items = []
        self.active = None
        self.cid_press = fig.canvas.mpl_connect('button_press_event', self.on_press)
        self.cid_release = fig.canvas.mpl_connect('button_release_event', self.on_release)
        self.cid_motion = fig.canvas.mpl_connect('motion_notify_event', self.on_motion)

    def clear(self):
        self.items.clear()

    def register(self, patch, obj: Object3D):
        self.items.append(self.Item(patch, obj))

    def _clamp(self, x, y, w, h):
        x = min(max(0, x), L - w)
        y = min(max(0, y), W - h)
        w = max(10, min(w, L))
        h = max(10, min(h, W))
        if x + w > L:
            x = L - w
        if y + h > W:
            y = W - h
        return x, y, w, h

    def on_press(self, event):
        if (
            state.get('dim') != '2d'
            or state.get('view') != 'custom'
            or state.get('mode') != 'edit'
        ):
            return
        if event.inaxes != active_ax.get('ax'):
            return
        for item in self.items:
            contains, _ = item.patch.contains(event)
            if contains:
                x0 = item.patch.get_x()
                y0 = item.patch.get_y()
                w0 = item.patch.get_width()
                h0 = item.patch.get_height()
                near_left = abs(event.xdata - x0) < self.EDGE_TOL
                near_right = abs(event.xdata - (x0 + w0)) < self.EDGE_TOL
                near_bottom = abs(event.ydata - y0) < self.EDGE_TOL
                near_top = abs(event.ydata - (y0 + h0)) < self.EDGE_TOL
                edges = set()
                if near_left:
                    edges.add('left')
                if near_right:
                    edges.add('right')
                if near_bottom:
                    edges.add('bottom')
                if near_top:
                    edges.add('top')
                mode = 'resize' if edges else 'move'
                self.active = {
                    'item': item,
                    'mode': mode,
                    'edges': edges,
                    'press': (event.xdata, event.ydata),
                    'orig': (x0, y0, w0, h0),
                    'moved': False,
                }
                global selected_name
                selected_name = item.obj.name
                prop_panel.highlight(selected_name)
                if mode == 'move':
                    self.fig.canvas.set_cursor(cursors.MOVE)
                else:
                    if (
                        ('left' in edges or 'right' in edges)
                        and ('top' in edges or 'bottom' in edges)
                    ):
                        cur = cursors.MOVE
                    elif 'left' in edges or 'right' in edges:
                        cur = cursors.RESIZE_HORIZONTAL
                    else:
                        cur = cursors.RESIZE_VERTICAL
                    self.fig.canvas.set_cursor(cur)
                break

    def on_motion(self, event):
        if event.inaxes != active_ax.get('ax') or state.get('mode') != 'edit':
            return

        if not self.active:
            for item in self.items:
                contains, _ = item.patch.contains(event)
                if contains:
                    x0 = item.patch.get_x()
                    y0 = item.patch.get_y()
                    w0 = item.patch.get_width()
                    h0 = item.patch.get_height()
                    near_left = abs(event.xdata - x0) < self.EDGE_TOL
                    near_right = abs(event.xdata - (x0 + w0)) < self.EDGE_TOL
                    near_bottom = abs(event.ydata - y0) < self.EDGE_TOL
                    near_top = abs(event.ydata - (y0 + h0)) < self.EDGE_TOL
                    if (near_left or near_right) and (near_top or near_bottom):
                        cur = cursors.MOVE
                    elif near_left or near_right:
                        cur = cursors.RESIZE_HORIZONTAL
                    elif near_top or near_bottom:
                        cur = cursors.RESIZE_VERTICAL
                    else:
                        cur = cursors.MOVE
                    self.fig.canvas.set_cursor(cur)
                    break
            else:
                self.fig.canvas.set_cursor(cursors.POINTER)
            return

        xpress, ypress = self.active['press']
        dx = event.xdata - xpress
        dy = event.ydata - ypress
        if abs(dx) > 1 or abs(dy) > 1:
            self.active['moved'] = True

        item = self.active['item']
        patch = item.patch
        ox, oy, ow, oh = self.active['orig']

        if self.active['mode'] == 'move':
            self.fig.canvas.set_cursor(cursors.MOVE)
            nx, ny = ox + dx, oy + dy
            nx, ny, _, _ = self._clamp(nx, ny, ow, oh)
            patch.set_x(nx)
            patch.set_y(ny)
        else:  # resize
            edges = self.active['edges']
            if ('left' in edges or 'right' in edges) and ('top' in edges or 'bottom' in edges):
                cur = cursors.MOVE
            elif 'left' in edges or 'right' in edges:
                cur = cursors.RESIZE_HORIZONTAL
            else:
                cur = cursors.RESIZE_VERTICAL
            self.fig.canvas.set_cursor(cur)
            nx, ny, nw, nh = ox, oy, ow, oh
            if 'left' in self.active['edges']:
                nx = ox + dx
                nw = ow - dx
            if 'right' in self.active['edges']:
                nw = ow + dx
            if 'bottom' in self.active['edges']:
                ny = oy + dy
                nh = oh - dy
            if 'top' in self.active['edges']:
                nh = oh + dy
            nx, ny, nw, nh = self._clamp(nx, ny, nw, nh)
            patch.set_x(nx)
            patch.set_y(ny)
            patch.set_width(nw)
            patch.set_height(nh)

        self.fig.canvas.draw_idle()

    def _apply_patch(self, item):
        p = item.patch
        obj = item.obj
        obj.x = p.get_x()
        obj.y = p.get_y()
        obj.l = p.get_width()
        obj.w = p.get_height()
        update_remain_blocks()
        prop_panel.update()

    def _rotate(self, item):
        p = item.patch
        cx = p.get_x() + p.get_width() / 2
        cy = p.get_y() + p.get_height() / 2
        new_w = p.get_width()
        new_h = p.get_height()
        new_w, new_h = new_h, new_w
        nx = cx - new_w / 2
        ny = cy - new_h / 2
        nx, ny, new_w, new_h = self._clamp(nx, ny, new_w, new_h)
        p.set_x(nx)
        p.set_y(ny)
        p.set_width(new_w)
        p.set_height(new_h)
        obj = item.obj
        obj.x = nx
        obj.y = ny
        obj.l = new_w
        obj.w = new_h
        update_remain_blocks()
        prop_panel.update()
        redraw()

    def on_release(self, event):
        if not self.active:
            return
        if state.get('mode') != 'edit':
            self.active = None
            return
        item = self.active['item']
        if self.active['mode'] == 'move' and not self.active['moved']:
            self._rotate(item)
        else:
            self._apply_patch(item)
            redraw()
        global selected_name
        selected_name = item.obj.name
        prop_panel.highlight(selected_name)
        self.active = None
        self.fig.canvas.set_cursor(cursors.POINTER)


def update_remain_blocks():
    """Recalculate remain_blocks using updated coordinates."""
    door = objects.get("门口")
    desk = objects.get("桌子")
    if door:
        remain_blocks[0]["l"] = door.x
        remain_blocks[1]["xy"] = (door.x + door.l, 0)
        remain_blocks[1]["l"] = L - (door.x + door.l)
    if desk:
        remain_blocks[2]["xy"] = (0, W - desk.w)


def clamp_position(x, y, w, h):
    x = min(max(0, x), L - w)
    y = min(max(0, y), W - h)
    if x + w > L:
        x = L - w
    if y + h > W:
        y = W - h
    return x, y


class CubeInteractor:
    """Simple XY-plane move interaction in 3D view."""

    def __init__(self, fig):
        self.fig = fig
        self.active = None
        self.cid_press = fig.canvas.mpl_connect("button_press_event", self.on_press)
        self.cid_release = fig.canvas.mpl_connect("button_release_event", self.on_release)
        self.cid_motion = fig.canvas.mpl_connect("motion_notify_event", self.on_motion)

    def on_press(self, event):
        if (
            state.get("dim") != "3d"
            or state.get("view") != "custom"
            or state.get("mode") != "edit"
        ):
            return
        if event.inaxes != active_ax.get("ax"):
            return
        for obj in objects.values():
            if obj.x <= event.xdata <= obj.x + obj.l and obj.y <= event.ydata <= obj.y + obj.w:
                self.active = {
                    "obj": obj,
                    "press": (event.xdata, event.ydata),
                    "orig": (obj.x, obj.y),
                }
                global selected_name
                selected_name = obj.name
                prop_panel.highlight(selected_name)
                self.fig.canvas.set_cursor(cursors.MOVE)
                break

    def on_motion(self, event):
        if not self.active or event.inaxes != active_ax.get("ax"):
            return
        dx = event.xdata - self.active["press"][0]
        dy = event.ydata - self.active["press"][1]
        obj = self.active["obj"]
        nx, ny = clamp_position(self.active["orig"][0] + dx, self.active["orig"][1] + dy, obj.l, obj.w)
        obj.x, obj.y = nx, ny
        redraw()

    def on_release(self, event):
        if not self.active:
            return
        self.active = None
        self.fig.canvas.set_cursor(cursors.POINTER)


class PropertyPanel:
    """Show and edit object properties in real time using Tk widgets."""

    def __init__(self, parent: tk.Frame, objects: Dict[str, Object3D]):
        self.parent = parent
        self.parent.configure(bg="#F7F7F7")
        self.objects = objects
        self.vars: Dict[str, Dict[str, tk.StringVar]] = {}
        self.entries: Dict[str, Dict[str, tk.Entry]] = {}
        self.cards: Dict[str, tk.LabelFrame] = {}
        self.editable = True
        self.build()

    def build(self):
        for child in self.parent.winfo_children():
            child.destroy()
        self.vars.clear()
        self.entries.clear()
        self.cards.clear()
        for name, obj in self.objects.items():
            self._create_card(name, obj)
        self.highlight(selected_name)
        self.set_editable(self.editable)

    def _create_card(self, name: str, obj: Object3D):
        frame = tk.LabelFrame(self.parent, text=name, padx=5, pady=5, bg="#F7F7F7")
        frame.pack(fill="x", padx=5, pady=5, anchor="n")
        self.vars[name] = {}
        self.entries[name] = {}
        self.cards[name] = frame
        props = [("X", "x"), ("Y", "y"), ("长", "l"), ("宽", "w"), ("高", "h")]
        for i, (label, attr) in enumerate(props):
            r = i // 2
            c = (i % 2) * 2
            tk.Label(frame, text=label, width=4, anchor="e", bg="#F7F7F7").grid(row=r, column=c, sticky="e", padx=2, pady=2)
            var = tk.StringVar(value=f"{getattr(obj, attr):.1f}")
            ent = tk.Entry(frame, textvariable=var, width=7, bg="#f7f7f7", relief="solid", bd=1)
            ent.grid(row=r, column=c + 1, sticky="w", padx=(0, 8), pady=2)
            ent.bind("<Return>", lambda e, n=name, p=attr, v=var: self.update_prop(n, p, v.get()))
            ent.bind("<FocusOut>", lambda e, n=name, p=attr, v=var: self.update_prop(n, p, v.get()))
            self.vars[name][attr] = var
            self.entries[name][attr] = ent

    def update_prop(self, name: str, prop: str, value: str):
        if not self.editable:
            return
        try:
            val = float(value)
        except ValueError:
            self.vars[name][prop].set(f"{getattr(self.objects[name], prop):.1f}")
            return
        setattr(self.objects[name], prop, val)
        update_remain_blocks()
        redraw()

    def update(self):
        for name, obj in self.objects.items():
            if name not in self.vars:
                continue
            for prop, var in self.vars[name].items():
                var.set(f"{getattr(obj, prop):.1f}")
        self.highlight(selected_name)

    def highlight(self, selected: Optional[str]):
        for name, frame in self.cards.items():
            if name == selected:
                frame.configure(highlightbackground="orange", highlightthickness=2)
            else:
                frame.configure(highlightthickness=0)

    def set_editable(self, editable: bool):
        self.editable = editable
        state_val = "normal" if editable else "readonly"
        for entry_dict in self.entries.values():
            for ent in entry_dict.values():
                ent.configure(state=state_val)



# ────────────────────── ❶ 主空间 & 物体参数 ──────────────────────
L, W, H = 3260, 1840, 2600            # 主空间 (长×宽×高)

# 初始化物体列表
objects: Dict[str, Object3D] = {
    "loft bed": Object3D("loft bed", 2000, 1070, 576.5, L - 2000, W - 1070, 0, COLOR_LOFT),
    "门口": Object3D("门口", 905, 30, 2060, L - 905, 0, 0, COLOR_DOOR),
    "桌子": Object3D("桌子", 1260, 500, 690, 0, W - 500, 0, COLOR_DESK),
}


# ─────────────── ❷ 剩余空间近似区块（可继续补充） ───────────────
door = objects.get("门口")
desk = objects.get("桌子")
remain_blocks = [
    {"xy": (0, 0),        "l": door.x if door else 0, "w": door.w if door else 0, "h": H,
     "label": "左下剩余空间", "color": "purple"},
    {"xy": (door.x + door.l, 0) if door else (0, 0),
     "l": L - (door.x + door.l) if door else L,
     "w": door.w if door else 0, "h": H,
     "label": "底边剩余空间", "color": "lime"},
    {"xy": (0, W - desk.w) if desk else (0, 0), "l": 0, "w": 0, "h": H,
     "label": "", "color": ""},  # 示例，无其他区块
]


# ────────────────────── ❸ 绘制函数：2D ──────────────────────
def draw_2d(ax, mode="custom"):
    ax.clear()
    ax.set_facecolor(DARK_BG if dark_mode else LIGHT_BG)
    ax.set_aspect("equal")
    ax.set_xlim(0, L)
    ax.set_ylim(0, W)
    ax.set_title("俯视图（长-宽）", pad=20)
    ax.set_xticks([]); ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)

    if mode == "custom":
        rects = [
            {"xy": (0, 0), "l": L, "w": W, "ec": COLOR_MAIN, "label": "主空间", "obj": None}
        ]
        for obj in objects.values():
            rects.append({
                "xy": (obj.x, obj.y),
                "l": obj.l,
                "w": obj.w,
                "ec": obj.color,
                "label": obj.name,
                "obj": obj,
            })
        handles, labels = [], []
        drag_mgr.clear()
        for r in rects:
            patch = plt.Rectangle(r["xy"], r["l"], r["w"], fill=None, edgecolor=r["ec"], lw=2)
            ax.add_patch(patch)
            handles.append(patch)
            labels.append(r["label"])
            if r["obj"] is not None:
                drag_mgr.register(patch, r["obj"])

            # 标注长宽
            cx, cy = r["xy"]
            ax.text(
                cx + r["l"] / 2,
                cy + r["w"],
                f"{r['l']:.1f} mm",
                color=r["ec"],
                va="bottom",
                ha="center",
                fontsize=9,
                fontweight="bold",
            )
            ax.text(
                cx + r["l"],
                cy + r["w"] / 2,
                f"{r['w']:.1f} mm",
                color=r["ec"],
                va="center",
                ha="left",
                fontsize=9,
                fontweight="bold",
            )

        # Legend shown in Tk control panel

    else:  # mode == 'remain'
        main_patch = plt.Rectangle((0, 0), L, W, fill=None, edgecolor=COLOR_MAIN, lw=2)
        ax.add_patch(main_patch)
        handles, labels = [main_patch], ["主空间"]

        for blk in remain_blocks:
            if blk["label"]:
                patch = plt.Rectangle(blk["xy"], blk["l"], blk["w"],
                                      fill=None, edgecolor=blk["color"],
                                      lw=2, linestyle="--")
                ax.add_patch(patch)
                handles.append(patch); labels.append(blk["label"])
                cx, cy = blk["xy"]
                ax.text(
                    cx + blk["l"] / 2,
                    cy + blk["w"] / 2,
                    blk["label"],
                    color=blk["color"],
                    ha="center",
                    va="center",
                    fontsize=9,
                    fontweight="bold",
                )
        # Legend shown in Tk control panel


def draw_3d(ax, mode="custom"):
    ax.clear()
    ax.set_facecolor(DARK_BG if dark_mode else LIGHT_BG)
    ax.set_box_aspect((L, W, H))
    ax.set_xticks([])
    ax.set_yticks([])
    ax.set_zticks([])

    def plot(origin, size, color, label="", offset=0.05, dashed=False):
        ox, oy, oz = origin
        l, w, h = size
        verts = [
            (ox, oy, oz),
            (ox + l, oy, oz),
            (ox + l, oy + w, oz),
            (ox, oy + w, oz),
            (ox, oy, oz + h),
            (ox + l, oy, oz + h),
            (ox + l, oy + w, oz + h),
            (ox, oy + w, oz + h),
        ]
        edges = [
            (0, 1),
            (1, 2),
            (2, 3),
            (3, 0),
            (4, 5),
            (5, 6),
            (6, 7),
            (7, 4),
            (0, 4),
            (1, 5),
            (2, 6),
            (3, 7),
        ]
        for s, e in edges:
            xs = [verts[s][0], verts[e][0]]
            ys = [verts[s][1], verts[e][1]]
            zs = [verts[s][2], verts[e][2]]
            ax.plot(
                xs,
                ys,
                zs,
                color=color,
                linestyle="--" if dashed else "-",
            )
        if label:
            ax.text(
                ox + l / 2,
                oy + w / 2,
                oz + h + offset * H,
                label,
                color=color,
                ha="center",
                va="bottom",
                fontsize=9,
                fontweight="bold",
            )

    if mode == "custom":
        plot((0, 0, 0), (L, W, H), COLOR_MAIN, f"{L:.1f}×{W:.1f}×{H:.1f}", 0.07)
        for obj in objects.values():
            plot(
                (obj.x, obj.y, obj.z),
                (obj.l, obj.w, obj.h),
                obj.color,
                f"{obj.l:.1f}×{obj.w:.1f}×{obj.h:.1f}",
                0.18,
            )
    else:  # mode == 'remain'
        plot((0, 0, 0), (L, W, H), COLOR_MAIN, f"{L:.1f}×{W:.1f}×{H:.1f}", 0.07)
        for blk in remain_blocks:
            if blk["label"]:
                ox, oy = blk["xy"]
                plot(
                    (ox, oy, 0),
                    (blk["l"], blk["w"], blk["h"]),
                    blk["color"],
                    blk["label"],
                    0.28,
                    dashed=True,
                )

    ax.set_xlim(0, L)
    ax.set_ylim(0, W)
    ax.set_zlim(0, H)

# ────────────────────── ❺ 交互主界面 ──────────────────────
root = tk.Tk()
root.title("Precise CAD")
root.geometry("1200x600")
root.configure(bg=DARK_BG if dark_mode else LIGHT_BG)

frame_left = tk.Frame(root, bg=DARK_BG if dark_mode else LIGHT_BG)
frame_left.pack(side="left", fill="both", expand=True)

separator_lr = tk.Frame(root, width=3, bg="#666666")
separator_lr.pack(side="left", fill="y")

frame_right = tk.Frame(root, bg=DARK_BG if dark_mode else LIGHT_BG)
frame_right.pack(side="left", fill="y")

control_frame = tk.Frame(frame_right, bg=DARK_BG if dark_mode else LIGHT_BG)
control_frame.pack(fill="x", pady=5)

separator_hr = tk.Frame(frame_right, height=3, bg="#666666")
separator_hr.pack(fill="x")

prop_canvas = tk.Canvas(frame_right, bg=DARK_BG if dark_mode else LIGHT_BG, highlightthickness=0)
scrollbar = tk.Scrollbar(frame_right, orient="vertical", command=prop_canvas.yview)
prop_canvas.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")
prop_canvas.pack(side="left", fill="both", expand=True)

prop_frame = tk.Frame(prop_canvas, bg="#F7F7F7")
prop_canvas.create_window((0, 0), window=prop_frame, anchor="nw")

def _update_scrollregion(event):
    prop_canvas.configure(scrollregion=prop_canvas.bbox("all"))

prop_frame.bind("<Configure>", _update_scrollregion)

def _on_mousewheel(event):
    prop_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

prop_canvas.bind_all("<MouseWheel>", _on_mousewheel)

fig = Figure(figsize=(8, 5))
fig.patch.set_facecolor(DARK_BG if dark_mode else LIGHT_BG)
canvas = FigureCanvasTkAgg(fig, master=frame_left)
canvas.get_tk_widget().pack(fill="both", expand=True)

MAIN_REGION = [0.1, 0.1, 0.75, 0.8]
state      = {"view": "custom", "dim": "2d", "mode": "edit"}
active_ax  = {"ax": None}
drag_mgr   = RectInteractor(fig)
cube_mgr   = CubeInteractor(fig)
prop_panel = PropertyPanel(prop_frame, objects)

btn_toggle = tk.Button(control_frame, text="切换2D/3D")
btn_toggle.pack(fill="x", pady=2)

btn_mode = tk.Button(control_frame, text="切换到查看模式")
btn_mode.pack(fill="x", pady=2)

btn_theme = tk.Button(control_frame, text="切换主题")
btn_theme.pack(fill="x", pady=2)

view_var = tk.StringVar(value="custom")
rb_custom = tk.Radiobutton(control_frame, text="定制物体", variable=view_var,
                           value="custom")
rb_remain = tk.Radiobutton(control_frame, text="剩余空间", variable=view_var,
                           value="remain")
rb_custom.pack(anchor="w")
rb_remain.pack(anchor="w")

btn_new = tk.Button(control_frame, text="新建物体")
btn_new.pack(fill="x", pady=5)

btn_export = tk.Button(control_frame, text="导出")
btn_export.pack(fill="x", pady=2)

btn_import = tk.Button(control_frame, text="导入")
btn_import.pack(fill="x", pady=2)

legend_labels = []
for txt, col in [
    ("loft bed", COLOR_LOFT),
    ("门口", COLOR_DOOR),
    ("桌子", COLOR_DESK),
]:
    lbl = tk.Label(control_frame, text=txt, fg=col)
    lbl.pack(anchor="w")
    legend_labels.append(lbl)


# ──────────────── 重绘逻辑 ────────────────
def redraw():
    if active_ax["ax"] is not None:
        fig.delaxes(active_ax["ax"])
        active_ax["ax"] = None

    if state["dim"] == "2d":
        ax = fig.add_axes(MAIN_REGION)
        draw_2d(ax, state["view"])
        active_ax["ax"] = ax
    else:
        ax = fig.add_axes(MAIN_REGION, projection="3d")
        draw_3d(ax, state["view"])
        active_ax["ax"] = ax

    prop_panel.update()
    canvas.draw_idle()

def toggle_dim():
    state["dim"] = "3d" if state["dim"] == "2d" else "2d"
    redraw()

def toggle_mode():
    state["mode"] = "view" if state["mode"] == "edit" else "edit"
    prop_panel.set_editable(state["mode"] == "edit")
    if state["mode"] == "edit":
        btn_mode.config(text="切换到查看模式")
    else:
        btn_mode.config(text="切换到编辑模式")
    redraw()

def on_view_change():
    state["view"] = view_var.get()
    redraw()

def create_object():
    name = simpledialog.askstring("名称", "物体名称:", parent=root)
    if not name:
        return
    try:
        l = float(simpledialog.askstring("尺寸", "长(mm):", parent=root))
        w = float(simpledialog.askstring("尺寸", "宽(mm):", parent=root))
        h = float(simpledialog.askstring("尺寸", "高(mm):", parent=root))
        x = float(simpledialog.askstring("位置", "X(mm):", parent=root))
        y = float(simpledialog.askstring("位置", "Y(mm):", parent=root))
    except (TypeError, ValueError):
        return
    color = colorchooser.askcolor(parent=root)[1] or "gray"
    objects[name] = Object3D(name, l, w, h, x, y, 0, color)
    global selected_name
    selected_name = name
    prop_panel.build()
    prop_panel.set_editable(state["mode"] == "edit")
    update_remain_blocks()
    redraw()

def export_scene():
    file_path = filedialog.asksaveasfilename(
        defaultextension=".json",
        filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        parent=root,
    )
    if not file_path:
        return
    data = {name: asdict(obj) for name, obj in objects.items()}
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def import_scene():
    file_path = filedialog.askopenfilename(
        filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        parent=root,
    )
    if not file_path:
        return
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return
    objects.clear()
    for name, info in data.items():
        objects[name] = Object3D(**info)
    global selected_name
    selected_name = next(iter(objects.keys()), None)
    prop_panel.build()
    prop_panel.set_editable(state["mode"] == "edit")
    update_remain_blocks()
    redraw()

def apply_theme():
    bg = DARK_BG if dark_mode else LIGHT_BG
    fg = "#FFFFFF" if dark_mode else "#000000"
    root.configure(bg=bg)
    frame_left.configure(bg=bg)
    frame_right.configure(bg=bg)
    control_frame.configure(bg=bg)
    prop_canvas.configure(bg=bg)
    fig.patch.set_facecolor(bg)
    prop_frame.configure(bg="#F7F7F7")
    separator_lr.configure(bg="#666666")
    separator_hr.configure(bg="#666666")
    for w in [btn_toggle, btn_mode, btn_theme, btn_new, btn_export, btn_import]:
        w.configure(bg=bg, fg=fg, activebackground=bg, activeforeground=fg)
    for rb in [rb_custom, rb_remain]:
        rb.configure(bg=bg, fg=fg, activebackground=bg, activeforeground=fg, selectcolor=bg)
    for lbl in legend_labels:
        lbl.configure(bg=bg)
    prop_panel.build()
    prop_panel.set_editable(state["mode"] == "edit")
    if state["mode"] == "edit":
        btn_mode.config(text="切换到查看模式")
    else:
        btn_mode.config(text="切换到编辑模式")

def toggle_theme():
    global dark_mode
    dark_mode = not dark_mode
    apply_theme()
    redraw()

btn_toggle.config(command=toggle_dim)
btn_mode.config(command=toggle_mode)
btn_theme.config(command=toggle_theme)
rb_custom.config(command=on_view_change)
rb_remain.config(command=on_view_change)
btn_new.config(command=create_object)
btn_export.config(command=export_scene)
btn_import.config(command=import_scene)

# 首次绘制
ax_init = fig.add_axes(MAIN_REGION)
draw_2d(ax_init, state["view"])
active_ax["ax"] = ax_init
apply_theme()
redraw()
canvas.draw()
root.mainloop()
