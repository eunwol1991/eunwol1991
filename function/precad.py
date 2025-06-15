import matplotlib
matplotlib.rcParams["font.family"] = "Microsoft YaHei"     # 中文字体
matplotlib.rcParams["axes.unicode_minus"] = False          # 负号正常显示

import matplotlib.pyplot as plt
from matplotlib.widgets import Button, RadioButtons
from mpl_toolkits.mplot3d.art3d import Poly3DCollection


class RectInteractor:
    """Handle move/rotate/resize for rectangles in 2D view."""

    EDGE_TOL = 30  # px distance considered near an edge

    class Item:
        def __init__(self, patch, refs):
            self.patch = patch
            self.refs = refs  # names of globals for x,y,l,w

    def __init__(self, fig):
        self.fig = fig
        self.items = []
        self.active = None
        self.cid_press = fig.canvas.mpl_connect('button_press_event', self.on_press)
        self.cid_release = fig.canvas.mpl_connect('button_release_event', self.on_release)
        self.cid_motion = fig.canvas.mpl_connect('motion_notify_event', self.on_motion)

    def clear(self):
        self.items.clear()

    def register(self, patch, refs):
        self.items.append(self.Item(patch, refs))

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
        if state.get('dim') != '2d' or state.get('view') != 'custom':
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
                break

    def on_motion(self, event):
        if not self.active:
            return
        if event.inaxes != active_ax.get('ax'):
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
            nx, ny = ox + dx, oy + dy
            nx, ny, _, _ = self._clamp(nx, ny, ow, oh)
            patch.set_x(nx)
            patch.set_y(ny)
        else:  # resize
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
        refs = item.refs
        if 'x' in refs:
            globals()[refs['x']] = p.get_x()
        if 'y' in refs:
            globals()[refs['y']] = p.get_y()
        if 'l' in refs:
            globals()[refs['l']] = p.get_width()
        if 'w' in refs:
            globals()[refs['w']] = p.get_height()

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
        refs = item.refs
        if 'x' in refs:
            globals()[refs['x']] = nx
        if 'y' in refs:
            globals()[refs['y']] = ny
        if 'l' in refs:
            globals()[refs['l']] = new_w
        if 'w' in refs:
            globals()[refs['w']] = new_h
        update_remain_blocks()
        redraw()

    def on_release(self, event):
        if not self.active:
            return
        item = self.active['item']
        if self.active['mode'] == 'move' and not self.active['moved']:
            self._rotate(item)
        else:
            self._apply_patch(item)
            update_remain_blocks()
            redraw()
        self.active = None


def update_remain_blocks():
    """Recalculate remain_blocks using updated coordinates."""
    remain_blocks[0]['l'] = x4
    remain_blocks[1]['xy'] = (x4 + l4, 0)
    remain_blocks[1]['l'] = L - (x4 + l4)
    remain_blocks[2]['xy'] = (0, W - w5)



# ────────────────────── ❶ 主空间 & 物体参数 ──────────────────────
L, W, H = 3260, 1840, 2600            # 主空间 (长×宽×高)

# loft bed（移到右上角）
l2, w2, h2 = 2000, 1070, 576.5
x2, y2, z2 = L - l2, W - w2, 0        # 右上角

# 斜梯：已移除，不再使用
l3 = w3 = h3 = 0
x3 = y3 = z3 = 0

# 门口
l4, w4, h4 = 905, 30, 2060
x4, y4, z4 = L - l4, 0, 0

# 桌子（放在左上角）
l5, w5, h5 = 1260, 500, 690
x5, y5, z5 = 0, W - w5, 0            # 左上角


# ─────────────── ❷ 剩余空间近似区块（可继续补充） ───────────────
remain_blocks = [
    {"xy": (0, 0),        "l": x4,            "w": w4, "h": H,
     "label": "左下剩余空间", "color": "purple"},
    {"xy": (x4 + l4, 0),  "l": L - (x4 + l4), "w": w4, "h": H,
     "label": "底边剩余空间", "color": "lime"},
    {"xy": (0, W - w5),   "l": 0,              "w": 0,  "h": H,
     "label": "",            "color": ""},  # 示例，无其他区块
]


# ────────────────────── ❸ 绘制函数：2D ──────────────────────
def draw_2d(ax, mode="custom"):
    ax.clear()
    ax.set_aspect("equal")
    ax.set_xlim(0, L)
    ax.set_ylim(0, W)
    ax.set_title("俯视图（长-宽）")
    ax.set_xticks([]); ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)

    if mode == "custom":
        rects = [
            {"xy": (0, 0),   "l": L,  "w": W,  "ec": "b",     "label": "主空间"},
            {"xy": (x2, y2), "l": l2, "w": w2, "ec": "r",     "label": "loft bed"},
            {"xy": (x4, y4), "l": l4, "w": w4, "ec": "brown", "label": "门口"},
            {"xy": (x5, y5), "l": l5, "w": w5, "ec": "c",     "label": "桌子"},
        ]
        handles, labels = [], []
        drag_mgr.clear()
        for r in rects:
            patch = plt.Rectangle(r["xy"], r["l"], r["w"], fill=None,
                                  edgecolor=r["ec"], lw=2)
            ax.add_patch(patch)
            handles.append(patch); labels.append(r["label"])
            if r["label"] != "主空间":
                if r["label"] == "loft bed":
                    drag_mgr.register(patch, {"x": "x2", "y": "y2", "l": "l2", "w": "w2"})
                elif r["label"] == "门口":
                    drag_mgr.register(patch, {"x": "x4", "y": "y4", "l": "l4", "w": "w4"})
                elif r["label"] == "桌子":
                    drag_mgr.register(patch, {"x": "x5", "y": "y5", "l": "l5", "w": "w5"})

            # 标注长宽
            cx, cy = r["xy"]
            ax.text(cx + r["l"]/2, cy + r["w"], f"{r['l']} mm",
                    color=r["ec"], va="bottom", ha="center", fontsize=9, fontweight="bold")
            ax.text(cx + r["l"],   cy + r["w"]/2, f"{r['w']} mm",
                    color=r["ec"], va="center", ha="left", fontsize=9, fontweight="bold")

        ax.legend(handles, labels, loc="upper right", fontsize=10, frameon=True)

    else:  # mode == 'remain'
        main_patch = plt.Rectangle((0, 0), L, W, fill=None, edgecolor="b", lw=2)
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
                ax.text(cx + blk["l"]/2, cy + blk["w"]/2, blk["label"],
                        color=blk["color"], ha="center", va="center",
@@ -126,50 +305,51 @@ def draw_3d(ax, mode="custom"):
            ax.text(cx, cy, cz, label, color=color, ha="center", va="center",
                    fontsize=9, fontweight="bold")

    if mode == "custom":
        plot((0, 0, 0),        (L, W, H),    "b",     f"{L}×{W}×{H}", 0.07)
        plot((x2, y2, z2),     (l2, w2, h2), "r",     f"{l2}×{w2}×{h2}", 0.18)
        plot((x4, y4, z4),     (l4, w4, h4), "brown", f"{l4}×{w4}×{h4}", 0.22)
        plot((x5, y5, z5),     (l5, w5, h5), "c",     f"{l5}×{w5}×{h5}", 0.35)

    else:  # mode == 'remain'
        plot((0, 0, 0),        (L, W, H),    "b",     f"{L}×{W}×{H}", 0.07)
        for blk in remain_blocks:
            if blk["label"]:
                ox, oy = blk["xy"]
                plot((ox, oy, 0),  (blk["l"], blk["w"], blk["h"]),
                     blk["color"], blk["label"], 0.28, dashed=True)

    ax.set_xlim(0, L); ax.set_ylim(0, W); ax.set_zlim(0, H)


# ────────────────────── ❺ 交互主界面 ──────────────────────
fig = plt.figure(figsize=(10, 5))
MAIN_REGION = [0.2, 0.15, 0.6, 0.8]          # 主绘图区
state      = {"view": "custom", "dim": "2d"} # 初始 = 2D+定制
active_ax  = {"ax": None}
drag_mgr   = RectInteractor(fig)

# 按钮 & 单选框
ax_btn   = plt.axes([0.82, 0.01, 0.15, 0.08])
btn      = Button(ax_btn, "2D/3D切换", color="lightgray", hovercolor="orange")
ax_radio = plt.axes([0.02, 0.01, 0.16, 0.18])
radio    = RadioButtons(ax_radio, ("定制物体", "剩余空间"), active=0)

# 颜色说明
legend_texts = [
    ("红色：loft bed", "r"),
    ("褐色：门口",    "brown"),
    ("青色：桌子",    "c"),
]
for i, (txt, col) in enumerate(legend_texts):
    fig.text(0.82, 0.35 + i * 0.05, txt, color=col,
             ha="left", va="center", fontsize=9)


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

    fig.canvas.draw_idle()

btn.on_clicked(lambda event: (state.update({"dim": "3d" if state["dim"] == "2d" else "2d"}), redraw()))
radio.on_clicked(lambda label: (state.update({"view": "custom" if label == "定制物体" else "remain"}), redraw()))

# 首次绘制
ax_init = fig.add_axes(MAIN_REGION)
draw_2d(ax_init, state["view"])
active_ax["ax"] = ax_init

plt.show()
