import matplotlib
matplotlib.rcParams["font.family"] = "Microsoft YaHei"
matplotlib.rcParams["axes.unicode_minus"] = False

import matplotlib.pyplot as plt
from matplotlib.widgets import Button, RadioButtons
from mpl_toolkits.mplot3d.art3d import Poly3DCollection

# 主空间和物体参数
L, W, H = 3260, 1840, 2600
l2, w2, h2 = 2000, 900, 1800
x2, y2, z2 = 0, W - w2, 0
l3, w3, h3 = 500, 900, 1800
x3, y3, z3 = x2 + l2, y2, 0
l4, w4, h4 = 600, 650, 1800
x4, y4, z4 = L - l4, 0, 0

# 剩余空间主要区块（近似为最大长方体，不考虑复杂拼接）
remain_blocks = [
    # 右上角剩余区块
    {"xy": (x3+l3, y2), "l": L-(x3+l3), "w": w3, "h": H, "label": "右上剩余空间", "color": "c"},
    # 左下角底边剩余区块
    {"xy": (0, 0), "l": x4, "w": w4, "h": H, "label": "左下剩余空间", "color": "purple"},
    # 主空间底部剩余区块（去除门口区域）
    {"xy": (x4+l4, 0), "l": L-(x4+l4), "w": w4, "h": H, "label": "底边剩余空间", "color": "lime"},
    # 可根据需求增加更多区块
]

def draw_2d(ax, mode='custom'):
    ax.clear()
    ax.set_aspect('equal')
    ax.set_xlim(0, L)
    ax.set_ylim(0, W)
    ax.set_title("俯视图（长-宽）")
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)

    if mode == 'custom':
        rects = [
            {"xy": (0, 0), "l": L, "w": W, "ec": "b", "label": "主空间"},
            {"xy": (x2, y2), "l": l2, "w": w2, "ec": "r", "label": "loft bed"},
            {"xy": (x3, y3), "l": l3, "w": w3, "ec": "g", "label": "loft bed 的斜梯"},
            {"xy": (x4, y4), "l": l4, "w": w4, "ec": "brown", "label": "门口"},
        ]
        handles = []
        labels = []
        for r in rects:
            patch = plt.Rectangle(r["xy"], r["l"], r["w"], fill=None, edgecolor=r["ec"], lw=2)
            ax.add_patch(patch)
            handles.append(patch)
            labels.append(r["label"])
        def annotate_rect(x, y, l, w, color):
            ax.text(x + l/2, y + w, f"{l} mm", color=color, va="bottom", ha="center", fontsize=10, fontweight="bold")
            ax.text(x + l, y + w/2, f"{w} mm", color=color, va="center", ha="left", fontsize=10, fontweight="bold")
        annotate_rect(0, 0, L, W, "b")
        annotate_rect(x2, y2, l2, w2, "r")
        annotate_rect(x3, y3, l3, w3, "g")
        annotate_rect(x4, y4, l4, w4, "brown")
        ax.legend(handles, labels, loc="upper right", fontsize=10, frameon=True)
    elif mode == 'remain':
        main_patch = plt.Rectangle((0, 0), L, W, fill=None, edgecolor="b", lw=2)
        ax.add_patch(main_patch)
        handles = [main_patch]
        labels = ["主空间"]
        for block in remain_blocks:
            patch = plt.Rectangle(block["xy"], block["l"], block["w"], fill=None,
                                  edgecolor=block["color"], lw=2, linestyle="--")
            ax.add_patch(patch)
            handles.append(patch)
            labels.append(block["label"])
            cx, cy = block["xy"]
            ax.text(cx + block["l"] / 2, cy + block["w"] / 2, block["label"],
                    color=block["color"], ha="center", va="center",
                    fontsize=10, fontweight="bold")
        ax.legend(handles, labels, loc="upper right", fontsize=10, frameon=True)

def draw_3d(ax, mode='custom'):
    ax.clear()
    ax.set_title("3D示意图")
    ax.set_box_aspect([L, W, H])
    ax.set_axis_off()

    def cuboid_data(origin, size):
        ox, oy, oz = origin
        l, w, h = size
        x = [ox, ox+l]
        y = [oy, oy+w]
        z = [oz, oz+h]
        return [
            [ [x[0], y[0], z[0]], [x[1], y[0], z[0]], [x[1], y[1], z[0]], [x[0], y[1], z[0]],
              [x[0], y[0], z[1]], [x[1], y[0], z[1]], [x[1], y[1], z[1]], [x[0], y[1], z[1]] ] ]
    def plot_cube_at(ax, origin, size, color, label=None, alpha=0.15, linestyle='solid'):
        p = cuboid_data(origin, size)[0]
        verts = [
            [p[0], p[1], p[2], p[3]],
            [p[4], p[5], p[6], p[7]],
            [p[0], p[1], p[5], p[4]],
            [p[2], p[3], p[7], p[6]],
            [p[1], p[2], p[6], p[5]],
            [p[3], p[0], p[4], p[7]],
        ]
        poly = Poly3DCollection(verts, facecolors=color, linewidths=1, edgecolors=color, alpha=alpha, linestyle=linestyle)
        ax.add_collection3d(poly)
        if label:
            cx = origin[0] + size[0]/2
            cy = origin[1] + size[1]/2
            cz = origin[2] + size[2]/2
            ax.text(cx, cy, cz, label, color=color, ha="center", va="center", fontsize=10, fontweight="bold")

    if mode == 'custom':
        plot_cube_at(ax, (0,0,0), (L,W,H), "b", label=f"{L}×{W}×{H}", alpha=0.07)
        plot_cube_at(ax, (x2,y2,z2), (l2,w2,h2), "r", label=f"{l2}×{w2}×{h2}", alpha=0.18)
        plot_cube_at(ax, (x3,y3,z3), (l3,w3,h3), "g", label=f"{l3}×{w3}×{h3}", alpha=0.18)
        plot_cube_at(ax, (x4,y4,z4), (l4,w4,h4), "brown", label=f"{l4}×{w4}×{h4}", alpha=0.22)
    elif mode == 'remain':
        plot_cube_at(ax, (0,0,0), (L,W,H), "b", label=f"{L}×{W}×{H}", alpha=0.07)
        for block in remain_blocks:
            ox, oy = block["xy"]
            l, w, h = block["l"], block["w"], block["h"]
            plot_cube_at(ax, (ox, oy, 0), (l, w, h), block["color"], label=block["label"], alpha=0.28, linestyle='dashed')

    ax.set_xlim(0, L)
    ax.set_ylim(0, W)
    ax.set_zlim(0, H)

# 界面主逻辑
fig = plt.figure(figsize=(10, 5))
MAIN_REGION = [0.2, 0.15, 0.6, 0.8]
state = {"view": "custom", "dim": "2d"}
active_ax = {"ax": None}

ax_button = plt.axes([0.82, 0.01, 0.15, 0.08])
btn = Button(ax_button, "2D/3D切换", color="lightgray", hovercolor="orange")
ax_radio = plt.axes([0.02, 0.01, 0.16, 0.18])
radio = RadioButtons(ax_radio, ("定制物体", "剩余空间"), active=0)

# 颜色说明文字（放置在主视图区右侧空白区域）
color_labels = [
    ("红色：loft bed", "r"),
    ("绿色：loft bed 的斜梯", "g"),
    ("褐色：门口", "brown"),
]
_color_texts = []
for i, (txt, c) in enumerate(color_labels):
    t = fig.text(0.82, 0.35 + i*0.05, txt, color=c,
                 ha="left", va="center", fontsize=9)
    _color_texts.append(t)

def redraw():
    # 清除旧主axes
    if active_ax["ax"] is not None:
        fig.delaxes(active_ax["ax"])
        active_ax["ax"] = None

    # 新建axes并绘图
    if state["dim"] == "2d":
        ax = fig.add_axes(MAIN_REGION)
        draw_2d(ax, mode=state["view"])
        active_ax["ax"] = ax
    else:
        ax3d = fig.add_axes(MAIN_REGION, projection="3d")
        draw_3d(ax3d, mode=state["view"])
        active_ax["ax"] = ax3d
    fig.canvas.draw_idle()

def on_toggle(event):
    state["dim"] = "3d" if state["dim"] == "2d" else "2d"
    redraw()

def on_radio(label):
    state["view"] = "custom" if label == "定制物体" else "remain"
    redraw()

btn.on_clicked(on_toggle)
radio.on_clicked(on_radio)

# 首次绘制
ax = fig.add_axes(MAIN_REGION)
draw_2d(ax, mode=state["view"])
active_ax["ax"] = ax

plt.show()