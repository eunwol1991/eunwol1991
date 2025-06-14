import matplotlib
matplotlib.rcParams["font.family"] = "Microsoft YaHei"     # 中文字体
matplotlib.rcParams["axes.unicode_minus"] = False          # 负号正常显示

import matplotlib.pyplot as plt
from matplotlib.widgets import Button, RadioButtons
from mpl_toolkits.mplot3d.art3d import Poly3DCollection


# ────────────────────── ❶ 主空间 & 物体参数 ──────────────────────
L, W, H = 3260, 1840, 2600            # 主空间 (长×宽×高)

# loft bed
l2, w2, h2 = 2000, 1200, 1800
x2, y2, z2 = 0, W - w2, 0             # 左上角

# 斜梯
l3, w3, h3 = 500, 900, 1800
x3, y3, z3 = x2 + l2, y2, 0

# 门口
l4, w4, h4 = 600, 650, 1800
x4, y4, z4 = L - l4, 0, 0

# 桌子（放在 loft bed 下方，略短）
l5, w5, h5 = 1950, 900, 750
x5, y5, z5 = 0, y2, 0                 # 与床左对齐


# ─────────────── ❷ 剩余空间近似区块（可继续补充） ───────────────
remain_blocks = [
    {"xy": (x3 + l3, y2), "l": L - (x3 + l3), "w": w3, "h": H,
     "label": "右上剩余空间", "color": "c"},

    {"xy": (0, 0),        "l": x4,            "w": w4, "h": H,
     "label": "左下剩余空间", "color": "purple"},

    {"xy": (x4 + l4, 0),  "l": L - (x4 + l4), "w": w4, "h": H,
     "label": "底边剩余空间", "color": "lime"},
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
            {"xy": (x3, y3), "l": l3, "w": w3, "ec": "g",     "label": "斜梯"},
            {"xy": (x4, y4), "l": l4, "w": w4, "ec": "brown", "label": "门口"},
            {"xy": (x5, y5), "l": l5, "w": w5, "ec": "c",     "label": "桌子"},
        ]
        handles, labels = [], []
        for r in rects:
            patch = plt.Rectangle(r["xy"], r["l"], r["w"], fill=None,
                                  edgecolor=r["ec"], lw=2)
            ax.add_patch(patch)
            handles.append(patch); labels.append(r["label"])

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
            patch = plt.Rectangle(blk["xy"], blk["l"], blk["w"],
                                  fill=None, edgecolor=blk["color"],
                                  lw=2, linestyle="--")
            ax.add_patch(patch)
            handles.append(patch); labels.append(blk["label"])

            cx, cy = blk["xy"]
            ax.text(cx + blk["l"]/2, cy + blk["w"]/2, blk["label"],
                    color=blk["color"], ha="center", va="center",
                    fontsize=9, fontweight="bold")

        ax.legend(handles, labels, loc="upper right", fontsize=10, frameon=True)


# ────────────────────── ❹ 绘制函数：3D ──────────────────────
def draw_3d(ax, mode="custom"):
    ax.clear()
    ax.set_title("3D示意图")
    ax.set_box_aspect([L, W, H])
    ax.set_axis_off()

    def cuboid(origin, size):
        ox, oy, oz = origin
        l, w, h = size
        x = [ox, ox + l]; y = [oy, oy + w]; z = [oz, oz + h]
        return [[
            [x[0], y[0], z[0]], [x[1], y[0], z[0]], [x[1], y[1], z[0]], [x[0], y[1], z[0]],
            [x[0], y[0], z[1]], [x[1], y[0], z[1]], [x[1], y[1], z[1]], [x[0], y[1], z[1]]
        ]]

    def plot(origin, size, color, label=None, alpha=0.15, dashed=False):
        p = cuboid(origin, size)[0]
        faces = [
            [p[0], p[1], p[2], p[3]], [p[4], p[5], p[6], p[7]],
            [p[0], p[1], p[5], p[4]], [p[2], p[3], p[7], p[6]],
            [p[1], p[2], p[6], p[5]], [p[3], p[0], p[4], p[7]],
        ]
        poly = Poly3DCollection(faces, facecolors=color, linewidths=1,
                                edgecolors=color, alpha=alpha,
                                linestyle="dashed" if dashed else "solid")
        ax.add_collection3d(poly)
        if label:
            cx = origin[0] + size[0]/2
            cy = origin[1] + size[1]/2
            cz = origin[2] + size[2]/2
            ax.text(cx, cy, cz, label, color=color, ha="center", va="center",
                    fontsize=9, fontweight="bold")

    if mode == "custom":
        plot((0, 0, 0),        (L, W, H),    "b",     f"{L}×{W}×{H}", 0.07)
        plot((x2, y2, z2),     (l2, w2, h2), "r",     f"{l2}×{w2}×{h2}", 0.18)
        plot((x3, y3, z3),     (l3, w3, h3), "g",     f"{l3}×{w3}×{h3}", 0.18)
        plot((x4, y4, z4),     (l4, w4, h4), "brown", f"{l4}×{w4}×{h4}", 0.22)
        plot((x5, y5, z5),     (l5, w5, h5), "c",     f"{l5}×{w5}×{h5}", 0.35)
    else:                      # mode == 'remain'
        plot((0, 0, 0),        (L, W, H),    "b",     f"{L}×{W}×{H}", 0.07)
        for blk in remain_blocks:
            ox, oy = blk["xy"]
            plot((ox, oy, 0),  (blk["l"], blk["w"], blk["h"]),
                 blk["color"], blk["label"], 0.28, dashed=True)

    ax.set_xlim(0, L); ax.set_ylim(0, W); ax.set_zlim(0, H)


# ────────────────────── ❺ 交互主界面 ──────────────────────
fig = plt.figure(figsize=(10, 5))
MAIN_REGION = [0.2, 0.15, 0.6, 0.8]          # 主绘图区 (left, bottom, width, height)
state      = {"view": "custom", "dim": "2d"} # 初始 = 2D+定制
active_ax  = {"ax": None}

# 按钮 & 单选框
ax_btn   = plt.axes([0.82, 0.01, 0.15, 0.08])
btn      = Button(ax_btn, "2D/3D切换", color="lightgray", hovercolor="orange")
ax_radio = plt.axes([0.02, 0.01, 0.16, 0.18])
radio    = RadioButtons(ax_radio, ("定制物体", "剩余空间"), active=0)

# 颜色说明
legend_texts = [
    ("红色：loft bed",        "r"),
    ("绿色：斜梯",            "g"),
    ("褐色：门口",            "brown"),
    ("青色：桌子",            "c"),
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
