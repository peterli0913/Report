#!/usr/bin/env python3
"""生成生产管理场景常用图标的透明 PNG，只依赖 Pillow。

为什么自己画而不下载图标库：汇报材料常在没有外网的内网机器上生成，图标文件缺失
会让整页排版塌掉；而且下载来的图标授权状态往往说不清。这里的图标是几何绘制，
颜色和线宽跟着主题走，离线可用、授权无争议。

用法：
    python3 make_icon.py list                          # 列出全部图标名
    python3 make_icon.py quality out.png --color E5B620 --size 256
    python3 make_icon.py all icons/ --color F2F5F9     # 全部导出到目录

在代码里用：
    from make_icon import draw_icon, ICONS
    draw_icon("production", "icon.png", color="E5B620", size=256)
"""

import math
import os
import sys

from PIL import Image, ImageDraw

# 画布按 100x100 设计，导出时按 size 缩放；线宽 8 在 256px 下视觉舒适
BOX = 100.0


def _rgb(hexstr, alpha=255):
    h = str(hexstr).lstrip("#")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), alpha)


def _line(d, pts, c, w, joint="curve"):
    d.line([tuple(p) for p in pts], fill=c, width=w, joint=joint)


def _circle(d, cx, cy, r, c, w=None, fill=None):
    box = [cx - r, cy - r, cx + r, cy + r]
    if fill is not None:
        d.ellipse(box, fill=fill)
    if w:
        d.ellipse(box, outline=c, width=w)


# --- 各图标的绘制函数：入参 (draw, color, w=线宽) ------------------------

def _production(d, c, w):
    """产量/产能：递增柱状"""
    for i, h in enumerate((26, 44, 62)):
        x = 22 + i * 24
        d.rounded_rectangle([x, 84 - h, x + 14, 84], radius=3, fill=c)
    _line(d, [(18, 90), (86, 90)], c, w)


def _quality(d, c, w):
    """质量：盾牌 + 勾"""
    _line(d, [(50, 14), (82, 26), (82, 52), (50, 86), (18, 52), (18, 26), (50, 14)],
          c, w)
    _line(d, [(35, 50), (46, 62), (67, 38)], c, w + 1)


def _safety(d, c, w):
    """安全：安全帽（半圆帽壳 + 通长帽檐 + 顶脊）"""
    d.arc([20, 22, 80, 82], start=180, end=360, fill=c, width=w)  # 帽壳，底在 y=52
    _line(d, [(20, 52), (20, 62)], c, w)      # 两侧下沿接到帽檐
    _line(d, [(80, 52), (80, 62)], c, w)
    _line(d, [(14, 64), (86, 64)], c, w + 2)  # 帽檐，比帽壳略宽
    _line(d, [(50, 22), (50, 13)], c, w)      # 顶脊


def _equipment(d, c, w):
    """设备：齿轮"""
    r_out, r_in = 34, 20
    for i in range(8):
        a = math.radians(i * 45)
        d.line([(50 + r_in * math.cos(a), 50 + r_in * math.sin(a)),
                (50 + r_out * math.cos(a), 50 + r_out * math.sin(a))],
               fill=c, width=w + 2)
    _circle(d, 50, 50, r_in + 3, c, w)
    _circle(d, 50, 50, 8, c, w)


def _people(d, c, w):
    """人员：两个人形"""
    _circle(d, 38, 34, 13, c, w)
    d.arc([16, 50, 60, 96], start=180, end=360, fill=c, width=w)
    _circle(d, 70, 40, 10, c, w)
    d.arc([54, 56, 92, 94], start=185, end=355, fill=c, width=w)


def _cost(d, c, w):
    """成本：圆内货币符号"""
    _circle(d, 50, 50, 34, c, w)
    _line(d, [(38, 32), (50, 50), (62, 32)], c, w)
    _line(d, [(50, 50), (50, 70)], c, w)
    _line(d, [(36, 52), (64, 52)], c, w)
    _line(d, [(36, 62), (64, 62)], c, w)


def _delivery(d, c, w):
    """交付：包装箱"""
    _line(d, [(20, 34), (50, 20), (80, 34), (80, 70), (50, 84), (20, 70), (20, 34)],
          c, w)
    _line(d, [(20, 34), (50, 48), (80, 34)], c, w)
    _line(d, [(50, 48), (50, 84)], c, w)


def _risk(d, c, w):
    """风险：三角警示"""
    _line(d, [(50, 16), (88, 82), (12, 82), (50, 16)], c, w)
    _line(d, [(50, 40), (50, 62)], c, w + 1)
    _circle(d, 50, 72, 3, c, fill=c)


def _improve(d, c, w):
    """改善：上升折线 + 箭头"""
    _line(d, [(16, 74), (38, 52), (54, 64), (84, 28)], c, w + 1)
    _line(d, [(84, 28), (66, 28)], c, w)
    _line(d, [(84, 28), (84, 46)], c, w)


def _target(d, c, w):
    """目标：靶心"""
    _circle(d, 50, 50, 34, c, w)
    _circle(d, 50, 50, 21, c, w)
    _circle(d, 50, 50, 7, c, fill=c)


def _time(d, c, w):
    """周期：时钟"""
    _circle(d, 50, 50, 34, c, w)
    _line(d, [(50, 50), (50, 30)], c, w)
    _line(d, [(50, 50), (66, 58)], c, w)


def _check(d, c, w):
    """完成：圆内勾"""
    _circle(d, 50, 50, 34, c, w)
    _line(d, [(33, 51), (45, 64), (68, 37)], c, w + 1)


def _alert(d, c, w):
    """关注：圆内感叹号"""
    _circle(d, 50, 50, 34, c, w)
    _line(d, [(50, 30), (50, 56)], c, w + 1)
    _circle(d, 50, 68, 3.5, c, fill=c)


def _flask(d, c, w):
    """工艺/化学：锥形瓶"""
    _line(d, [(38, 18), (62, 18)], c, w)
    _line(d, [(42, 18), (42, 40), (20, 82), (80, 82), (58, 40), (58, 18)], c, w)
    _line(d, [(30, 64), (70, 64)], c, w)


def _doc(d, c, w):
    """文档/合规：文件"""
    _line(d, [(26, 14), (62, 14), (76, 30), (76, 88), (26, 88), (26, 14)], c, w)
    _line(d, [(62, 14), (62, 30), (76, 30)], c, w)
    for y in (48, 60, 72):
        _line(d, [(38, y), (64, y)], c, w - 1 if w > 3 else w)


def _search(d, c, w):
    """检查/审计：放大镜"""
    _circle(d, 44, 42, 24, c, w)
    _line(d, [(61, 60), (82, 82)], c, w + 2)


def _factory(d, c, w):
    """厂区/产线：厂房"""
    _line(d, [(14, 84), (14, 46), (38, 60), (38, 46), (62, 60), (62, 32), (86, 32),
              (86, 84)], c, w)
    _line(d, [(10, 84), (90, 84)], c, w)
    _line(d, [(72, 32), (72, 18)], c, w)


def _flow(d, c, w):
    """流程：三段带箭头"""
    for i in range(3):
        x = 8 + i * 32
        d.rounded_rectangle([x, 36, x + 24, 64], radius=5, outline=c, width=w)
        if i < 2:
            ax = x + 24
            _line(d, [(ax + 2, 50), (ax + 8, 50)], c, w)
            _line(d, [(ax + 8, 50), (ax + 3, 45)], c, w)
            _line(d, [(ax + 8, 50), (ax + 3, 55)], c, w)


ICONS = {
    "production": _production, "quality": _quality, "safety": _safety,
    "equipment": _equipment, "people": _people, "cost": _cost,
    "delivery": _delivery, "risk": _risk, "improve": _improve,
    "target": _target, "time": _time, "check": _check, "alert": _alert,
    "flask": _flask, "doc": _doc, "search": _search, "factory": _factory,
    "flow": _flow,
}

# 中文别名，方便按业务词直接取图标
ALIASES = {
    "产量": "production", "产能": "production", "质量": "quality",
    "安全": "safety", "设备": "equipment", "人员": "people", "人力": "people",
    "成本": "cost", "交付": "delivery", "风险": "risk", "改善": "improve",
    "目标": "target", "周期": "time", "完成": "check", "关注": "alert",
    "工艺": "flask", "合规": "doc", "文档": "doc", "审计": "search",
    "检查": "search", "厂区": "factory", "产线": "factory", "流程": "flow",
}


def draw_icon(name, path, color="F2F5F9", size=256, stroke=None, padding=0.10):
    """画一个图标存成透明 PNG。stroke 为设计坐标下的线宽（默认按尺寸自适应）。"""
    key = ALIASES.get(name, name)
    if key not in ICONS:
        raise ValueError("未知图标 %r。可用：%s" % (name, ", ".join(sorted(ICONS))))
    # 超采样后缩小，得到平滑边缘（PIL 不做抗锯齿）
    ss = 4
    side = int(BOX * ss)
    img = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    class Scaled:
        """把设计坐标放大到超采样画布。"""

        def __init__(self, dd):
            self.d = dd

        def line(self, pts, fill, width, joint="curve"):
            self.d.line([(p[0] * ss, p[1] * ss) for p in pts], fill=fill,
                        width=max(1, int(width * ss)), joint=joint)

        def ellipse(self, box, outline=None, fill=None, width=1):
            self.d.ellipse([v * ss for v in box], outline=outline, fill=fill,
                           width=max(1, int(width * ss)))

        def arc(self, box, start, end, fill, width):
            self.d.arc([v * ss for v in box], start=start, end=end, fill=fill,
                       width=max(1, int(width * ss)))

        def rounded_rectangle(self, box, radius, outline=None, fill=None, width=1):
            self.d.rounded_rectangle([v * ss for v in box], radius=radius * ss,
                                     outline=outline, fill=fill,
                                     width=max(1, int(width * ss)))

    w = stroke if stroke else 7
    ICONS[key](Scaled(d), _rgb(color), w)

    inner = int(size * (1 - 2 * padding))
    img = img.resize((inner, inner), Image.LANCZOS)
    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(img, ((size - inner) // 2, (size - inner) // 2), img)
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    out.save(path)
    return path


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        return 0
    cmd = sys.argv[1]
    if cmd == "list":
        print("可用图标（%d 个）：" % len(ICONS))
        for k in sorted(ICONS):
            zh = [z for z, e in ALIASES.items() if e == k]
            print("  %-12s %s" % (k, " ".join(zh)))
        return 0

    color = "F2F5F9"
    size = 256
    for i, a in enumerate(sys.argv):
        if a == "--color" and i + 1 < len(sys.argv):
            color = sys.argv[i + 1]
        if a == "--size" and i + 1 < len(sys.argv):
            size = int(sys.argv[i + 1])

    if cmd == "all":
        outdir = sys.argv[2] if len(sys.argv) > 2 else "icons"
        for k in sorted(ICONS):
            draw_icon(k, os.path.join(outdir, k + ".png"), color, size)
        print("已导出 %d 个图标到 %s/" % (len(ICONS), outdir))
        return 0

    if len(sys.argv) < 3:
        print("缺少输出路径")
        return 1
    print("已生成：" + draw_icon(cmd, sys.argv[2], color, size))
    return 0


if __name__ == "__main__":
    sys.exit(main())
