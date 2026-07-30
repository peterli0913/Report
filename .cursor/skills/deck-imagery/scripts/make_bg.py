#!/usr/bin/env python3
"""生成汇报用的背景图与装饰底纹，只依赖 Pillow。

背景图的唯一职责是给页面质感，不能抢文字的注意力。所以这里生成的都是低对比、
无焦点、不含具体形象的图案 —— 照片型背景很容易让压在上面的标题读不清。

用法：
    python3 make_bg.py gradient bg.png --from 0A1F3C --to 1E6FBF --angle 135
    python3 make_bg.py glow bg.png --base 0A1F3C --accent 1E6FBF
    python3 make_bg.py mesh bg.png --base 0A1F3C --line 2C4E7A
    python3 make_bg.py molecule bg.png --base 0A1F3C --accent 2AA9DB   # 分子网络
    python3 make_bg.py dots bg.png --base 0A1F3C --accent 2C4E7A

通用参数：--size 1920x1080  --opacity 0-100（图案强度，默认按类型给合适值）

在代码里用：
    from make_bg import gradient, glow, mesh, molecule, dots
    gradient("bg.png", "0A1F3C", "1E6FBF", angle=135, size=(1920, 1080))
"""

import math
import os
import random
import sys

from PIL import Image, ImageDraw, ImageFilter

DEFAULT_SIZE = (1920, 1080)


def _rgb(h):
    h = str(h).lstrip("#")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _mix(c1, c2, t):
    return tuple(int(round(a + (b - a) * t)) for a, b in zip(c1, c2))


def _save(img, path):
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    img.convert("RGB").save(path, quality=92)
    return path


def gradient(path, c_from, c_to, angle=135, size=DEFAULT_SIZE):
    """线性渐变。angle 为度数，0=从左到右，90=从上到下，135=左上到右下。"""
    w, h = size
    a, b = _rgb(c_from), _rgb(c_to)
    rad = math.radians(angle)
    dx, dy = math.cos(rad), math.sin(rad)
    # 先在小图上算渐变再放大，比逐像素快几个数量级且视觉无差别
    sw, sh = 160, 90
    small = Image.new("RGB", (sw, sh))
    px = small.load()
    span = abs(dx) * sw + abs(dy) * sh
    off = (sw if dx < 0 else 0) * abs(dx) + (sh if dy < 0 else 0) * abs(dy)
    for y in range(sh):
        for x in range(sw):
            t = (x * dx + y * dy + off) / (span or 1)
            px[x, y] = _mix(a, b, max(0.0, min(1.0, t)))
    return _save(small.resize((w, h), Image.BICUBIC), path)


def glow(path, base, accent, size=DEFAULT_SIZE, cx=0.72, cy=0.28, radius=0.85,
         opacity=55):
    """径向光晕：底色 + 一处柔和亮区。文字放在光晕对侧（默认左下）。"""
    w, h = size
    bg = Image.new("RGB", (w, h), _rgb(base))
    sw, sh = 240, 135
    layer = Image.new("L", (sw, sh), 0)
    px = layer.load()
    ccx, ccy = cx * sw, cy * sh
    rr = radius * max(sw, sh)
    for y in range(sh):
        for x in range(sw):
            d = math.hypot(x - ccx, y - ccy) / rr
            px[x, y] = int(max(0.0, 1.0 - d) ** 2 * 255)
    layer = layer.resize((w, h), Image.BICUBIC)
    tint = Image.new("RGB", (w, h), _rgb(accent))
    layer = layer.point(lambda v: int(v * opacity / 100))
    return _save(Image.composite(tint, bg, layer), path)


def mesh(path, base, line, size=DEFAULT_SIZE, spacing=90, opacity=40, width=1):
    """细网格线：给页面一点"工程图"质感，适合技术类汇报。"""
    w, h = size
    bg = Image.new("RGB", (w, h), _rgb(base))
    layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    c = _rgb(line) + (int(255 * opacity / 100),)
    for x in range(0, w + 1, spacing):
        d.line([(x, 0), (x, h)], fill=c, width=width)
    for y in range(0, h + 1, spacing):
        d.line([(0, y), (w, y)], fill=c, width=width)
    bg = bg.convert("RGBA")
    bg.alpha_composite(layer)
    return _save(bg, path)


def dots(path, base, accent, size=DEFAULT_SIZE, spacing=46, radius=3, opacity=62):
    """点阵底纹：比网格更轻，几乎不干扰文字。"""
    w, h = size
    bg = Image.new("RGBA", (w, h), _rgb(base) + (255,))
    layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    c = _rgb(accent) + (int(255 * opacity / 100),)
    for y in range(spacing // 2, h, spacing):
        for x in range(spacing // 2, w, spacing):
            d.ellipse([x - radius, y - radius, x + radius, y + radius], fill=c)
    bg.alpha_composite(layer)
    return _save(bg, path)


def molecule(path, base, accent, size=DEFAULT_SIZE, nodes=26, opacity=42,
             seed=7, link_dist=0.30):
    """分子/网络结构：节点 + 连线，医药化工类汇报的封面很贴合。

    节点分布在四周、中部留空，这样中间放标题不会被线条穿过。
    """
    w, h = size
    rnd = random.Random(seed)
    bg = Image.new("RGBA", (w, h), _rgb(base) + (255,))
    layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    a = int(255 * opacity / 100)
    c_line = _rgb(accent) + (int(a * 0.55),)
    c_node = _rgb(accent) + (a,)

    # 分层采样：每个网格格子里放一个带抖动的节点，保证四周分布均匀。
    # 纯随机撒点在节点数不多时经常聚成一团、另一侧空白。
    pts = []
    cols, rows = 7, 5
    for r in range(rows):
        for c in range(cols):
            x = (c + 0.15 + 0.70 * rnd.random()) / cols
            y = (r + 0.15 + 0.70 * rnd.random()) / rows
            if 0.24 < x < 0.76 and 0.26 < y < 0.72:
                continue        # 中间区域留给标题，不放节点
            pts.append((x * w, y * h))
    rnd.shuffle(pts)
    pts = pts[:nodes]

    lim = link_dist * math.hypot(w, h)
    for i, p in enumerate(pts):
        for q in pts[i + 1:]:
            if math.dist(p, q) < lim:
                d.line([p, q], fill=c_line, width=2)
    for x, y in pts:
        r = rnd.choice((4, 5, 7))
        d.ellipse([x - r, y - r, x + r, y + r], fill=c_node)
    layer = layer.filter(ImageFilter.GaussianBlur(0.6))
    bg.alpha_composite(layer)
    return _save(bg, path)


def scrim(path, image, base="000000", opacity=55, size=None):
    """给已有图片压一层半透明遮罩，让压在上面的文字可读。

    深色底上放照片时这一步不能省 —— 照片的高光区域会把白色标题吃掉。
    """
    img = Image.open(image).convert("RGB")
    if size:
        img = img.resize(size, Image.LANCZOS)
    overlay = Image.new("RGB", img.size, _rgb(base))
    return _save(Image.blend(img, overlay, opacity / 100.0), path)


MODES = {"gradient": gradient, "glow": glow, "mesh": mesh, "dots": dots,
         "molecule": molecule, "scrim": scrim}


def main():
    if len(sys.argv) < 3 or sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        return 0
    mode, out = sys.argv[1], sys.argv[2]
    if mode not in MODES:
        print("未知模式 %r，可用：%s" % (mode, ", ".join(MODES)))
        return 1

    def opt(name, default=None, cast=str):
        for i, a in enumerate(sys.argv):
            if a == "--" + name and i + 1 < len(sys.argv):
                return cast(sys.argv[i + 1])
        return default

    size = opt("size", "1920x1080")
    size = tuple(int(v) for v in str(size).lower().split("x"))
    base = opt("base", "0A1F3C")
    accent = opt("accent", "1E6FBF")
    op = opt("opacity", None, int)

    if mode == "gradient":
        p = gradient(out, opt("from", base), opt("to", accent),
                     opt("angle", 135, float), size)
    elif mode == "scrim":
        p = scrim(out, opt("image"), base, op or 55, size)
    else:
        kw = {"size": size}
        if op is not None:
            kw["opacity"] = op
        if mode == "mesh":
            p = mesh(out, base, opt("line", "2C4E7A"), **kw)
        else:
            p = MODES[mode](out, base, accent, **kw)
    print("已生成：" + p)
    return 0


if __name__ == "__main__":
    sys.exit(main())
