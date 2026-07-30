#!/usr/bin/env python3
"""生成内联 SVG 图表，供 HTML 汇报页嵌入。零依赖，纯字符串输出。

为什么不用 Chart.js / ECharts：汇报材料经常要在内网、离线、或者直接双击本地
html 文件的环境下打开，CDN 拉不到就整页空白；而且图表库的画布在打印时经常
输出空白。内联 SVG 没有这些问题，打印和转 PDF 都正常。

颜色跟随 CSS 变量（var(--primary) 等），所以换主题、切浅色模式时图表自动跟着变。

用法：
    from svg_chart import bar, hbar, line, stacked, donut
    svg = bar(["1月","2月"], [("2025",[148,132]), ("2026",[186,171])])
    html = "<div class='chart-wrap'>%s</div>" % svg

命令行自测：
    python3 svg_chart.py > demo.html
"""

import html as _html

# 默认取 CSS 变量，这样图表颜色跟随主题
PALETTE = ["var(--primary)", "var(--accent)", "var(--secondary)", "var(--good)",
           "#9b7bd4", "var(--bad)"]

SEM = {"good": "var(--good)", "warn": "var(--warn)", "bad": "var(--bad)",
       "neutral": "var(--neutral)", "primary": "var(--primary)",
       "secondary": "var(--secondary)", "accent": "var(--accent)",
       "hairline": "var(--hairline)"}


def _c(v, i=0):
    if v is None:
        return PALETTE[i % len(PALETTE)]
    return SEM.get(str(v), str(v) if str(v).startswith(("var(", "#"))
                   else "#" + str(v).lstrip("#"))


def _fmt(v, unit=""):
    if v == int(v):
        s = "{:,}".format(int(v))
    else:
        s = "{:,.1f}".format(v)
    return s + unit


def _nice_max(v, ticks=5):
    """把轴上限取整到好读的刻度（150 / 300 / 2,500 这种）。"""
    if v <= 0:
        return ticks
    import math
    raw = v / ticks
    mag = 10 ** math.floor(math.log10(raw))
    for m in (1, 2, 2.5, 5, 10):
        if raw <= mag * m:
            return mag * m * ticks
    return mag * 10 * ticks


def _open(w, h):
    return ('<svg class="chart" viewBox="0 0 %d %d" '
            'preserveAspectRatio="xMidYMid meet" role="img">' % (w, h))


def _esc(s):
    return _html.escape(str(s), quote=True)


def _text_w(s, fs=13):
    """估算文本像素宽度。中文按全角算，否则会低估近一半。"""
    return sum(1.0 if ord(c) > 0x2E80 else 0.55 for c in str(s)) * fs


def legend(series, colors=None):
    """图例。多序列时必须给，否则读者不知道哪个颜色是哪年。"""
    out = ['<div class="legend">']
    for i, (name, _) in enumerate(series):
        out.append('<span><i style="background:%s"></i>%s</span>'
                   % (_c(colors[i] if colors else None, i), _esc(name)))
    out.append("</div>")
    return "".join(out)


def bar(categories, series, colors=None, width=1160, height=340, unit="",
        max_scale=None, show_values=True, ticks=5):
    """纵向分组柱状图。series = [(名称, [数值...]), ...]"""
    L, R, T, B = 58, 12, 24, 34
    pw, ph = width - L - R, height - T - B
    vals = [v for _, s in series for v in s]
    vmax = max_scale or _nice_max(max(vals + [0]), ticks)
    n, m = len(categories), len(series)
    group = pw / n
    bw = min(56.0, (group * 0.66) / m)
    gap = 6 if m > 1 else 0
    total = bw * m + gap * (m - 1)

    o = [_open(width, height)]
    for i in range(ticks + 1):
        y = T + ph - ph * i / ticks
        o.append('<line class="grid" x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f"/>'
                 % (L, y, width - R, y))
        o.append('<text class="lbl" x="%.1f" y="%.1f" text-anchor="end">%s</text>'
                 % (L - 10, y + 4, _fmt(vmax * i / ticks)))
    o.append('<line class="axis" x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f"/>'
             % (L, T + ph, width - R, T + ph))

    for gi, cat in enumerate(categories):
        cx = L + group * (gi + 0.5)
        for si, (_, sv) in enumerate(series):
            v = sv[gi] if gi < len(sv) else 0
            bh = ph * max(0.0, v) / vmax
            x = cx - total / 2 + si * (bw + gap)
            o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" rx="2" '
                     'fill="%s"/>' % (x, T + ph - bh, bw, bh,
                                      _c(colors[si] if colors else None, si)))
            if show_values:
                o.append('<text class="val" x="%.1f" y="%.1f" '
                         'text-anchor="middle">%s</text>'
                         % (x + bw / 2, T + ph - bh - 7, _fmt(v, unit)))
        o.append('<text class="lbl" x="%.1f" y="%.1f" text-anchor="middle">%s</text>'
                 % (cx, height - 12, _esc(cat)))
    o.append("</svg>")
    return "".join(o)


def hbar(categories, values, colors=None, width=1160, height=340, unit="",
         max_scale=None, label_w=None):
    """横向条形图。类目名长、或要排名比较时用这个而不是纵向柱状。

    第一项显示在最上方（读者习惯最大的在顶部），传入顺序即显示顺序。
    """
    lw = label_w or min(260, max(90, max(len(str(c)) for c in categories) * 15 + 16))
    L, R, T, B = lw, 60, 8, 8
    pw, ph = width - L - R, height - T - B
    vmax = max_scale or _nice_max(max(list(values) + [0]), 4)
    n = len(categories)
    row = ph / n
    bh = min(46.0, row * 0.6)

    o = [_open(width, height)]
    for i, (cat, v) in enumerate(zip(categories, values)):
        cy = T + row * (i + 0.5)
        bw = pw * max(0.0, v) / vmax
        o.append('<text class="lbl" x="%.1f" y="%.1f" text-anchor="end">%s</text>'
                 % (L - 14, cy + 5, _esc(cat)))
        o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" rx="3" '
                 'fill="%s"/>' % (L, cy - bh / 2, max(bw, 1), bh,
                                  _c(colors[i] if colors else None, i)))
        o.append('<text class="val" x="%.1f" y="%.1f">%s</text>'
                 % (L + bw + 10, cy + 5, _fmt(v, unit)))
    o.append("</svg>")
    return "".join(o)


def line(categories, series, colors=None, width=1160, height=340, unit="",
         max_scale=None, min_scale=None, show_dots=True, ticks=5, area=False,
         target=None, target_label=None, show_values=False):
    """折线图。看趋势用这个，不要用饼图。

    unit 加在纵轴刻度上（折线默认不标数据点，标了容易糊成一片）。
    target 画一条横向参考线，用于"趋势 vs 目标"。
    show_values=True 时在每个数据点上方标数值。
    """
    # 右边距按最后一个类目标签的半宽预留：固定边距会把"2026Q2"这类长标签裁成
    # "2026Q"，而裁完看起来仍像个正常标签，不逐页看图根本发现不了。
    L = 58
    R = max(14, _text_w(categories[-1]) / 2 + 6) if categories else 14
    T, B = (26 if show_values else 24), 34
    pw, ph = width - L - R, height - T - B
    vals = [v for _, s in series for v in s] + ([target] if target is not None else [])
    vmax = max_scale or _nice_max(max(vals + [0]), ticks)
    vmin = 0 if min_scale is None else min_scale
    span = (vmax - vmin) or 1
    n = len(categories)
    step = pw / max(n - 1, 1)

    o = [_open(width, height)]
    for i in range(ticks + 1):
        y = T + ph - ph * i / ticks
        o.append('<line class="grid" x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f"/>'
                 % (L, y, width - R, y))
        o.append('<text class="lbl" x="%.1f" y="%.1f" text-anchor="end">%s</text>'
                 % (L - 10, y + 4, _fmt(vmin + span * i / ticks, unit)))
    o.append('<line class="axis" x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f"/>'
             % (L, T + ph, width - R, T + ph))

    if target is not None and vmin <= target <= vmax:
        ty = T + ph - ph * (target - vmin) / span
        o.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" '
                 'stroke="var(--ink-muted)" stroke-width="1.5" '
                 'stroke-dasharray="6 5"/>' % (L, ty, width - R, ty))
        o.append('<text class="lbl" x="%.1f" y="%.1f">%s</text>'
                 % (L + 8, ty - 7,
                    _esc(target_label or ("目标 " + _fmt(target, unit)))))

    for si, (_, sv) in enumerate(series):
        col = _c(colors[si] if colors else None, si)
        pts = [(L + step * i, T + ph - ph * (v - vmin) / span)
               for i, v in enumerate(sv)]
        d = " ".join("%s%.1f,%.1f" % ("M" if i == 0 else "L", x, y)
                     for i, (x, y) in enumerate(pts))
        if area:
            o.append('<path d="%s L%.1f,%.1f L%.1f,%.1f Z" fill="%s" '
                     'opacity="0.16"/>' % (d, pts[-1][0], T + ph, pts[0][0],
                                           T + ph, col))
        o.append('<path d="%s" fill="none" stroke="%s" stroke-width="3" '
                 'stroke-linejoin="round" stroke-linecap="round"/>' % (d, col))
        if show_dots:
            for x, y in pts:
                o.append('<circle cx="%.1f" cy="%.1f" r="4.5" fill="%s" '
                         'stroke="var(--bg)" stroke-width="2"/>' % (x, y, col))
        if show_values:
            for (x, y), v in zip(pts, sv):
                o.append('<text class="val" x="%.1f" y="%.1f" '
                         'text-anchor="middle">%s</text>'
                         % (x, y - 12, _fmt(v, unit)))
    for i, cat in enumerate(categories):
        o.append('<text class="lbl" x="%.1f" y="%.1f" text-anchor="middle">%s</text>'
                 % (L + step * i, height - 12, _esc(cat)))
    o.append("</svg>")
    return "".join(o)


def stacked(categories, series, colors=None, width=1160, height=340,
            horizontal=True, unit="%", show_values=True):
    """堆积条形图，用于构成占比。比多个饼图并排好读得多。

    标签压在色段上，所以段太窄时自动隐藏标签（挤在一起反而读不出来）。
    """
    if horizontal:
        lw = min(200, max(80, max(len(str(c)) for c in categories) * 15 + 16))
        L, R, T, B = lw, 20, 8, 26
        pw, ph = width - L - R, height - T - B
        n = len(categories)
        row = ph / n
        bh = min(46.0, row * 0.6)
        totals = [sum(s[i] for _, s in series if i < len(s)) for i in range(n)]
        vmax = max(totals + [1])
        o = [_open(width, height)]
        for i, cat in enumerate(categories):
            cy = T + row * (i + 0.5)
            o.append('<text class="lbl" x="%.1f" y="%.1f" text-anchor="end">%s</text>'
                     % (L - 14, cy + 5, _esc(cat)))
            x = L
            for si, (_, sv) in enumerate(series):
                v = sv[i] if i < len(sv) else 0
                w = pw * v / vmax
                col = _c(colors[si] if colors else None, si)
                o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" '
                         'fill="%s"/>' % (x, cy - bh / 2, w, bh, col))
                if show_values and w > 42:
                    o.append('<text class="val" x="%.1f" y="%.1f" '
                             'text-anchor="middle" fill="%s">%s</text>'
                             % (x + w / 2, cy + 5, _label_ink(col), _fmt(v, unit)))
                x += w
        o.append("</svg>")
        return "".join(o)

    L, R, T, B = 58, 12, 16, 34
    pw, ph = width - L - R, height - T - B
    n = len(categories)
    group = pw / n
    bw = min(76.0, group * 0.5)
    totals = [sum(s[i] for _, s in series if i < len(s)) for i in range(n)]
    vmax = max(totals + [1])
    o = [_open(width, height)]
    for i, cat in enumerate(categories):
        cx = L + group * (i + 0.5)
        y = T + ph
        for si, (_, sv) in enumerate(series):
            v = sv[i] if i < len(sv) else 0
            h = ph * v / vmax
            col = _c(colors[si] if colors else None, si)
            o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" '
                     'fill="%s"/>' % (cx - bw / 2, y - h, bw, h, col))
            if show_values and h > 22:
                o.append('<text class="val" x="%.1f" y="%.1f" '
                         'text-anchor="middle" fill="%s">%s</text>'
                         % (cx, y - h / 2 + 5, _label_ink(col), _fmt(v, unit)))
            y -= h
        o.append('<text class="lbl" x="%.1f" y="%.1f" text-anchor="middle">%s</text>'
                 % (cx, height - 12, _esc(cat)))
    o.append("</svg>")
    return "".join(o)


# 压在色段上的标签文字色。CSS 变量在 Python 里算不出亮度，所以按调色板位置
# 给出预判值：金色和浅蓝段配深字，其余配白字。
_DARK_LABEL_ON = {"var(--accent)", "var(--warn)", "var(--secondary)"}


def _label_ink(color):
    return "var(--ink-on-accent)" if color in _DARK_LABEL_ON else "#ffffff"


def donut(labels, values, colors=None, width=460, height=340, hole=0.58,
          unit="%", show_legend=True):
    """环形图。只用于"一个整体的构成"，且分段不超过 5 个 —— 更多就换 hbar。"""
    import math
    cx, cy = width / 2, height / 2 - (6 if show_legend else 0)
    r = min(width, height) * 0.38
    ri = r * hole
    total = sum(values) or 1
    o = [_open(width, height)]
    a0 = -math.pi / 2
    for i, v in enumerate(values):
        a1 = a0 + 2 * math.pi * v / total
        col = _c(colors[i] if colors else None, i)
        large = 1 if (a1 - a0) > math.pi else 0
        x0, y0 = cx + r * math.cos(a0), cy + r * math.sin(a0)
        x1, y1 = cx + r * math.cos(a1), cy + r * math.sin(a1)
        xi1, yi1 = cx + ri * math.cos(a1), cy + ri * math.sin(a1)
        xi0, yi0 = cx + ri * math.cos(a0), cy + ri * math.sin(a0)
        o.append('<path d="M%.1f,%.1f A%.1f,%.1f 0 %d 1 %.1f,%.1f '
                 'L%.1f,%.1f A%.1f,%.1f 0 %d 0 %.1f,%.1f Z" fill="%s" '
                 'stroke="var(--bg)" stroke-width="2"/>'
                 % (x0, y0, r, r, large, x1, y1, xi1, yi1, ri, ri, large,
                    xi0, yi0, col))
        am = (a0 + a1) / 2
        rl = (r + ri) / 2
        if v / total > 0.06:
            o.append('<text class="val" x="%.1f" y="%.1f" text-anchor="middle" '
                     'fill="%s">%s</text>'
                     % (cx + rl * math.cos(am), cy + rl * math.sin(am) + 5,
                        _label_ink(col), _fmt(v, unit)))
        a0 = a1
    o.append("</svg>")
    svg = "".join(o)
    if show_legend:
        items = "".join('<span><i style="background:%s"></i>%s</span>'
                        % (_c(colors[i] if colors else None, i), _esc(l))
                        for i, l in enumerate(labels))
        svg += '<div class="legend">%s</div>' % items
    return svg


def progress(items, target=100.0, show_target=True, min_scale=None,
             target_label=None):
    """达成率条组（HTML，不是 SVG）。items = [{label, value, text?, color?}]

    在 target 处画参考线，否则读者无法看出谁达标了。

    **值域集中时一定要传 min_scale。** 例如四个指标在 94-98 之间、目标 97，默认从 0
    起算的话四根条长度差不到 5%，肉眼分不出来，"谁没达标"只能靠颜色和数字撑着。
    传 min_scale=90 把起点抬高，长度差才读得出来。
    """
    vals = [float(i.get("value", 0)) for i in items]
    lo = 0.0 if min_scale is None else float(min_scale)
    hi = max([target] + vals) * (1.04 if min_scale is None else 1.01)
    span = (hi - lo) or 1.0

    def pct(v):
        return max(0.0, min(100.0, (float(v) - lo) / span * 100))

    out = ['<div class="progress-list">']
    if show_target:
        out.append('<div class="pr pr-axis"><div class="l"></div>'
                   '<div class="track-note"><span style="left:%.1f%%">%s</span>'
                   '</div><div class="v"></div></div>'
                   % (pct(target),
                      _esc(target_label or ("目标 %g%%" % target))))
    for it, v in zip(items, vals):
        # 阈值与 PPT 版的 progress_bars 保持一致：差 5% 以内算"接近"
        c = it.get("color") or ("good" if v >= target
                                else "warn" if v >= target * 0.95 else "bad")
        mark = ('<i class="target" style="left:%.1f%%"></i>' % pct(target)
                if show_target else "")
        out.append(
            '<div class="pr"><div class="l">%s</div>'
            '<div class="track"><div class="fill" style="width:%.1f%%;'
            'background:%s"></div>%s</div>'
            '<div class="v t-%s">%s</div></div>'
            % (_esc(it.get("label", "")), pct(v), _c(c), mark,
               c if c in SEM else "accent",
               _esc(it.get("text") or ("%g%%" % v))))
    if min_scale is not None:
        out.append('<div class="pr-scale">横轴起点 %g，用于放大差异</div>' % lo)
    out.append("</div>")
    return "".join(out)


if __name__ == "__main__":
    cats = ["1月", "2月", "3月", "4月", "5月", "6月"]
    s2 = [("2025 年", [148, 132, 165, 171, 178, 182]),
          ("2026 年", [186, 171, 208, 232, 241, 246])]
    print("<!doctype html><meta charset='utf-8'>"
          "<link rel='stylesheet' href='../assets/deck.css'>"
          "<body style='background:#0a1f3c;padding:24px'>")
    for title, svg in (("bar", legend(s2) + bar(cats, s2, unit=" 吨")),
                       ("line", line(cats, s2)),
                       ("hbar", hbar(["结晶周期偏长", "设备非计划停机", "人员不足"],
                                     [52, 26, 15], colors=["bad", "warn", "neutral"],
                                     unit="%")),
                       ("stacked", stacked(["A 产品", "B 产品"],
                                           [("前处理", [12, 14]), ("合成", [26, 24]),
                                            ("结晶", [38, 41]), ("包装", [24, 21])])),
                       ("donut", donut(["达标", "接近", "未达标"], [62, 26, 12]))):
        print("<h3 style='color:#fff;font:700 18px sans-serif'>%s</h3>"
              "<div style='height:340px;margin-bottom:28px'>%s</div>" % (title, svg))
    print("</body>")
