"""deckkit —— 面向中文高层汇报的 PPT 生成组件库（基于 python-pptx）。

设计目标：让调用方只描述"内容和结论"，不必计算坐标、不必和 OOXML 细节搏斗。

三个必须在库层面解决的问题（否则每次生成都会踩）：
  1. 东亚字体：python-pptx 的 font.name 只写 <a:latin>，中文会回落到主题字体。
     本库对每个 run 同时写 latin / ea / cs，中文才真正生效。
  2. 文本溢出：PPT 不会自动缩小文字，超出就被裁掉。本库按全角/半角估算宽度，
     预测行数并自动降档字号（fit_size），并提供 check_overflow 事后核查。
  3. 视觉一致性：颜色、字号、间距、阴影全部走 Theme，禁止散落魔法数字。

坐标系统：所有对外的坐标和尺寸都以"设计英寸"表示，基准画布 13.333 x 7.5 英寸。
换用 26.67 x 15 的大画布时只需改 canvas 参数，组件代码完全不用动 —— Deck 内部
用 self.scale 统一换算。
"""

from __future__ import annotations

import copy
import math
import re
from dataclasses import dataclass, field, replace

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# 基准画布：所有设计单位以此为准
BASE_W = 13.3333
BASE_H = 7.5

CANVAS = {
    "wide": (13.3333, 7.5),      # 16:9 标准，PowerPoint 默认
    "large": (26.6667, 15.0),    # 16:9 双倍画布，兼容既有企业模板
    "a4": (11.6929, 8.2677),     # 4:3 偏纸质，打印汇报用
}


# --------------------------------------------------------------------------
# 主题
# --------------------------------------------------------------------------

@dataclass
class Theme:
    """一套完整的视觉规范。字号是基准画布下的 pt 值，Deck 会按 scale 换算。"""

    name: str = "asymchem-dark"

    # 底色与层次
    bg: str = "0A1F3C"           # 页面底色
    bg_alt: str = "07172C"       # 次级底色（章节页/封面）
    surface: str = "122E52"      # 卡片底色
    surface_alt: str = "1B3E68"  # 卡片次级底色 / 表头
    hairline: str = "2C4E7A"     # 分隔线

    # 文字
    ink: str = "F2F5F9"          # 主文字
    ink_muted: str = "9FB3CC"    # 次要文字、坐标轴、脚注
    ink_on_accent: str = "0A1F3C"  # 压在亮色块上的文字

    # 品牌色
    primary: str = "1E6FBF"
    secondary: str = "2AA9DB"
    accent: str = "E5B620"       # 强调色，用于关键数字和唯一重点

    # 语义色（达成/关注/风险）—— 这一组是"填充色"，用于色块、图表、进度条。
    # 亮度要么足够暗（配白字）要么足够亮（配深字），落在中间地带的颜色两种字色都压不住。
    good: str = "27A567"
    warn: str = "E5A32B"
    bad: str = "C93B33"
    neutral: str = "7E93AD"

    # 语义色的"文字版"。填充够醒目和文字够可读是两个不同要求：饱和的绿做色块很好，
    # 但当成 12pt 小字压在深底上就达不到 WCAG 4.5:1。所以两组分开定义。
    good_text: str = "2DBF77"
    warn_text: str = "E5A32B"
    bad_text: str = "FF7A6B"
    neutral_text: str = "A0B0C4"
    accent_text: str = "E5B620"
    primary_text: str = "6FB0EA"
    secondary_text: str = "44B4DF"

    # 图表序列色，按顺序取用
    series: tuple = ("1E6FBF", "E5B620", "2AA9DB", "27A567", "9B7BD4", "C93B33")

    # 字体：中文与西文分开，库会同时写 latin/ea/cs
    font_cn: str = "微软雅黑"
    font_en: str = "Arial"
    font_num: str = "Arial"      # 数字用等宽感更强的西文字体

    # 字号阶梯（基准画布 pt）
    # 投屏汇报的正文下限是 14pt：再小后排就看不清，宁可删内容也别缩字。
    size_cover_title: float = 40
    size_cover_sub: float = 18
    size_section_title: float = 34
    size_title: float = 23       # 内容页行动标题
    size_eyebrow: float = 11.5   # 标题上方的栏目标签
    size_kpi: float = 40         # KPI 大数字
    size_kpi_label: float = 11.5
    size_h: float = 16           # 小节标题
    size_body: float = 14        # 正文
    size_small: float = 12
    size_note: float = 9.5       # 脚注/数据来源

    # 版面
    margin: float = 0.62         # 页边距
    gap: float = 0.24            # 元素间距
    radius: float = 0.045        # 圆角占短边比例
    dark: bool = True            # 深色主题标记，影响图表网格线等细节

    def color(self, key: str) -> str:
        """按语义名或直接 hex 取色，方便调用方写 'good' / 'accent' / 'FF0000'。"""
        if not key:
            return self.ink
        k = str(key).strip().lstrip("#")
        if re.fullmatch(r"[0-9A-Fa-f]{6}", k):
            return k.upper()
        return getattr(self, str(key), self.ink)

    def text_color(self, key: str) -> str:
        """语义色当文字用时走这里，自动换成对比度达标的文字版。

        调用方只需记住：fill= 用填充色名，color= 用同一个名字，库负责换算。
        """
        m = {"good": self.good_text, "warn": self.warn_text,
             "bad": self.bad_text, "neutral": self.neutral_text,
             "accent": self.accent_text, "primary": self.primary_text,
             "secondary": self.secondary_text}
        v = m.get(str(key))
        return self.color(v) if v else self.color(key)

    def luminance(self, key: str) -> float:
        """WCAG 相对亮度，0（黑）到 1（白）。"""
        h = self.color(key)
        vals = []
        for i in (0, 2, 4):
            c = int(h[i:i + 2], 16) / 255.0
            vals.append(c / 12.92 if c <= 0.03928
                        else ((c + 0.055) / 1.055) ** 2.4)
        return 0.2126 * vals[0] + 0.7152 * vals[1] + 0.0722 * vals[2]

    def ink_on(self, bg: str) -> str:
        """返回压在 bg 上对比度更高的文字色（hex）。

        色块上的文字色必须这样算出来，不能按"哪些颜色算亮色"写死 —— 换主题或
        换品牌色后，写死的判断会产出绿底深字这种读不清的组合。
        """
        lb = self.luminance(bg)
        ld = self.luminance(self.ink_on_accent)
        ratio_light = 1.05 / (lb + 0.05)
        ratio_dark = (lb + 0.05) / (ld + 0.05)
        return "FFFFFF" if ratio_light > ratio_dark else self.color(self.ink_on_accent)

    def block_ink(self, fill: str) -> tuple:
        """返回 (主文字色, 次文字色)。低饱和底用常规墨色，饱和色块用反差色。"""
        if str(fill) in ("surface", "surface_alt", "bg", "bg_alt", None):
            return self.ink, self.ink_muted
        ink = self.ink_on(fill)
        return ink, ink

    def variant(self, **kw) -> "Theme":
        return replace(self, **kw)

    def with_readable_text(self, min_ratio: float = 4.6) -> "Theme":
        """按当前底色重算全部语义文字色，保证在任何一档底色上都达到 min_ratio。

        改过 bg / surface（例如换成企业 VI 色）之后调一次 —— 底色一变，原来达标的
        文字色就可能不达标了，而这类问题在屏幕上不明显、投屏时却直接消失。

            T = DARK.variant(bg="003669", surface="0B4880").with_readable_text()
        """
        import colorsys

        bgs = [self.bg, self.bg_alt, self.surface, self.surface_alt]
        lums = [self.luminance(b) for b in bgs]
        # 深底以最亮的一档为约束（文字要够亮），浅底以最暗的一档为约束
        if sum(lums) / len(lums) < 0.4:
            target = min_ratio * (max(lums) + 0.05) - 0.05
            want_brighter = True
        else:
            target = (min(lums) + 0.05) / min_ratio - 0.05
            want_brighter = False

        def tune(hexc):
            h = self.color(hexc)
            r, g, b = [int(h[i:i + 2], 16) / 255.0 for i in (0, 2, 4)]
            hh, _, ss = colorsys.rgb_to_hls(r, g, b)
            best = None
            for i in range(2, 99):
                rr, gg, bb = colorsys.hls_to_rgb(hh, i / 100.0, ss)
                cand = "%02X%02X%02X" % (round(rr * 255), round(gg * 255),
                                         round(bb * 255))
                lc = self.luminance(cand)
                ok = lc >= target if want_brighter else lc <= target
                if ok:
                    # 取刚好达标的那一档，避免颜色被推到发白或发黑
                    return cand
                best = cand
            return best or h

        fixed = {}
        for f in ("good_text", "warn_text", "bad_text", "neutral_text",
                  "accent_text", "primary_text", "secondary_text", "ink_muted"):
            cur = self.color(getattr(self, f))
            lc = self.luminance(cur)
            if (lc >= target) if want_brighter else (lc <= target):
                continue                      # 已经达标就别动，保留原有色感
            fixed[f] = tune(cur)
        return replace(self, **fixed) if fixed else self


# 浅色主题：打印、投屏亮环境、以及需要贴合既有白底模板时使用
LIGHT = Theme(
    name="asymchem-light",
    bg="FFFFFF", bg_alt="F4F7FB", surface="F4F7FB", surface_alt="E6EDF6",
    hairline="D2DDEA",
    ink="14243C", ink_muted="576C89", ink_on_accent="14243C",
    primary="1E6FBF", secondary="2AA9DB", accent="C8880C",
    good="1A7F50", warn="C2851A", bad="C2372F", neutral="8496AC",
    # 浅底上文字要往暗走（深底是往亮走），所以整组与 DARK 相反
    good_text="1A7A4D", warn_text="8F6213", bad_text="B8332B",
    neutral_text="596C84", accent_text="906209", primary_text="1D6CBA",
    secondary_text="197396",
    series=("1E6FBF", "C8880C", "2AA9DB", "1A7F50", "7B5CB8", "C2372F"),
    dark=False,
)

DARK = Theme()

# 深灰中性主题：不想用蓝的场合
SLATE = Theme(
    name="slate",
    bg="1C2229", bg_alt="14181D", surface="262E37", surface_alt="333D48",
    hairline="3D4854",
    ink="F0F3F6", ink_muted="A3AEBB",
    primary="4C8DBF", secondary="6BB8C9", accent="E2A93B",
    good_text="2DBD76", bad_text="FF7A6B", neutral_text="97A8BD",
    accent_text="E2A93B", primary_text="7CACD0", secondary_text="6BB8C9",
    series=("4C8DBF", "E2A93B", "6BB8C9", "5FA97F", "A78BC4", "CC6155"),
)

THEMES = {"dark": DARK, "light": LIGHT, "slate": SLATE,
          "asymchem-dark": DARK, "asymchem-light": LIGHT}


# --------------------------------------------------------------------------
# 矩形区域：组件的定位语言
# --------------------------------------------------------------------------

@dataclass
class Rect:
    """设计英寸下的矩形。所有 split/grid 都返回新的 Rect，不改自身。"""

    x: float
    y: float
    w: float
    h: float

    @property
    def right(self) -> float:
        return self.x + self.w

    @property
    def bottom(self) -> float:
        return self.y + self.h

    @property
    def cx(self) -> float:
        return self.x + self.w / 2

    @property
    def cy(self) -> float:
        return self.y + self.h / 2

    def inset(self, d: float = 0.0, dy: float | None = None) -> "Rect":
        dy = d if dy is None else dy
        return Rect(self.x + d, self.y + dy, self.w - 2 * d, self.h - 2 * dy)

    def pad(self, top=0.0, right=0.0, bottom=0.0, left=0.0) -> "Rect":
        return Rect(self.x + left, self.y + top,
                    self.w - left - right, self.h - top - bottom)

    def split_h(self, *ratios: float, gap: float = 0.0) -> list["Rect"]:
        """横向切分。ratios 为相对比例，自动归一化。"""
        total = sum(ratios)
        avail = self.w - gap * (len(ratios) - 1)
        out, cur = [], self.x
        for r in ratios:
            w = avail * r / total
            out.append(Rect(cur, self.y, w, self.h))
            cur += w + gap
        return out

    def split_v(self, *ratios: float, gap: float = 0.0) -> list["Rect"]:
        total = sum(ratios)
        avail = self.h - gap * (len(ratios) - 1)
        out, cur = [], self.y
        for r in ratios:
            h = avail * r / total
            out.append(Rect(self.x, cur, self.w, h))
            cur += h + gap
        return out

    def grid(self, cols: int = 1, rows: int = 1, gap: float = 0.24,
             gap_y: float | None = None) -> list["Rect"]:
        """行优先返回 cols*rows 个等分格。"""
        gy = gap if gap_y is None else gap_y
        cw = (self.w - gap * (cols - 1)) / cols
        ch = (self.h - gy * (rows - 1)) / rows
        return [Rect(self.x + c * (cw + gap), self.y + r * (ch + gy), cw, ch)
                for r in range(rows) for c in range(cols)]


# --------------------------------------------------------------------------
# 文本度量：中文排版的核心，用于预测溢出
# --------------------------------------------------------------------------

# 单字符宽度相对字号的倍数（微软雅黑/Arial 实测近似）
_W_FULL = 1.0      # CJK 汉字、全角标点
_W_UPPER = 0.62
_W_LOWER = 0.50
_W_DIGIT = 0.55
_W_SPACE = 0.28
_W_NARROW = 0.30   # .,:;'|! 等窄标点

_CJK = (
    (0x4E00, 0x9FFF), (0x3400, 0x4DBF), (0xF900, 0xFAFF),
    (0x3000, 0x303F), (0xFF00, 0xFF60), (0xFFE0, 0xFFE6),
    (0x3040, 0x30FF), (0xAC00, 0xD7AF),
)


def char_width(ch: str) -> float:
    """返回单字符宽度相对字号的倍数。"""
    o = ord(ch)
    for lo, hi in _CJK:
        if lo <= o <= hi:
            return _W_FULL
    if ch == " ":
        return _W_SPACE
    if ch in ".,:;'`|!i lIjt[]()":
        return _W_NARROW
    if ch.isdigit():
        return _W_DIGIT
    if ch.isupper():
        return _W_UPPER
    return _W_LOWER


def text_width(text: str, size_pt: float) -> float:
    """估算单行文本宽度（英寸）。size_pt 为设计单位字号。"""
    return sum(char_width(c) for c in text) * size_pt / 72.0


def wrapped_lines(text: str, box_w: float, size_pt: float) -> int:
    """估算在给定宽度下的折行数。中文可任意断行，西文按空格断词。"""
    if not text:
        return 1
    usable = max(box_w, 0.05)
    total = 0
    for para in str(text).split("\n"):
        if not para:
            total += 1
            continue
        # 按"CJK 单字 + 西文单词"切成不可分割的最小单元
        tokens = re.findall(r"[A-Za-z0-9][A-Za-z0-9.\-%/,]*|\s|.", para)
        line = 0.0
        n = 1
        for tk in tokens:
            w = text_width(tk, size_pt)
            if line + w > usable and line > 0:
                n += 1
                line = 0.0 if tk.isspace() else w
            else:
                line += w
        total += n
    return total


def fit_size(text: str, box_w: float, box_h: float, size_pt: float,
             min_pt: float = 8.0, line_spacing: float = 1.28,
             step: float = 0.5) -> float:
    """从 size_pt 起逐档下调，直到预测行数能装进 box_h。返回可用字号。

    这是"文字被裁掉"的主要防线：PPT 不会替你缩字。
    """
    s = size_pt
    while s > min_pt:
        lines = wrapped_lines(text, box_w, s)
        if lines * s * line_spacing / 72.0 <= box_h:
            return s
        s -= step
    return min_pt


def text_height(text: str, box_w: float, size_pt: float,
                line_spacing: float = 1.28) -> float:
    """估算文本块高度（英寸），用于自适应排版。"""
    return wrapped_lines(text, box_w, size_pt) * size_pt * line_spacing / 72.0


# --------------------------------------------------------------------------
# OOXML 低层工具
# --------------------------------------------------------------------------

def _el(tag: str):
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls
    return parse_xml("<%s %s/>" % (tag, nsdecls("a")))


def apply_font(font, theme: Theme, size=None, bold=None, italic=None,
               color=None, font_en=None, font_cn=None):
    """给 python-pptx 的 Font 对象同时写 latin / ea / cs。

    只设 font.name 的话，中文会走主题的东亚字体（通常是宋体或 Calibri 的回落），
    这是中文 PPT 最常见的"字体没生效"根因。
    """
    latin = font_en or theme.font_en
    ea = font_cn or theme.font_cn
    font.name = latin                      # 写入 <a:latin>
    if size is not None:
        font.size = Pt(size)
    if bold is not None:
        font.bold = bold
    if italic is not None:
        font.italic = italic
    if color is not None:
        # 走 text_color：语义色自动换成文字版，hex 与 ink/ink_muted 原样通过
        font.color.rgb = RGBColor.from_string(theme.text_color(color))

    rPr = font._element
    for tag, face in (("a:ea", ea), ("a:cs", latin)):
        node = rPr.find(qn(tag))
        if node is None:
            node = _el(tag)
            # rPr 子元素顺序固定：latin 之后紧跟 ea，再 cs
            anchor = rPr.find(qn("a:latin"))
            if anchor is not None:
                anchor.addnext(node) if tag == "a:ea" else anchor.addnext(node)
            else:
                rPr.append(node)
        node.set("typeface", face)
    # 确保 cs 在 ea 之后
    ea_node, cs_node = rPr.find(qn("a:ea")), rPr.find(qn("a:cs"))
    if ea_node is not None and cs_node is not None:
        if list(rPr).index(cs_node) < list(rPr).index(ea_node):
            ea_node.addnext(cs_node)


def add_shadow(shape, blur=10.0, dist=4.0, direction=5400000,
               color="000000", alpha=22):
    """给形状加外阴影。blur/dist 单位为 pt。alpha 为不透明度百分比。"""
    spPr = shape._element.spPr
    for old in spPr.findall(qn("a:effectLst")):
        spPr.remove(old)
    eff = _el("a:effectLst")
    shdw = _el("a:outerShdw")
    shdw.set("blurRad", str(int(blur * 12700)))
    shdw.set("dist", str(int(dist * 12700)))
    shdw.set("dir", str(int(direction)))
    shdw.set("rotWithShape", "0")
    clr = _el("a:srgbClr")
    clr.set("val", color.lstrip("#").upper())
    a = _el("a:alpha")
    a.set("val", str(int(alpha * 1000)))
    clr.append(a)
    shdw.append(clr)
    eff.append(shdw)
    # spPr 子元素顺序：... ln, effectLst, effectDag, scene3d, sp3d, extLst
    spPr.insert_element_before(eff, "a:effectDag", "a:scene3d", "a:sp3d", "a:extLst")
    return shape


def set_fill_alpha(shape, transparency: float):
    """给纯色填充加透明度（0=不透明，100=全透明）。

    python-pptx 没有暴露填充透明度，但压在背景图上的遮罩必须半透明 ——
    不透明的遮罩会把图整张盖住，等于没放图。
    """
    if not transparency:
        return shape
    spPr = shape._element.spPr
    solid = spPr.find(qn("a:solidFill"))
    if solid is None:
        return shape
    clr = solid.find(qn("a:srgbClr"))
    if clr is None:
        return shape
    for old in clr.findall(qn("a:alpha")):
        clr.remove(old)
    a = _el("a:alpha")
    a.set("val", str(int(max(0.0, min(100.0, 100 - transparency)) * 1000)))
    clr.append(a)
    return shape


def clear_shadow(shape):
    """去掉形状阴影（python-pptx 默认从主题继承，卡片叠卡片时要关掉）。"""
    spPr = shape._element.spPr
    for old in spPr.findall(qn("a:effectLst")):
        spPr.remove(old)
    spPr.insert_element_before(_el("a:effectLst"), "a:effectDag", "a:scene3d",
                               "a:sp3d", "a:extLst")
    return shape


# --------------------------------------------------------------------------
# Slide 包装：组件都挂在这里
# --------------------------------------------------------------------------

class Slide:
    """一页幻灯片。坐标一律用设计英寸，内部乘 deck.scale 后写入。"""

    def __init__(self, deck: "Deck", slide, bg: str | None = None):
        self.deck = deck
        self.raw = slide
        self.theme = deck.theme
        self._title_bottom = deck.theme.margin
        if bg is not None:
            self.set_bg(bg)

    # ---- 单位换算 ----
    def _i(self, v: float) -> Emu:
        return Inches(v * self.deck.scale)

    def _p(self, v: float) -> Pt:
        return Pt(v * self.deck.scale)

    # ---- 版面区域 ----
    @property
    def page(self) -> Rect:
        return Rect(0, 0, BASE_W, BASE_H)

    @property
    def safe(self) -> Rect:
        m = self.theme.margin
        return Rect(m, m, BASE_W - 2 * m, BASE_H - 2 * m)

    @property
    def body(self) -> Rect:
        """标题之下、脚注之上的内容区。加过标题后自动下移。"""
        m = self.theme.margin
        top = self._title_bottom
        return Rect(m, top, BASE_W - 2 * m, BASE_H - top - m - 0.12)

    # ---- 背景 ----
    def set_bg(self, color: str):
        fill = self.raw.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(self.theme.color(color))
        return self

    def bg_image(self, path: str, transparency: int | None = None):
        """整页铺底图。会插到最底层。"""
        pic = self.raw.shapes.add_picture(path, 0, 0,
                                         self._i(BASE_W), self._i(BASE_H))
        self.raw.shapes._spTree.remove(pic._element)
        self.raw.shapes._spTree.insert(2, pic._element)
        return pic

    # ---- 基础图形 ----
    def rect(self, at: Rect, fill=None, line=None, line_w=1.0,
             radius: float | None = None, shadow=False, shape=None,
             transparency: float = 0):
        """画一个矩形/圆角矩形。radius 为圆角占短边比例，None 用主题值。

        transparency 0-100，用于压在背景图上的半透明遮罩。
        """
        if shape is None:
            shape = MSO_SHAPE.ROUNDED_RECTANGLE if (radius is None or radius > 0) \
                else MSO_SHAPE.RECTANGLE
        sh = self.raw.shapes.add_shape(shape, self._i(at.x), self._i(at.y),
                                       self._i(at.w), self._i(at.h))
        if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            r = self.theme.radius if radius is None else radius
            try:
                sh.adjustments[0] = max(0.0, min(0.5, r))
            except (IndexError, ValueError):
                pass
        if fill is None:
            sh.fill.background()
        else:
            sh.fill.solid()
            sh.fill.fore_color.rgb = RGBColor.from_string(self.theme.color(fill))
            set_fill_alpha(sh, transparency)
        if line is None:
            sh.line.fill.background()
        else:
            sh.line.color.rgb = RGBColor.from_string(self.theme.color(line))
            sh.line.width = self._p(line_w)
        add_shadow(sh) if shadow else clear_shadow(sh)
        sh.text_frame.word_wrap = True
        return sh

    def line(self, x1: float, y1: float, x2: float, y2: float,
             color="hairline", width=1.0, dash=None):
        from pptx.enum.shapes import MSO_CONNECTOR
        cn = self.raw.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                          self._i(x1), self._i(y1),
                                          self._i(x2), self._i(y2))
        cn.line.color.rgb = RGBColor.from_string(self.theme.color(color))
        cn.line.width = self._p(width)
        if dash:
            from pptx.enum.dml import MSO_LINE_DASH_STYLE
            cn.line.dash_style = getattr(MSO_LINE_DASH_STYLE, dash, None)
        return cn

    # ---- 文本 ----
    def text(self, at: Rect, content, size=None, bold=False, color="ink",
             align="left", anchor="top", line_spacing=1.28, italic=False,
             font_en=None, font_cn=None, fit=False, min_pt=8.0,
             space_after=0.0, shrink_wrap=False, vert=None):
        """放一个文本框。

        content 可以是 str，也可以是 [(文字, 覆盖属性dict), ...] 的富文本片段列表，
        用于在一行里混排不同字号/颜色（例如把数字染成 accent）。
        fit=True 时按区域高度自动降档字号。
        vert='vert270' 竖排（从下往上读，用于纵轴标签）；'vert' 为从上往下。
        """
        theme = self.theme
        size = theme.size_body if size is None else size
        box = self.raw.shapes.add_textbox(self._i(at.x), self._i(at.y),
                                          self._i(at.w), self._i(at.h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                              "bottom": MSO_ANCHOR.BOTTOM}[anchor]
        if vert:
            # 真正的文字方向旋转，而不是靠窄文本框逐字折行来伪造竖排
            tf._txBody.bodyPr.set("vert", vert)

        if fit and isinstance(content, str):
            size = fit_size(content, at.w, at.h, size, min_pt=min_pt,
                            line_spacing=line_spacing)

        paras = content if isinstance(content, list) else [(content, {})]
        # 纯字符串按换行拆段；富文本列表整体放一段
        if isinstance(content, str):
            paras = [(seg, {}) for seg in content.split("\n")] or [("", {})]
            multi_para = True
        else:
            multi_para = False

        first = True
        p = tf.paragraphs[0]
        for seg, over in paras:
            if multi_para and not first:
                p = tf.add_paragraph()
            p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                           "right": PP_ALIGN.RIGHT,
                           "justify": PP_ALIGN.JUSTIFY}[over.get("align", align)]
            p.line_spacing = over.get("line_spacing", line_spacing)
            if space_after:
                p.space_after = self._p(space_after)
            run = p.add_run()
            run.text = str(seg)
            apply_font(run.font, theme,
                       size=(over.get("size", size)) * self.deck.scale,
                       bold=over.get("bold", bold),
                       italic=over.get("italic", italic),
                       color=over.get("color", color),
                       font_en=over.get("font_en", font_en),
                       font_cn=over.get("font_cn", font_cn))
            first = False

        if shrink_wrap:
            box.height = self._i(min(at.h, text_height(
                content if isinstance(content, str) else
                "".join(str(s) for s, _ in paras), at.w, size, line_spacing) + 0.06))
        return box

    # ---- 页面骨架 ----
    def title(self, text, eyebrow=None, sub=None, color="ink", rule=False):
        """页面标题。

        写"行动标题"：把结论写进标题，而不是写栏目名。
        eyebrow 是标题上方的小字栏目标签，用来交代这一页在结构中的位置。
        """
        theme = self.theme
        m = theme.margin
        y = m
        w = BASE_W - 2 * m
        if eyebrow:
            self.text(Rect(m, y, w, 0.24), eyebrow, size=theme.size_eyebrow,
                      bold=True, color="accent", anchor="top")
            y += 0.30
        size = fit_size(str(text), w, 1.0, theme.size_title, min_pt=15)
        h = text_height(str(text), w, size, 1.18) + 0.04
        self.text(Rect(m, y, w, h), text, size=size, bold=True, color=color,
                  line_spacing=1.18, anchor="top")
        y += h + 0.10
        if sub:
            sh = text_height(str(sub), w, theme.size_small, 1.3) + 0.02
            self.text(Rect(m, y, w, sh), sub, size=theme.size_small,
                      color="ink_muted", line_spacing=1.3)
            y += sh + 0.08
        if rule:
            self.line(m, y, BASE_W - m, y, "hairline", 0.75)
            y += 0.12
        self._title_bottom = y + 0.06
        return self

    def note(self, text, color="ink_muted"):
        """页脚注释/数据来源。高层汇报里的数字必须能溯源。"""
        m = self.theme.margin
        self.text(Rect(m, BASE_H - m + 0.04, BASE_W - 2 * m, 0.24), text,
                  size=self.theme.size_note, color=color, anchor="top")
        return self

    def page_number(self, n, total=None):
        txt = "%d" % n if total is None else "%d / %d" % (n, total)
        self.text(Rect(BASE_W - self.theme.margin - 1.0,
                       BASE_H - self.theme.margin + 0.04, 1.0, 0.24), txt,
                  size=self.theme.size_note, color="ink_muted", align="right")
        return self

    # ------------------------------------------------------------------
    # 组件
    # ------------------------------------------------------------------

    def kpi_row(self, items, at: Rect | None = None, cols=None, gap=None,
                card=True, height=None):
        """KPI 大数字卡片行 —— 高层汇报第一屏的标准配置。

        items: [{value, label, unit?, delta?, trend?('up'/'down'/'flat'), note?, color?}]
        trend 决定同比箭头的语义颜色；如果"下降是好事"（如能耗、偏差数），
        直接传 color='good' 覆盖。
        """
        theme = self.theme
        at = at or self.body
        gap = theme.gap if gap is None else gap
        n = len(items)
        cols = cols or n
        rows = math.ceil(n / cols)
        # 不给高度时用固定高度而非撑满 body：KPI 卡拉得过高会在卡内留出大片空白
        h = min(at.h, 1.95 * rows + gap * (rows - 1)) if height is None else height
        cells = Rect(at.x, at.y, at.w, h).grid(cols, rows, gap)
        out = []
        for it, cell in zip(items, cells):
            if card:
                self.rect(cell, fill="surface", shadow=False)
            inner = cell.inset(0.22, 0.18)
            val = str(it.get("value", ""))
            unit = it.get("unit", "")
            vcolor = it.get("color", "ink")
            # 数字尽量大，但必须装得下
            vsize = fit_size(val + unit, inner.w, inner.h * 0.55,
                             theme.size_kpi, min_pt=16)
            vh = vsize * 1.15 / 72.0

            # 先量出内容总高，再整体垂直居中 —— 否则内容贴顶、卡片下半空着
            blocks = [vh, 0.06, 0.24]
            delta = it.get("delta")
            if delta:
                blocks.append(0.26)
            note_h = 0.0
            if it.get("note"):
                note_h = text_height(it["note"], inner.w, theme.size_note, 1.3)
                blocks.append(0.04 + note_h)
            total = sum(blocks)
            y = inner.y + max(0.0, (inner.h - total) / 2)

            segs = [(val, {"size": vsize, "bold": True, "color": vcolor,
                           "font_en": theme.font_num})]
            if unit:
                segs.append((" " + unit, {"size": max(vsize * 0.42, 9),
                                          "bold": False, "color": "ink_muted"}))
            self.text(Rect(inner.x, y, inner.w, vh), segs, anchor="bottom")
            y += vh + 0.06
            self.text(Rect(inner.x, y, inner.w, 0.24), it.get("label", ""),
                      size=theme.size_kpi_label, color="ink_muted", anchor="top")
            y += 0.24
            if delta:
                trend = it.get("trend", "flat")
                arrow = {"up": "▲", "down": "▼", "flat": "▬"}.get(trend, "")
                dcolor = it.get("delta_color") or {
                    "up": "good", "down": "bad", "flat": "neutral"}.get(trend, "neutral")
                self.text(Rect(inner.x, y, inner.w, 0.26),
                          [(arrow + " ", {"size": theme.size_small * 0.85,
                                          "color": dcolor}),
                           (str(delta), {"size": theme.size_small, "bold": True,
                                         "color": dcolor})], anchor="top")
                y += 0.26
            if note_h:
                self.text(Rect(inner.x, y + 0.04, inner.w, note_h),
                          it["note"], size=theme.size_note, color="ink_muted",
                          line_spacing=1.3, anchor="top")
            out.append(cell)
        return out

    def bullets(self, items, at: Rect | None = None, size=None, marker="dot",
                gap=0.14, color="ink", bold_head=True, marker_color="accent",
                desc_color="ink_muted"):
        """要点列表。items 元素可以是 str，或 (标题, 说明) 二元组。

        marker: 'dot' 圆点 / 'num' 序号 / 'none'
        用 (标题, 说明) 形式能形成"粗体结论 + 灰色解释"的两级结构，比纯 bullet 好读。
        """
        theme = self.theme
        at = at or self.body
        size = theme.size_body if size is None else size
        y = at.y
        line_h = size * 1.3 / 72.0
        for i, it in enumerate(items, 1):
            head, desc = (it, None) if isinstance(it, str) else (it[0], it[1])
            mx = at.x
            tx = at.x
            if marker == "dot":
                d = size / 72.0 * 0.34
                # 圆点对齐首行文字的垂直中心，否则会显得浮在文字上方
                self.rect(Rect(mx, y + (line_h - d) / 2, d, d), fill=marker_color,
                          shape=MSO_SHAPE.OVAL, radius=0)
                tx = mx + d + 0.14
            elif marker == "num":
                self.text(Rect(mx, y, 0.38, line_h), "%02d" % i,
                          size=size * 0.9, bold=True, color=marker_color,
                          anchor="middle")
                tx = mx + 0.46
            tw = at.right - tx
            hh = text_height(str(head), tw, size, 1.3)
            self.text(Rect(tx, y, tw, hh), head, size=size, bold=bold_head,
                      color=color, line_spacing=1.3)
            y += hh + 0.04
            if desc:
                dsize = size * 0.92
                dh = text_height(str(desc), tw, dsize, 1.34)
                self.text(Rect(tx, y, tw, dh), desc, size=dsize,
                          color=desc_color, line_spacing=1.34)
                y += dh
            y += gap
            if y > at.bottom:
                break
        return y

    def icon_cards(self, items, at: Rect | None = None, cols=3, gap=None,
                   icon_size=0.46):
        """图标卡片组。items: [{icon(单字符或emoji), title, desc, color?}]

        icon 用文字符号（■ ● ▲ ✓ 数字 等）画在色圆里，不依赖外部图片，
        跨环境不会缺图。要用真图标见 deck-imagery skill。
        """
        theme = self.theme
        at = at or self.body
        gap = theme.gap if gap is None else gap
        rows = math.ceil(len(items) / cols)
        cells = at.grid(cols, rows, gap)
        for it, cell in zip(items, cells):
            self.rect(cell, fill="surface")
            inner = cell.inset(0.22, 0.2)
            c = it.get("color", "primary")
            # 先量总高再整体垂直居中，卡片被拉高时内容不会贴在顶部
            th = text_height(it.get("title", ""), inner.w, theme.size_h, 1.24)
            dsize = theme.size_small
            dh = 0.0
            if it.get("desc"):
                dsize = fit_size(it["desc"], inner.w,
                                 max(inner.h - icon_size - th - 0.3, 0.3),
                                 theme.size_small, min_pt=9, line_spacing=1.34)
                dh = text_height(it["desc"], inner.w, dsize, 1.34)
            total = icon_size + 0.14 + th + (0.07 + dh if dh else 0)
            y = inner.y + max(0.0, (inner.h - total) / 2)

            self.rect(Rect(inner.x, y, icon_size, icon_size), fill=c,
                      shape=MSO_SHAPE.OVAL, radius=0)
            self.text(Rect(inner.x, y + 0.02, icon_size, icon_size - 0.04),
                      it.get("icon", "●"), size=icon_size * 72 * 0.42, bold=True,
                      color=theme.ink_on(c), align="center", anchor="middle")
            y += icon_size + 0.14
            self.text(Rect(inner.x, y, inner.w, th), it.get("title", ""),
                      size=theme.size_h, bold=True, line_spacing=1.24)
            y += th + 0.07
            if dh:
                self.text(Rect(inner.x, y, inner.w, dh), it["desc"],
                          size=dsize, color="ink_muted", line_spacing=1.34)
        return cells

    def callout(self, text, at: Rect | None = None, kind="accent",
                label=None, size=None, icon=None):
        """结论条 / 强调块。每页最多一个，否则重点就不是重点了。

        kind: accent(结论) / good(达成) / warn(关注) / bad(风险) / surface(中性)
        """
        theme = self.theme
        at = at or Rect(theme.margin, BASE_H - theme.margin - 0.86,
                        BASE_W - 2 * theme.margin, 0.86)
        fill = kind if kind in ("accent", "good", "warn", "bad", "primary",
                                "secondary") else "surface"
        self.rect(at, fill=fill)
        inner = at.inset(0.24, 0.16)
        ink = theme.block_ink(fill)[0]
        x = inner.x
        if icon:
            self.text(Rect(x, inner.y, 0.36, inner.h), icon,
                      size=(size or theme.size_h) * 1.1, bold=True, color=ink,
                      anchor="middle")
            x += 0.44
        if label:
            lw = text_width(label, theme.size_small) + 0.06
            self.text(Rect(x, inner.y, lw, inner.h), label,
                      size=theme.size_small, bold=True, color=ink,
                      anchor="middle")
            x += lw + 0.18
            self.line(x - 0.09, inner.y + 0.06, x - 0.09, inner.bottom - 0.06,
                      ink, 0.75)
        self.text(Rect(x, inner.y, inner.right - x, inner.h), text,
                  size=size or theme.size_body, bold=True, color=ink,
                  anchor="middle", line_spacing=1.26, fit=True, min_pt=9)
        return at

    def table(self, headers, rows, at: Rect | None = None, col_widths=None,
              align=None, row_h=None, header_h=None, cell_colors=None,
              size=None, zebra=True):
        """数据表。

        align: 每列对齐方式列表，如 ['left','right','right','center']
               数字列一律右对齐，否则位数对不齐、没法快速比较。
        cell_colors: {(行,列): 'good'} 形式的单元格文字染色，用于标红/标绿。
        """
        theme = self.theme
        at = at or self.body
        size = (theme.size_small if size is None else size)
        ncol = len(headers)
        nrow = len(rows)
        header_h = header_h or 0.34
        row_h = row_h or min(0.40, max(0.26, (at.h - header_h) / max(nrow, 1)))
        total_h = header_h + row_h * nrow
        gf = self.raw.shapes.add_table(nrow + 1, ncol, self._i(at.x),
                                       self._i(at.y), self._i(at.w),
                                       self._i(min(total_h, at.h)))
        tbl = gf.table
        # 关掉 PowerPoint 自带的花哨样式，自己控制
        tblPr = tbl._tbl.find(qn("a:tblPr"))
        if tblPr is not None:
            tblPr.set("firstRow", "1")
            tblPr.set("bandRow", "0")
        if col_widths:
            tot = sum(col_widths)
            for i, cw in enumerate(col_widths):
                tbl.columns[i].width = self._i(at.w * cw / tot)
        tbl.rows[0].height = self._i(header_h)
        for r in range(nrow):
            tbl.rows[r + 1].height = self._i(row_h)

        aligns = align or (["left"] + ["right"] * (ncol - 1))
        aligns = (aligns + ["left"] * ncol)[:ncol]
        amap = {"left": PP_ALIGN.LEFT, "right": PP_ALIGN.RIGHT,
                "center": PP_ALIGN.CENTER}

        def style_cell(cell, txt, *, bold, fill, fg, al, fsize):
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(theme.color(fill))
            cell.margin_left = cell.margin_right = self._i(0.09)
            cell.margin_top = cell.margin_bottom = self._i(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = amap[al]
            run = p.add_run()
            run.text = "" if txt is None else str(txt)
            apply_font(run.font, theme, size=fsize * self.deck.scale,
                       bold=bold, color=fg)

        for c, htxt in enumerate(headers):
            style_cell(tbl.cell(0, c), htxt, bold=True, fill="surface_alt",
                       fg="ink", al=aligns[c], fsize=size)
        for r, row in enumerate(rows):
            fill = "surface" if (not zebra or r % 2 == 0) else "bg"
            for c in range(ncol):
                val = row[c] if c < len(row) else ""
                fg = (cell_colors or {}).get((r, c), "ink")
                style_cell(tbl.cell(r + 1, c), val, bold=(c == 0), fill=fill,
                           fg=fg, al=aligns[c], fsize=size)
        return tbl

    def progress_bars(self, items, at: Rect | None = None, size=None,
                      bar_h=0.20, gap=0.30, show_target=True, target=100.0,
                      align="middle"):
        """达成率条组。items: [{label, value(数值，通常 0-120), text?, color?}]

        比饼图更适合"多个指标 vs 目标"：条形起点对齐，一眼看出谁没达标。

        条长按实际最大值缩放，不硬截断到 100 —— 截断会把 103% 和 108% 画成一样长，
        读者没法比较。target 处画一条参考线标出达标位置。
        align: middle 垂直居中 / top 贴顶 / fill 撑满区域
        """
        theme = self.theme
        at = at or self.body
        size = theme.size_body if size is None else size
        vals = [float(it.get("value", 0) or 0) for it in items]
        vmax = max([target] + vals) * 1.04
        label_w = max([text_width(str(i.get("label", "")), size) for i in items]
                      + [1.0])
        label_w = min(label_w + 0.14, at.w * 0.40)
        val_w = max([text_width(str(i.get("text") or "100%"), size)
                     for i in items] + [0.6]) + 0.18
        track_x = at.x + label_w + 0.16
        track_w = at.right - val_w - 0.12 - track_x
        row_h = bar_h + gap
        n = len(items)
        if align == "fill" and n > 1:
            step = at.h / n
            y = at.y + (step - bar_h) / 2
        else:
            step = row_h
            span = row_h * n - gap
            y = at.y + (max(0.0, at.h - span) / 2 if align == "middle" else 0.0)

        if show_target:
            tx = track_x + track_w * min(target, vmax) / vmax
            top = min(y - 0.10, at.y)
            self.line(tx, top + 0.22, tx, y + step * (n - 1) + bar_h + 0.10,
                      "ink_muted", 1.0, dash="DASH")
            self.text(Rect(tx - 0.6, top - 0.04, 1.2, 0.24),
                      "目标 %g%%" % target, size=theme.size_note,
                      color="ink_muted", align="center")
        for it, v in zip(items, vals):
            c = it.get("color") or ("good" if v >= target else
                                    "warn" if v >= target * 0.9 else "bad")
            self.text(Rect(at.x, y - 0.04, label_w, bar_h + 0.08),
                      it.get("label", ""), size=size, color="ink",
                      anchor="middle")
            self.rect(Rect(track_x, y, track_w, bar_h), fill="surface_alt",
                      radius=0.5)
            fw = track_w * max(0.0, v) / vmax
            if fw > 0.02:
                self.rect(Rect(track_x, y, fw, bar_h), fill=c, radius=0.5)
            self.text(Rect(at.right - val_w, y - 0.04, val_w, bar_h + 0.08),
                      it.get("text") or ("%.0f%%" % v), size=size, bold=True,
                      color=c, align="right", anchor="middle")
            y += step
        return at

    def timeline(self, items, at: Rect | None = None, orient="h", size=None):
        """里程碑时间轴。items: [{when, title, desc?, status?('done'/'doing'/'todo')}]"""
        theme = self.theme
        at = at or self.body
        size = theme.size_body if size is None else size
        smap = {"done": "good", "doing": "accent", "todo": "neutral"}
        n = len(items)
        if orient == "h":
            axis_y = at.y + at.h * 0.42
            self.line(at.x + 0.1, axis_y, at.right - 0.1, axis_y, "hairline", 1.5)
            cw = at.w / n
            for i, it in enumerate(items):
                cx = at.x + cw * (i + 0.5)
                c = smap.get(it.get("status", "todo"), "neutral")
                d = 0.2
                self.rect(Rect(cx - d / 2, axis_y - d / 2, d, d), fill=c,
                          shape=MSO_SHAPE.OVAL, radius=0)
                self.text(Rect(cx - cw / 2 + 0.06, at.y, cw - 0.12,
                               at.h * 0.42 - 0.16), it.get("when", ""),
                          size=theme.size_small, bold=True, color=c,
                          align="center", anchor="bottom")
                ty = axis_y + 0.20
                th = text_height(it.get("title", ""), cw - 0.12, size, 1.24)
                self.text(Rect(cx - cw / 2 + 0.06, ty, cw - 0.12, th),
                          it.get("title", ""), size=size, bold=True,
                          align="center", line_spacing=1.24)
                if it.get("desc"):
                    self.text(Rect(cx - cw / 2 + 0.06, ty + th + 0.05,
                                   cw - 0.12, at.bottom - ty - th - 0.05),
                              it["desc"], size=theme.size_note, color="ink_muted",
                              align="center", line_spacing=1.3, fit=True)
        else:
            rh = at.h / n
            ax = at.x + 0.14
            self.line(ax, at.y + 0.08, ax, at.bottom - 0.08, "hairline", 1.5)
            for i, it in enumerate(items):
                cy = at.y + rh * (i + 0.5)
                c = smap.get(it.get("status", "todo"), "neutral")
                d = 0.18
                self.rect(Rect(ax - d / 2, cy - d / 2, d, d), fill=c,
                          shape=MSO_SHAPE.OVAL, radius=0)
                tx = ax + 0.30
                self.text(Rect(tx, cy - rh / 2 + 0.02, 0.95, rh - 0.04),
                          it.get("when", ""), size=theme.size_small, bold=True,
                          color=c, anchor="middle")
                bx = tx + 1.02
                self.text(Rect(bx, cy - rh / 2 + 0.02, at.right - bx, rh - 0.04),
                          it.get("title", ""), size=size, bold=True,
                          anchor="middle")
        return at

    def compare(self, left, right, at: Rect | None = None, gap=None):
        """左右对比块。left/right: {title, items[], color?, note?}

        用于"目标 vs 实际""改善前 vs 改善后""方案 A vs 方案 B"。
        """
        theme = self.theme
        at = at or self.body
        gap = theme.gap if gap is None else gap
        a, b = at.split_h(1, 1, gap=gap)
        for panel, cell in ((left, a), (right, b)):
            c = panel.get("color", "primary")
            self.rect(cell, fill="surface")
            head_h = 0.44
            self.rect(Rect(cell.x, cell.y, cell.w, head_h), fill=c,
                      radius=theme.radius)
            self.rect(Rect(cell.x, cell.y + head_h - 0.1, cell.w, 0.1), fill=c,
                      radius=0)
            self.text(Rect(cell.x + 0.2, cell.y, cell.w - 0.4, head_h),
                      panel.get("title", ""), size=theme.size_h, bold=True,
                      color=theme.ink_on(c), anchor="middle")
            inner = Rect(cell.x + 0.2, cell.y + head_h + 0.16,
                         cell.w - 0.4, cell.h - head_h - 0.34)
            self.bullets(panel.get("items", []), inner, size=theme.size_body,
                         bold_head=False, gap=0.16)
            if panel.get("note"):
                self.text(Rect(inner.x, cell.bottom - 0.34, inner.w, 0.26),
                          panel["note"], size=theme.size_note, color="ink_muted")
        return [a, b]

    def matrix(self, quadrants, at: Rect | None = None, x_label=None,
               y_label=None, gap=0.18):
        """2x2 四象限。quadrants 按 [左上, 右上, 左下, 右下] 顺序，
        元素为 {title, items[]?, desc?, color?}。用于风险矩阵、优先级排序。

        x_label / y_label 可以是一个字符串，也可以是 (低端, 高端) 二元组 ——
        二元组会把两个标签摆到轴的两端，比在一个字符串里塞空格可靠。
        """
        theme = self.theme
        at = at or self.body
        if x_label or y_label:
            at = at.pad(left=0.34 if y_label else 0,
                        bottom=0.32 if x_label else 0)
        cells = at.grid(2, 2, gap)
        for q, cell in zip(quadrants, cells):
            c = q.get("color", "surface")
            self.rect(cell, fill=c)
            inner = cell.inset(0.18, 0.14)
            ink, ink2 = theme.block_ink(c)
            self.text(Rect(inner.x, inner.y, inner.w, 0.28), q.get("title", ""),
                      size=theme.size_h, bold=True, color=ink)
            rest = Rect(inner.x, inner.y + 0.34, inner.w,
                        inner.bottom - inner.y - 0.34)
            if q.get("items"):
                self.bullets(q["items"], rest, size=theme.size_small,
                             bold_head=False, gap=0.06, color=ink,
                             marker_color=ink, desc_color=ink2)
            elif q.get("desc"):
                self.text(rest, q["desc"], size=theme.size_small, color=ink2,
                          line_spacing=1.34, fit=True)
        if x_label:
            row = Rect(at.x, at.bottom + 0.06, at.w, 0.24)
            if isinstance(x_label, (tuple, list)):
                self.text(row, "← " + str(x_label[0]), size=theme.size_note,
                          color="ink_muted", align="left")
                self.text(row, str(x_label[1]) + " →", size=theme.size_note,
                          color="ink_muted", align="right")
            else:
                self.text(row, str(x_label), size=theme.size_note,
                          color="ink_muted", align="center")
        if y_label:
            col = Rect(at.x - 0.34, at.y, 0.26, at.h)
            if isinstance(y_label, (tuple, list)):
                self.text(col, "← " + str(y_label[0]), size=theme.size_note,
                          color="ink_muted", align="left", vert="vert270")
                self.text(col, str(y_label[1]) + " →", size=theme.size_note,
                          color="ink_muted", align="right", vert="vert270")
            else:
                self.text(col, str(y_label), size=theme.size_note,
                          color="ink_muted", align="center", vert="vert270")
        return cells

    def image(self, path, at: Rect | None = None, mode="cover", radius=None,
              caption=None):
        """插图。mode='cover' 按区域裁剪填满（不变形），'fit' 完整放入。

        变形的照片在高层汇报里非常显眼，所以默认走 cover 裁剪而不是拉伸。
        """
        at = at or self.body
        try:
            from PIL import Image
            with Image.open(path) as im:
                iw, ih = im.size
        except Exception:
            iw = ih = None
        x, y, w, h = at.x, at.y, at.w, at.h
        pic = self.raw.shapes.add_picture(path, self._i(x), self._i(y),
                                          self._i(w), self._i(h))
        if iw and ih:
            box_ar, img_ar = w / h, iw / ih
            if mode == "cover":
                # 用 crop 保持比例填满，避免拉伸
                if img_ar > box_ar:
                    over = 1 - box_ar / img_ar
                    pic.crop_left = pic.crop_right = over / 2
                else:
                    over = 1 - img_ar / box_ar
                    pic.crop_top = pic.crop_bottom = over / 2
            else:
                if img_ar > box_ar:
                    nh = w / img_ar
                    pic.top = self._i(y + (h - nh) / 2)
                    pic.height = self._i(nh)
                else:
                    nw = h * img_ar
                    pic.left = self._i(x + (w - nw) / 2)
                    pic.width = self._i(nw)
        if caption:
            self.text(Rect(at.x, at.bottom + 0.05, at.w, 0.24), caption,
                      size=self.theme.size_note, color="ink_muted")
        return pic

    # ---- 图表 ----
    def chart(self, kind, categories, series, at: Rect | None = None,
              legend=None, data_labels=None, number_format=None,
              colors=None, gap_width=60, overlap=None, smooth=False,
              value_axis=True, category_axis=True, max_scale=None,
              min_scale=None, label_position=None, title=None):
        """原生 PowerPoint 图表（可在 PPT 里继续编辑数据，别用图片代替）。

        kind: bar(纵向柱) / hbar(横向条) / stacked / hstacked / line / area /
              pie / doughnut / scatter_line / radar
        series: [(名称, [数值...]), ...]
        默认样式：无边框、无标题、淡网格线、主题配色、单序列自动隐藏图例。
        """
        theme = self.theme
        at = at or self.body
        ctype = {
            "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "hbar": XL_CHART_TYPE.BAR_CLUSTERED,
            "stacked": XL_CHART_TYPE.COLUMN_STACKED,
            "hstacked": XL_CHART_TYPE.BAR_STACKED,
            "stacked100": XL_CHART_TYPE.COLUMN_STACKED_100,
            "hstacked100": XL_CHART_TYPE.BAR_STACKED_100,
            "line": XL_CHART_TYPE.LINE,
            "line_markers": XL_CHART_TYPE.LINE_MARKERS,
            "area": XL_CHART_TYPE.AREA,
            "pie": XL_CHART_TYPE.PIE,
            "doughnut": XL_CHART_TYPE.DOUGHNUT,
            "radar": XL_CHART_TYPE.RADAR,
        }.get(kind, XL_CHART_TYPE.COLUMN_CLUSTERED)

        # 横向条形图 PowerPoint 从下往上排，直接传入会让第一项落到最底部。
        # 阅读习惯是"最重要/最大的在最上面"，所以先把顺序倒过来。
        if kind in ("hbar", "hstacked", "hstacked100"):
            categories = list(categories)[::-1]
            series = [(nm, list(v)[::-1]) for nm, v in series]

        cd = CategoryChartData()
        cd.categories = list(categories)
        for name, vals in series:
            cd.add_series(name, tuple(vals))
        gf = self.raw.shapes.add_chart(ctype, self._i(at.x), self._i(at.y),
                                       self._i(at.w), self._i(at.h), cd)
        ch = gf.chart

        # 全局字体（含东亚字体，否则图表里的中文会串字体）
        apply_font(ch.font, theme, size=theme.size_small * self.deck.scale,
                   color="ink_muted")

        ch.has_title = bool(title)
        if title:
            ch.chart_title.text_frame.text = title
            apply_font(ch.chart_title.text_frame.paragraphs[0].font, theme,
                       size=theme.size_h * self.deck.scale, bold=True, color="ink")

        multi = len(series) > 1
        show_legend = multi if legend is None else legend
        ch.has_legend = bool(show_legend)
        if ch.has_legend:
            ch.legend.position = XL_LEGEND_POSITION.TOP
            ch.legend.include_in_layout = False
            apply_font(ch.legend.font, theme,
                       size=theme.size_small * self.deck.scale, color="ink_muted")

        pal = [theme.color(c) for c in (colors or theme.series)]
        is_pie = kind in ("pie", "doughnut")
        plot = ch.plots[0]
        if is_pie:
            # 饼图按数据点着色
            pts = plot.series[0].points
            for i, pt in enumerate(pts):
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = RGBColor.from_string(pal[i % len(pal)])
                pt.format.line.color.rgb = RGBColor.from_string(theme.color(theme.bg))
                pt.format.line.width = self._p(1.2)
        else:
            for i, s in enumerate(ch.series):
                col = RGBColor.from_string(pal[i % len(pal)])
                if kind in ("line", "line_markers", "radar"):
                    s.format.line.color.rgb = col
                    s.format.line.width = self._p(2.4)
                    s.smooth = smooth
                else:
                    s.format.fill.solid()
                    s.format.fill.fore_color.rgb = col
                    s.format.line.fill.background()

        # 数据标签：柱/条/饼默认开，折线默认关（否则容易糊成一片）
        if data_labels is None:
            data_labels = kind not in ("line", "line_markers", "area", "radar")
        stacked = kind in ("stacked", "hstacked", "stacked100", "hstacked100")
        plot.has_data_labels = bool(data_labels)
        if plot.has_data_labels:
            dl = plot.data_labels
            apply_font(dl.font, theme, size=theme.size_small * self.deck.scale,
                       bold=True, color="ink")
            if number_format:
                dl.number_format = number_format
                dl.number_format_is_linked = False
            if not is_pie:
                pos = label_position or ("inEnd" if stacked else "outEnd")
                # 堆积图用 outEnd 会让 PowerPoint 判定文件损坏
                if stacked and pos == "outEnd":
                    pos = "inEnd"
                try:
                    dl.position = {"outEnd": XL_LABEL_POSITION.OUTSIDE_END,
                                   "inEnd": XL_LABEL_POSITION.INSIDE_END,
                                   "ctr": XL_LABEL_POSITION.CENTER,
                                   "inBase": XL_LABEL_POSITION.INSIDE_BASE}[pos]
                except (KeyError, ValueError):
                    pass
            else:
                try:
                    dl.position = XL_LABEL_POSITION.OUTSIDE_END
                except ValueError:
                    pass       # 环形图不支持外侧标签，退回默认位置

            # 标签落在色块内时逐序列/逐点重设文字色 —— 统一用白字会在金色、
            # 浅蓝这类亮色段上完全读不出来。
            if stacked:
                for i, sr in enumerate(ch.series):
                    apply_font(sr.data_labels.font, theme,
                               size=theme.size_small * self.deck.scale, bold=True,
                               color=theme.ink_on(pal[i % len(pal)]))
                    if number_format:
                        sr.data_labels.number_format = number_format
                        sr.data_labels.number_format_is_linked = False
            elif kind == "doughnut":
                for i, pt in enumerate(plot.series[0].points):
                    apply_font(pt.data_label.font, theme,
                               size=theme.size_small * self.deck.scale, bold=True,
                               color=theme.ink_on(pal[i % len(pal)]))

        if not is_pie:
            if hasattr(plot, "gap_width"):
                plot.gap_width = gap_width
            if overlap is not None and hasattr(plot, "overlap"):
                plot.overlap = overlap
            elif kind in ("stacked", "hstacked", "stacked100") and hasattr(plot, "overlap"):
                plot.overlap = 100
            va, ca = ch.value_axis, ch.category_axis
            va.visible = bool(value_axis)
            ca.visible = bool(category_axis)
            va.has_major_gridlines = True
            gl = va.major_gridlines.format.line
            gl.color.rgb = RGBColor.from_string(theme.color(theme.hairline))
            gl.width = self._p(0.5)
            ca.has_major_gridlines = False
            for ax in (va, ca):
                ax.format.line.color.rgb = RGBColor.from_string(theme.color(theme.hairline))
                ax.format.line.width = self._p(0.75)
                apply_font(ax.tick_labels.font, theme,
                           size=theme.size_small * self.deck.scale,
                           color="ink_muted")
            if number_format:
                va.tick_labels.number_format = number_format
                va.tick_labels.number_format_is_linked = False
            # 百分比堆积图各类别合计恒为 100，让轴自动到 120 会留一段空白
            if max_scale is None and kind in ("stacked", "hstacked"):
                totals = [sum(vals[i] for _, vals in series if i < len(vals))
                          for i in range(len(categories))]
                if totals and all(abs(t - 100) < 0.51 for t in totals):
                    max_scale = 100
            if max_scale is not None:
                va.maximum_scale = max_scale
            if min_scale is not None:
                va.minimum_scale = min_scale
        else:
            if kind == "doughnut":
                try:
                    plot.hole_size = 55
                except Exception:
                    pass

        # 图表区与绘图区透明，融进页面底色
        for target in (ch, plot):
            try:
                target.format.fill.background()
                target.format.line.fill.background()
            except Exception:
                pass
        return ch

    def sparkline(self, values, at: Rect, color="accent", width=1.8,
                  fill_area=False):
        """极简折线（用形状画，不建图表对象）。用于 KPI 卡里的迷你趋势。"""
        if len(values) < 2:
            return None
        lo, hi = min(values), max(values)
        span = (hi - lo) or 1.0
        pts = [(at.x + at.w * i / (len(values) - 1),
                at.bottom - at.h * (v - lo) / span) for i, v in enumerate(values)]
        for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
            self.line(x1, y1, x2, y2, color, width)
        return pts


# --------------------------------------------------------------------------
# Deck
# --------------------------------------------------------------------------

class Deck:
    """一份演示文稿。"""

    def __init__(self, theme="dark", canvas="wide", template=None):
        self.prs = Presentation(template) if template else Presentation()
        if isinstance(canvas, str):
            cw, chh = CANVAS.get(canvas, CANVAS["wide"])
        else:
            cw, chh = canvas
        if template is None:
            self.prs.slide_width = Inches(cw)
            self.prs.slide_height = Inches(chh)
        else:
            cw = self.prs.slide_width / 914400
            chh = self.prs.slide_height / 914400
        self.canvas_w, self.canvas_h = cw, chh
        # 设计单位 -> 实际画布的缩放系数，组件因此与画布尺寸解耦
        self.scale = cw / BASE_W
        self.theme = THEMES.get(theme, DARK) if isinstance(theme, str) else theme
        self._blank = self._pick_blank_layout()
        self.slides: list[Slide] = []
        self._sync_theme_fonts()

    def _pick_blank_layout(self):
        for lo in self.prs.slide_layouts:
            if (lo.name or "").strip().lower() in ("blank", "空白"):
                return lo
        return self.prs.slide_layouts[-1]

    def _sync_theme_fonts(self):
        """把主题字体写进 theme part，这样用户后续手动加的文本框也是对的字体。"""
        try:
            from pptx.opc.constants import RELATIONSHIP_TYPE as RT
            part = self.prs.slide_masters[0].part.part_related_by(RT.THEME)
            xml = part.blob.decode("utf-8")
            t = self.theme
            xml = re.sub(r'(<a:(?:majorFont|minorFont)>\s*<a:latin[^/]*?typeface=")[^"]*(")',
                         lambda m: m.group(1) + t.font_en + m.group(2), xml)
            xml = re.sub(r'(<a:ea\s+typeface=")[^"]*(")',
                         lambda m: m.group(1) + t.font_cn + m.group(2), xml)
            part._blob = xml.encode("utf-8")
        except Exception:
            pass  # 拿不到 theme part 不影响正常输出，元素字体都是显式写的

    # ---- 页面 ----
    def slide(self, title=None, eyebrow=None, sub=None, bg=None, rule=False):
        s = Slide(self, self.prs.slides.add_slide(self._blank),
                  bg or self.theme.bg)
        if title:
            s.title(title, eyebrow=eyebrow, sub=sub, rule=rule)
        self.slides.append(s)
        return s

    def cover(self, title, subtitle=None, meta=None, kicker=None,
              image=None, accent_block=True, scrim=52):
        """封面。深色底 + 大标题 + 一条强调色块。

        传 image 时会压一层 scrim% 的半透明底色，保证标题可读；图案本身透出来。
        照片类背景可能需要更高的 scrim（60-70）。
        """
        t = self.theme
        s = Slide(self, self.prs.slides.add_slide(self._blank), t.bg_alt)
        self.slides.append(s)
        if image:
            s.bg_image(image)
            s.rect(s.page, fill=t.bg_alt, radius=0, transparency=100 - scrim)
        m = t.margin + 0.24
        box = Rect(m, 0, BASE_W - 2 * m, BASE_H)
        y = BASE_H * 0.30
        if kicker:
            s.text(Rect(m, y - 0.42, box.w, 0.30), kicker, size=t.size_eyebrow + 1,
                   bold=True, color="accent")
        if accent_block:
            s.rect(Rect(m, y, 0.09, 0.62), fill="accent", radius=0)
            tx = m + 0.30
        else:
            tx = m
        tsize = fit_size(str(title), box.w - (tx - m), 1.6, t.size_cover_title,
                         min_pt=22)
        th = text_height(str(title), box.w - (tx - m), tsize, 1.16)
        s.text(Rect(tx, y, box.w - (tx - m), th), title, size=tsize, bold=True,
               color="ink", line_spacing=1.16)
        y += th + 0.22
        if subtitle:
            sh = text_height(str(subtitle), box.w * 0.8, t.size_cover_sub, 1.3)
            s.text(Rect(tx, y, box.w * 0.8, sh), subtitle, size=t.size_cover_sub,
                   color="ink_muted", line_spacing=1.3)
            y += sh + 0.16
        if meta:
            s.line(m, BASE_H - t.margin - 0.62, BASE_W - m,
                   BASE_H - t.margin - 0.62, "hairline", 0.75)
            s.text(Rect(m, BASE_H - t.margin - 0.48, box.w, 0.30), meta,
                   size=t.size_small, color="ink_muted")
        return s

    def section(self, number, title, subtitle=None, agenda=None):
        """章节过渡页。给听众一个"我们讲到哪了"的锚点。"""
        t = self.theme
        s = Slide(self, self.prs.slides.add_slide(self._blank), t.bg_alt)
        self.slides.append(s)
        m = t.margin + 0.24
        w = BASE_W - 2 * m
        cy = BASE_H * 0.42
        if number:
            nsize = t.size_section_title * 1.5
            nh = nsize * 1.3 / 72.0     # 框高必须容得下字号，否则会被裁
            s.text(Rect(m, cy - nh - 0.14, 2.4, nh), str(number),
                   size=nsize, bold=True, color="accent", font_en=t.font_num,
                   anchor="bottom")
        tsize = fit_size(str(title), w, 1.2, t.size_section_title, min_pt=20)
        th = text_height(str(title), w, tsize, 1.18)
        s.text(Rect(m, cy, w, th), title, size=tsize, bold=True, color="ink",
               line_spacing=1.18)
        y = cy + th + 0.18
        if subtitle:
            s.text(Rect(m, y, w * 0.78, 0.6), subtitle, size=t.size_cover_sub,
                   color="ink_muted", line_spacing=1.3, fit=True)
        if agenda:
            ay = BASE_H - t.margin - 0.42
            xs = Rect(m, ay, w, 0.32).grid(len(agenda), 1, 0.16)
            for i, (item, cell) in enumerate(zip(agenda, xs), 1):
                on = str(i) == str(number).lstrip("0") or item == title
                s.text(cell, "%02d %s" % (i, item), size=t.size_note,
                       bold=on, color="accent" if on else "ink_muted")
        return s

    def closing(self, title="讨论与决策事项", items=None, meta=None):
        """结尾页：不要写"谢谢"，写你要对方决策什么。"""
        t = self.theme
        s = Slide(self, self.prs.slides.add_slide(self._blank), t.bg_alt)
        self.slides.append(s)
        m = t.margin + 0.24
        w = BASE_W - 2 * m
        y = BASE_H * (0.26 if items else 0.42)
        s.rect(Rect(m, y, 0.09, 0.56), fill="accent", radius=0)
        s.text(Rect(m + 0.30, y, w - 0.30, 0.62), title,
               size=t.size_section_title, bold=True, color="ink")
        y += 0.86
        if items:
            s.bullets(items, Rect(m + 0.30, y, w - 0.30,
                                  BASE_H - y - t.margin - 0.5),
                      size=t.size_h, marker="num", gap=0.18)
        if meta:
            s.text(Rect(m, BASE_H - t.margin - 0.40, w, 0.30), meta,
                   size=t.size_small, color="ink_muted")
        return s

    # ---- 收尾 ----
    def add_page_numbers(self, skip_first=True, total=True):
        n = len(self.slides)
        for i, s in enumerate(self.slides, 1):
            if skip_first and i == 1:
                continue
            s.page_number(i, n if total else None)
        return self

    def save(self, path):
        self.prs.save(path)
        return path


# --------------------------------------------------------------------------
# 事后校验：把"看起来没问题"变成"检查过没问题"
# --------------------------------------------------------------------------

def check_overflow(path, verbose=True):
    """扫描 pptx，报告文本溢出、越界、字号过小、缺东亚字体等问题。

    返回问题列表。生成后必须跑一遍 —— 溢出是中文 PPT 最高频且最显眼的缺陷。
    """
    prs = Presentation(path)
    pw = prs.slide_width / 914400
    ph = prs.slide_height / 914400
    scale = pw / BASE_W
    issues = []

    def add(sn, kind, msg):
        issues.append({"slide": sn, "kind": kind, "msg": msg})

    for sn, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            name = sh.name
            if sh.left is None or sh.top is None:
                continue
            x, y = sh.left / 914400, sh.top / 914400
            w = (sh.width or 0) / 914400
            h = (sh.height or 0) / 914400
            if x < -0.02 or y < -0.02 or x + w > pw + 0.02 or y + h > ph + 0.02:
                add(sn, "out_of_bounds",
                    "%s 超出画布: (%.2f,%.2f) %.2fx%.2f" % (name, x, y, w, h))
            if not sh.has_text_frame:
                continue
            tf = sh.text_frame
            txt = tf.text
            if not txt.strip():
                continue
            sizes, missing_ea = [], False
            for p in tf.paragraphs:
                for r in p.runs:
                    if r.font.size:
                        sizes.append(r.font.size.pt)
                    if r.text.strip() and re.search(r"[\u4e00-\u9fff]", r.text):
                        if r.font._element.find(qn("a:ea")) is None:
                            missing_ea = True
            if not sizes:
                continue
            sz = max(sizes) / scale       # 折算回设计单位
            eff = min(sizes) / scale
            if eff < 8.5:
                add(sn, "tiny_text", "%s 字号折算后仅 %.1fpt: %r" %
                    (name, eff, txt[:24]))
            if missing_ea:
                add(sn, "missing_ea_font",
                    "%s 中文缺少东亚字体设置，将回落主题字体: %r" % (name, txt[:24]))
            inner_w = w / scale - (tf.margin_left + tf.margin_right) / 914400 / scale
            inner_h = h / scale - (tf.margin_top + tf.margin_bottom) / 914400 / scale
            need = text_height(txt, max(inner_w, 0.05), sz, 1.28)
            if need > inner_h * 1.14 + 0.02:
                add(sn, "text_overflow",
                    "%s 文本预计溢出 需%.2f\" 实%.2f\": %r" %
                    (name, need, inner_h, txt[:24]))

    if verbose:
        if not issues:
            print("检查通过：%d 页，未发现溢出/越界问题。" % len(prs.slides))
        else:
            print("发现 %d 个问题：" % len(issues))
            for it in issues:
                print("  第%d页 [%s] %s" % (it["slide"], it["kind"], it["msg"]))
    return issues


__all__ = ["Deck", "Slide", "Theme", "Rect", "THEMES", "DARK", "LIGHT", "SLATE",
           "CANVAS", "BASE_W", "BASE_H", "apply_font", "add_shadow",
           "clear_shadow", "text_width", "text_height", "wrapped_lines",
           "fit_size", "check_overflow"]
