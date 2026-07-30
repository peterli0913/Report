"""把 合并成一页，调整格式.pptx 的两页合并为一页，并按汇报逻辑分层。

逻辑层级（按沟通确认）：
    现状背景：新态势下，对生产管理要求、挑战：容错空间越来越小
      → 重点举措：推行工厂常态化合规审计
        → 达成：飞检能力提升（Ⅰ.日清日结 / Ⅱ.软硬件提升 / Ⅲ.SME）

所以版面是"顶部背景条 + 左侧重点 + 箭头 + 右侧分组框"，而不是四栏并列 ——
飞检能力提升是常态化合规审计的结果，不与它平级。

页面上的每一句都取自源文件，未新增措辞。源第 2 页的三大方向（官方飞检准备及应对 /
基层GMP规范执行管理 / 自动化和高端生产设备GMP管理）按要求不再上页，数据仍留在
DIRECTIONS 里，把 SHOW_DIRECTIONS 改成 True 即可恢复成左侧第二列。

模板安全区按版式背景图 image10.jpg 实测：页眉止于 y 0.749"，页脚色条自 y 5.409" 起。
画布 10 x 5.625，deckkit 设计单位为 13.333 x 7.5，故实际字号 = 设计字号 x 0.75。

运行：python3 build_gmp_2026_onepage.py [输出路径.pptx]
"""

from __future__ import annotations

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(REPO, ".cursor", "skills", "exec-deck-builder",
                               "scripts"))

from deckkit import (LIGHT, Deck, Rect, add_shadow, apply_font,  # noqa: E402
                     check_overflow, wrapped_lines)
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402
from pptx.util import Pt  # noqa: E402
from lxml import etree  # noqa: E402

TEMPLATE = os.path.join(REPO, "合并成一页，调整格式.pptx")
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    HERE, "2026年GMP管理重点工作推进.pptx")

SHOW_DIRECTIONS = False   # 见模块 docstring

# 源文件配色：标题 132556、正文 252D30。小节标签原为 5B9BD5，白底对比度仅 3.0:1，
# 压深到 1F4E86（8.4:1）后仍是同一支蓝，投屏才看得清。
NAVY = "132556"
BLUE = "1F4E86"
INK = "252D30"
CARD_LINE = "C7D3E6"
CARD_BG = "F6F9FC"
GROUP_BG = "EDF2FA"

LINE_SPACING = 1.22
LINE_SPACING_EST = LINE_SPACING * 1.22   # 估高留余量，末行才不会顶出框
BULLET_INDENT = 0.20

T = LIGHT.variant(
    name="asymchem-gmp2026",
    bg="FFFFFF", bg_alt="F7FAFD",
    surface=CARD_BG, surface_alt=GROUP_BG,
    hairline=CARD_LINE,
    ink=INK, ink_muted="42546B", ink_on_accent="FFFFFF",
    primary=BLUE, secondary="2E5E9E", accent=BLUE,
    good="1A7F50", warn="A34E0F", bad="C0392B", neutral="5E7288",
    accent_text=BLUE, primary_text=NAVY, secondary_text="1F4E86",
    margin=0.40, margin_top=1.10, margin_bottom=0.41,
    size_title=26, size_h=15, size_body=13, size_small=12, size_note=10,
    font_cn="微软雅黑", font_en="Arial",
)

# 这个文件里有 37 个母版、好几个都叫「2_自定义版式」，按名字取会拿到另一套页眉。
# 先记下源页第 2 页实际用的版式 part 名，再在 Deck 里按 part 名精确取回同一个版式。
from pptx import Presentation  # noqa: E402

_probe = Presentation(TEMPLATE)
LAYOUT_PART = _probe.slides[1].slide_layout.part.partname
del _probe

deck = Deck(theme=T, template=TEMPLATE)
_target = None
for _m in deck.prs.slide_masters:
    for _lo in _m.slide_layouts:
        if _lo.part.partname == LAYOUT_PART:
            _target = _lo
            break
    if _target is not None:
        break
if _target is None:
    raise SystemExit("找不到源页所用版式：%s" % LAYOUT_PART)
deck._blank = _target

s = deck.slide()
SCALE = deck.scale


# --------------------------------------------------------------------------
# 文本工具：项目符号、估高、自动适配字号
# --------------------------------------------------------------------------

def set_bullet(paragraph, char="●"):
    pPr = paragraph._p.get_or_add_pPr()
    for tag in ("a:buFont", "a:buChar", "a:buNone", "a:buAutoNum"):
        node = pPr.find(qn(tag))
        if node is not None:
            pPr.remove(node)
    indent = int(BULLET_INDENT * 914400 * SCALE)
    pPr.set("marL", str(indent))
    pPr.set("indent", str(-indent))
    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buFont.set("typeface", "Arial")
    buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", char)


def plain_text(item: dict) -> str:
    if item.get("runs"):
        return "".join(seg.get("text", "") for seg in item["runs"])
    return item.get("text", "")


def block_height(lines: list[dict], box_w: float, size: float) -> float:
    """估算多段文本总高（设计英寸），含项目符号缩进与段前距。"""
    total = 0.0
    for item in lines:
        w = box_w - (BULLET_INDENT if item.get("bullet") else 0.0)
        sz = size * item.get("scale", 1.0)
        total += wrapped_lines(plain_text(item), max(w, 0.05), sz) \
            * sz * LINE_SPACING_EST / 72.0
        total += item.get("space_before", 0) / 72.0
    return total


def fit_size(groups, hi: float, lo: float, step: float = 0.2) -> float:
    """groups: [(lines, box_w, box_h)]。返回所有块都装得下的最大字号。"""
    size = hi
    while size > lo:
        if all(block_height(ls, w, size) <= h for ls, w, h in groups):
            return round(size, 1)
        size -= step
    return lo


def rich_box(at: Rect, lines: list[dict], size: float, lead: float = 0.0,
             anchor="top"):
    box = s.raw.shapes.add_textbox(s._i(at.x), s._i(at.y),
                                   s._i(at.w), s._i(at.h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Pt(0)
    tf.margin_top = tf.margin_bottom = Pt(0)
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP,
                          "middle": MSO_ANCHOR.MIDDLE}[anchor]

    first = True
    for item in lines:
        is_first = first
        p = tf.paragraphs[0] if is_first else tf.add_paragraph()
        first = False
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = LINE_SPACING
        sb = item.get("space_before", 0) + (0 if is_first else lead)
        if sb:
            p.space_before = Pt(sb * SCALE)
        if item.get("bullet"):
            set_bullet(p)
        runs = item.get("runs") or [{"text": item.get("text", "")}]
        for seg in runs:
            run = p.add_run()
            run.text = seg.get("text", "")
            apply_font(run.font, T,
                       size=size * item.get("scale", 1.0) * SCALE,
                       bold=seg.get("bold", item.get("bold", False)),
                       color=seg.get("color", item.get("color", "ink")))
    return box


def extra_leading(lines, box_w, size, box_h, cap=5.0):
    """把剩余高度摊到段间，短块底部就不会留一大条空白。"""
    slack = box_h - block_height(lines, box_w, size)
    gaps = max(len(lines) - 1, 1)
    return max(0.0, min(cap, slack * 72.0 * 0.6 / gaps))


def head_band(at: Rect, text: str, size: float, fill=NAVY, h=0.50):
    """深色标题带 + 白字。返回标题带下沿 y。"""
    s.rect(Rect(at.x, at.y, at.w, h), fill=fill, radius=0.05)
    s.rect(Rect(at.x, at.y + h - 0.12, at.w, 0.12), fill=fill, radius=0)
    s.text(Rect(at.x + 0.22, at.y, at.w - 0.44, h), text, size=size,
           bold=True, color="ink_on_accent", font_cn="微软雅黑",
           anchor="middle", line_spacing=1.10)
    return at.y + h


# ==========================================================================
# 页标题（右上，沿用源页第 2 页写法）
# ==========================================================================
s.text(Rect(5.40, 0.24, 7.53, 0.44), "2026年GMP管理重点工作推进",
       size=26, bold=True, color="primary_text", align="right",
       font_cn="微软雅黑")
s.text(Rect(5.40, 0.71, 7.53, 0.28),
       "2026 GMP Governance Priorities & Implementation",
       size=14, bold=True, color="secondary_text", align="right")


# ==========================================================================
# 内容（文字均取自源文件）
# ==========================================================================
BACKGROUND = "新态势下，对生产管理要求、挑战：容错空间越来越小"

AUDIT = {
    "head": "推行工厂常态化合规审计",
    "lines": [
        {"text": "各个部门：", "bold": True, "color": "accent"},
        {"text": "（1）各个部门按照既定的制度进行监控检查（每日，每工作日，每周等）；",
         "bullet": True},
        {"text": "（2）非生产区域文件存放、5S管理、电脑使用巡查----每月执行一次，"
                 "生产部、IT、综办管理人员", "bullet": True},
        {"text": "（3）生产区域合规性+交叉污染联合巡查----生产部巡查、质量组、"
                 "车间质量管理人员+QA", "bullet": True},
        {"text": "（4）专项检查---生产部巡查组执行：", "bullet": True},
        {"text": "（5）高风险专项检查（交叉污染和D级区管控）--生产部质量组负责人和质量主任",
         "bullet": True},
        {"text": "2. QA部门：每周按照审批的巡查计划，执行主题巡查；"
                 "每个月QA管理人员进行飞检", "bold": True, "space_before": 7},
    ],
}

CAPABILITY_HEAD = "飞检能力提升"
CAPABILITY = [
    {
        "head": "Ⅰ.日清日结",
        "lines": [
            {"text": "现场工作的日清日结", "bold": True, "color": "accent"},
            {"text": "定主题", "bullet": True},
            {"text": "定人员", "bullet": True},
            {"text": "定频率", "bullet": True},
            {"text": "定跟踪", "bullet": True},
            {"text": "GMP批次单批记录关闭", "bold": True, "color": "accent",
             "space_before": 6},
            {"text": "每周开展记录书写的专项培训", "bullet": True},
            {"text": "记录问题统计个人/批次，奖励长期执行较好的人员，"
                     "对于问题数较多的人员，强化培训或调岗", "bullet": True},
            {"text": "Ⅲ.生产记录及逻辑性关闭-7天", "bold": True,
             "color": "accent", "space_before": 6},
            {"text": "生产开始，每天进行逻辑性数据统计与确认-生产、项目、分析、物料，"
                     "QA最终复核", "bullet": True},
        ],
    },
    {
        "head": "Ⅱ.软硬件提升",
        "lines": [
            {"text": "计算机化系统", "bullet": True},
            {"text": "微生物空调系统改造，满足阳性室和限度室独立空调分区管理-",
             "bullet": True},
            {"text": "DCS连续记录", "bullet": True},
        ],
    },
    {
        "head": "Ⅲ.SME",
        "lines": [
            {"text": "各部门审计相关的Topic", "bullet": True},
            {"text": "各部门的主题培训，结合主题巡查与日常执行中的典型问题",
             "bullet": True},
            {"text": "关键的厂区能力建设，如分析仪器管理、DCS、成品管理、环境控制",
             "bullet": True},
        ],
    },
]

# 源第 2 页三大方向：按要求本次不上页，保留数据便于随时恢复
DIRECTIONS = {
    "head": "GMP 管理能力建设三大方向",
    "lines": [
        {"text": "官方飞检准备及应对", "bold": True, "color": "accent"},
        {"text": "持续提升SME人员能力、重要审计前的工厂预审计专项合规性巡查、"
                 "合规性整改的全生命周期管控。"},
        {"text": "基层GMP规范执行管理", "bold": True, "color": "accent",
         "space_before": 6},
        {"text": "培训规范化：结合案例讲解记录的重要性，定期考核与反馈； "
                 "加强监督：设立专人抽查，建立数据，记录，质量违规异常预警机制；"
                 "考核和问责：将记录质量纳入KPI；"
                 "管理层推动：各层级管理人员带头重视数据，记录文化，"
                 "定期审查记录并推动改进。"},
        {"text": "自动化和高端生产设备GMP管理", "bold": True, "color": "accent",
         "space_before": 6},
        {"text": "对于不断更新迭代的设备：硬件（设施设备）上合规，"
                 "在软件（工艺、数据、文件）上可靠，"
                 "更在人员（团队能力与文化）上专业且严谨。"},
    ],
}


# ==========================================================================
# 版面：背景条 → 左侧重点 → 箭头 → 右侧分组框
# ==========================================================================
AREA = Rect(T.margin, 1.10, 13.3333 - 2 * T.margin, 7.10 - 1.10)

# ---- 现状背景（全宽一条，用浅底 + 左侧蓝竖条，克制但显眼）----
BG_H = 0.56
bg = Rect(AREA.x, AREA.y, AREA.w, BG_H)
s.rect(bg, fill=GROUP_BG, line=CARD_LINE, line_w=1.0, radius=0.06)
s.rect(Rect(bg.x, bg.y, 0.09, bg.h), fill=BLUE, radius=0)
s.text(Rect(bg.x + 0.30, bg.y, bg.w - 0.60, bg.h), BACKGROUND,
       size=17, bold=True, color="primary_text", font_cn="微软雅黑",
       anchor="middle")

MAIN = Rect(AREA.x, bg.bottom + 0.18, AREA.w, AREA.bottom - bg.bottom - 0.18)

ARROW_W = 0.62
HEAD_H = 0.50
PAD = 0.22
SUB_GAP = 0.16
SUB_HEAD_H = 0.38
SUB_PAD = 0.16
SUB_HEAD_SIZE = 14.0


def split_subs(inner: Rect):
    """三张子卡的宽度：按正文字数加权，但每张都不能窄于自己标题排一行的宽度。"""
    from deckkit import text_width
    total = inner.w - SUB_GAP * (len(CAPABILITY) - 1)
    weights = [sum(len(plain_text(it)) for it in c["lines"]) ** 0.6
               for c in CAPABILITY]
    base = [total * w / sum(weights) for w in weights]
    floor = [text_width(c["head"], SUB_HEAD_SIZE) + 2 * SUB_PAD + 0.08
             for c in CAPABILITY]
    widths = [max(b, f) for b, f in zip(base, floor)]
    over = sum(widths) - total
    if over > 0:
        # 超出的宽度只从"还高于自己下限"的卡上按余量比例回收
        slack = [w - f for w, f in zip(widths, floor)]
        pool = sum(slack)
        if pool > 0:
            widths = [w - over * sl / pool for w, sl in zip(widths, slack)]
    out, x = [], inner.x
    for w in widths:
        out.append(Rect(x, inner.y, w, inner.h))
        x += w + SUB_GAP
    return out


def geometry(left_ratio: float):
    """给定左右比例，算出各文本框几何与可用字号。只做计算，不画图形。"""
    left = Rect(MAIN.x, MAIN.y, MAIN.w * left_ratio, MAIN.h)
    right = Rect(left.right + ARROW_W, MAIN.y,
                 MAIN.right - left.right - ARROW_W, MAIN.h)
    abox = Rect(left.x + PAD, left.y + HEAD_H + 0.14,
                left.w - 2 * PAD, left.h - HEAD_H - 0.44)
    inner = Rect(right.x + 0.16, right.y + HEAD_H + 0.14,
                 right.w - 0.32, right.h - HEAD_H - 0.44)
    subs = split_subs(inner)
    sub_body_h = inner.h - SUB_HEAD_H - 0.22
    size = fit_size(
        [(AUDIT["lines"], abox.w, abox.h)]
        + [(c["lines"], sub.w - 2 * SUB_PAD, sub_body_h)
           for c, sub in zip(CAPABILITY, subs)],
        hi=14.0, lo=9.0)
    return size, left, right, abox, inner, subs, sub_body_h


# 左右比例交给搜索：内容多的一侧自动拿到更多宽度，字号才能取到最大档
_best = None
for _i in range(10, 25):
    _r = _i / 50.0                      # 0.20 ~ 0.48
    _g = geometry(_r)
    if _best is None or _g[0] > _best[0][0]:
        _best = (_g, _r)
(BODY_SIZE, LEFT, RIGHT, audit_box, INNER, subs, sub_body_h), LEFT_RATIO = _best

# ---- 左侧：重点举措 ----
card = s.rect(LEFT, fill=CARD_BG, line=CARD_LINE, line_w=1.0, radius=0.04)
add_shadow(card, blur=8, dist=3, alpha=18)
head_band(LEFT, AUDIT["head"], 16.0, h=HEAD_H)

# ---- 右侧：飞检能力提升分组框 + 三张子卡 ----
gframe = s.rect(RIGHT, fill=GROUP_BG, line=CARD_LINE, line_w=1.25, radius=0.04)
add_shadow(gframe, blur=8, dist=3, alpha=18)
head_band(RIGHT, CAPABILITY_HEAD, 16.0, h=HEAD_H)

rich_box(audit_box, AUDIT["lines"], size=BODY_SIZE,
         lead=extra_leading(AUDIT["lines"], audit_box.w, BODY_SIZE,
                            audit_box.h))

for sub, data in zip(subs, CAPABILITY):
    sc = s.rect(sub, fill="FFFFFF", line=CARD_LINE, line_w=1.0, radius=0.04)
    add_shadow(sc, blur=5, dist=2, alpha=12)
    s.rect(Rect(sub.x, sub.y, sub.w, 0.055), fill=BLUE, radius=0)
    s.text(Rect(sub.x + SUB_PAD, sub.y + 0.10, sub.w - 2 * SUB_PAD, 0.28),
           data["head"], size=SUB_HEAD_SIZE, bold=True, color="primary_text",
           font_cn="微软雅黑")
    s.line(sub.x + SUB_PAD, sub.y + SUB_HEAD_H, sub.right - SUB_PAD,
           sub.y + SUB_HEAD_H, "hairline", 1.0)
    body = Rect(sub.x + SUB_PAD, sub.y + SUB_HEAD_H + 0.10,
                sub.w - 2 * SUB_PAD, sub_body_h)
    # 内容少的卡垂直居中，留白上下均分，三张卡看起来才是一排而不是吊在顶上
    rich_box(body, data["lines"], size=BODY_SIZE, anchor="middle",
             lead=extra_leading(data["lines"], body.w, BODY_SIZE, sub_body_h))

# ---- 中间箭头：常态化合规审计 → 达成飞检能力提升 ----
ar_h = 0.68
ar = Rect(LEFT.right + 0.10, MAIN.y + (MAIN.h - ar_h) / 2,
          ARROW_W - 0.20, ar_h)
s.rect(ar, fill=BLUE, shape=MSO_SHAPE.RIGHT_ARROW, radius=0)

if SHOW_DIRECTIONS:
    raise SystemExit("恢复三大方向需要重新分配左侧两列，请按需调整 LEFT 区域")


deck.save(OUT)
print("已生成：%s（%d 页，%.2f x %.2f 英寸）"
      % (OUT, len(deck.slides), deck.canvas_w, deck.canvas_h))
print("正文字号 %.1f（实际 %.1fpt）；左右比例 %.2f"
      % (BODY_SIZE, BODY_SIZE * SCALE, LEFT_RATIO))
print("重点区 占高 %.2f\" / 可用 %.2f\""
      % (block_height(AUDIT["lines"], audit_box.w, BODY_SIZE), audit_box.h))
for c, sub in zip(CAPABILITY, subs):
    print("  %-12s 宽 %.2f\" 占高 %.2f\" / 可用 %.2f\""
          % (c["head"], sub.w,
             block_height(c["lines"], sub.w - 2 * SUB_PAD, BODY_SIZE),
             sub_body_h))
check_overflow(OUT)
