"""把 合并成一页，调整格式.pptx 的两页合并为一页。

- 文字内容原样保留（含源文件自身的编号写法，如两处「Ⅲ.」）
- 四栏单排铺满全高：重点方向 / 常态化合规审计 / 日清日结 / 软硬件提升+SME
- 栏宽按各栏字数加权分配，正文字号自动适配到"能装下的最大值"并四栏统一

模板安全区按版式背景图 image10.jpg 实测：页眉止于 y 0.765"，页脚色条自 y 5.391" 起。
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
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402
from pptx.util import Pt  # noqa: E402
from lxml import etree  # noqa: E402

TEMPLATE = os.path.join(REPO, "合并成一页，调整格式.pptx")
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    HERE, "2026年GMP管理重点工作推进.pptx")

# 源文件配色：标题 132556、正文 252D30。小节标签原为 5B9BD5，白底对比度仅 3.0:1，
# 压深到 1F4E86（8.4:1）后仍是同一支蓝，投屏才看得清。
NAVY = "132556"
BLUE = "1F4E86"
INK = "252D30"
CARD_LINE = "C7D3E6"
CARD_BG = "F6F9FC"

LINE_SPACING = 1.22
LINE_SPACING_EST = LINE_SPACING * 1.22   # 估高留余量，末行才不会顶出框
BULLET_INDENT = 0.20

T = LIGHT.variant(
    name="asymchem-gmp2026",
    bg="FFFFFF", bg_alt="F7FAFD",
    surface=CARD_BG, surface_alt="EAF0F9",
    hairline=CARD_LINE,
    ink=INK, ink_muted="42546B", ink_on_accent="FFFFFF",
    primary=BLUE, secondary="2E5E9E", accent=BLUE,
    good="1A7F50", warn="A34E0F", bad="C0392B", neutral="5E7288",
    accent_text=BLUE, primary_text=NAVY, secondary_text="1F4E86",
    margin=0.40, margin_top=1.10, margin_bottom=0.41,
    size_title=26, size_h=15, size_body=12, size_small=11, size_note=10,
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
    """估算多段文本总高（设计英寸），含项目符号缩进、段前距、小标题字号差。"""
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


def rich_box(at: Rect, lines: list[dict], size: float, lead: float = 0.0):
    box = s.raw.shapes.add_textbox(s._i(at.x), s._i(at.y),
                                   s._i(at.w), s._i(at.h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Pt(0)
    tf.margin_top = tf.margin_bottom = Pt(0)
    tf.vertical_anchor = MSO_ANCHOR.TOP

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


def extra_leading(lines, box_w, size, box_h, cap=3.0):
    """把剩余高度摊到段间，短栏底部就不会留一大条空白。"""
    slack = box_h - block_height(lines, box_w, size)
    gaps = max(len(lines) - 1, 1)
    return max(0.0, min(cap, slack * 72.0 * 0.6 / gaps))


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
# 内容（文字与源文件一致）
# ==========================================================================
COL_DIRECTION = {
    # 栏头与其余三栏取齐：都是名词性主题短语。源文件那句挑战陈述移到正文首行保留。
    "head": "GMP 管理能力建设三大方向",
    "lines": [
        {"text": "新态势下，对生产管理要求、挑战：容错空间越来越小",
         "bold": True, "color": "ink_muted"},
        {"text": "官方飞检准备及应对", "bold": True, "color": "accent",
         "space_before": 5},
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

COL_AUDIT = {
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
                 "每个月QA管理人员进行飞检", "bold": True, "space_before": 6},
    ],
}

COL_DAILY = {
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
        {"text": "Ⅲ.生产记录及逻辑性关闭-7天", "bold": True, "color": "accent",
         "space_before": 6},
        {"text": "生产开始，每天进行逻辑性数据统计与确认-生产、项目、分析、物料，"
                 "QA最终复核", "bullet": True},
    ],
}

COL_TECH = {
    # 栏头直接用源文件的「Ⅱ.软硬件提升」，正文里就不再重复这一行标签
    "head": "Ⅱ.软硬件提升",
    "lines": [
        {"text": "计算机化系统", "bullet": True},
        {"text": "微生物空调系统改造，满足阳性室和限度室独立空调分区管理-",
         "bullet": True},
        {"text": "DCS连续记录", "bullet": True},
        {"text": "Ⅲ.SME", "bold": True, "color": "accent", "space_before": 6},
        {"text": "各部门审计相关的Topic", "bullet": True},
        {"text": "各部门的主题培训，结合主题巡查与日常执行中的典型问题",
         "bullet": True},
        {"text": "关键的厂区能力建设，如分析仪器管理、DCS、成品管理、环境控制",
         "bullet": True},
    ],
}

COLUMNS = [COL_DIRECTION, COL_AUDIT, COL_DAILY, COL_TECH]


# ==========================================================================
# 版面：一行小节标签 + 四栏单排铺满全高
# ==========================================================================
AREA = Rect(T.margin, 1.10, 13.3333 - 2 * T.margin, 7.09 - 1.10)

# 小节标签（源页第 1 页的标题与副标题）
s.rect(Rect(AREA.x, AREA.y + 0.04, 0.07, 0.24), fill=BLUE, radius=0)
s.text(Rect(AREA.x + 0.18, AREA.y, 6.00, 0.30),
       "重点质量管理提升 · 飞检能力提升", size=15, bold=True,
       color="primary_text", font_cn="微软雅黑")

COLS_TOP = AREA.y + 0.40
GAP = 0.22
PAD_X = 0.20
HEAD_H = 0.66          # 栏头统一高度：最长的那句能排两行，四栏正文才同起点
BODY_PAD_BOTTOM = 0.16

weights = [sum(len(plain_text(it)) for it in c["lines"]) ** 0.6
           for c in COLUMNS]
total_w = AREA.w - GAP * (len(COLUMNS) - 1)
widths = [total_w * w / sum(weights) for w in weights]

cols, x = [], AREA.x
for w in widths:
    cols.append(Rect(x, COLS_TOP, w, AREA.bottom - COLS_TOP))
    x += w + GAP

body_h = cols[0].h - HEAD_H - BODY_PAD_BOTTOM
BODY_SIZE = fit_size([(c["lines"], col.w - 2 * PAD_X, body_h)
                      for c, col in zip(COLUMNS, cols)], hi=13.5, lo=9.0)

# 栏头统一字号：取最长那句也能排进两行的档位，四栏色带看起来才是一套
HEAD_SIZE = 15.0
while HEAD_SIZE > 11.0:
    if all(wrapped_lines(c["head"], col.w - 2 * PAD_X, HEAD_SIZE) <= 2
           for c, col in zip(COLUMNS, cols)):
        break
    HEAD_SIZE -= 0.25

for col, data in zip(cols, COLUMNS):
    card = s.rect(col, fill=CARD_BG, line=CARD_LINE, line_w=1.0, radius=0.035)
    add_shadow(card, blur=7, dist=3, alpha=16)
    # 栏头深蓝色带 + 白字，四栏节奏一致
    s.rect(Rect(col.x, col.y, col.w, HEAD_H - 0.12), fill=NAVY, radius=0.035)
    s.rect(Rect(col.x, col.y + HEAD_H - 0.22, col.w, 0.10), fill=NAVY,
           radius=0)
    s.text(Rect(col.x + PAD_X, col.y, col.w - 2 * PAD_X, HEAD_H - 0.12),
           data["head"], size=HEAD_SIZE, bold=True,
           color="ink_on_accent", font_cn="微软雅黑", anchor="middle",
           line_spacing=1.12)

    lead = extra_leading(data["lines"], col.w - 2 * PAD_X, BODY_SIZE, body_h)
    rich_box(Rect(col.x + PAD_X, col.y + HEAD_H, col.w - 2 * PAD_X, body_h),
             data["lines"], size=BODY_SIZE, lead=lead)


deck.save(OUT)
print("已生成：%s（%d 页，%.2f x %.2f 英寸）"
      % (OUT, len(deck.slides), deck.canvas_w, deck.canvas_h))
print("正文字号 %.1f（实际 %.1fpt）；栏宽 %s"
      % (BODY_SIZE, BODY_SIZE * SCALE, ["%.2f" % w for w in widths]))
for c, col in zip(COLUMNS, cols):
    print("  %-14s 占高 %.2f\" / 可用 %.2f\""
          % (c["head"][:14],
             block_height(c["lines"], col.w - 2 * PAD_X, BODY_SIZE), body_h))
check_overflow(OUT)
