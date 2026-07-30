"""将 合成一页.pptx 的两页合并为一页（四栏：生产 / QC / 计算机化系统 / 库房）。

- 不增删改文字内容，仅重排版式
- 栏宽按各栏内容量加权分配，正文字号自动适配到"能装下的最大值"并四栏统一
- 沿用源文件凯莱英模板与视觉语言（大号序号、竖分隔线、圆角内容框、高风险红字）

运行：python3 build_gmp_onepage.py [输出路径.pptx]
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

TEMPLATE = os.path.join(REPO, "合成一页.pptx")
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    HERE, "工厂GMP合规管理.pptx")

BLUE = "1D42E1"
TITLE_BLUE = "132556"
INK = "1A1A1A"
# 源页高风险用纯红 FF0000，白底上对比度只有 4.0:1，投屏偏亮。压深一档到 D40000
# （5.5:1）既过 WCAG AA，视觉上仍是同一套警示语言。
RED = "D40000"
CARD_BG = "FFFFFF"

# 正文自动适配的上下限（设计单位；实际画布为 2 倍，故 9.0 → 18pt）
BODY_MAX = 10.0
BODY_MIN = 8.6
LINE_SPACING = 1.20
# 估高时按更大的行距算：PowerPoint / LibreOffice 的实际行高含字体升降部，
# 只按 line_spacing 折算会偏乐观，末尾几行就会顶进页脚色条。
LINE_SPACING_EST = LINE_SPACING * 1.22
BULLET_INDENT = 0.16          # 项目符号悬挂缩进，量高时要从可用宽度里扣掉

T = LIGHT.variant(
    name="asymchem-gmp",
    bg="FFFFFF", bg_alt="F7FAFD",
    surface=CARD_BG, surface_alt="F3F6FC",
    hairline="C5D0E6",
    ink=INK, ink_muted="4E6584", ink_on_accent="FFFFFF",
    primary=BLUE, secondary="3263A7", accent=BLUE,
    good="1A7F50", warn="A34E0F", bad=RED, neutral="5E7288",
    accent_text=BLUE, bad_text=RED, primary_text=TITLE_BLUE,
    margin=0.55, margin_top=1.00, margin_bottom=0.50,
    size_title=25, size_h=13, size_body=9, size_small=8.5, size_note=7,
    font_cn="微软雅黑", font_en="Arial",
)

deck = Deck(theme=T, template=TEMPLATE, canvas="large", layout="2_自定义版式")
s = deck.slide()


# --------------------------------------------------------------------------
# 文本工具
# --------------------------------------------------------------------------

def set_bullet(paragraph, char="●"):
    pPr = paragraph._p.get_or_add_pPr()
    for tag in ("a:buFont", "a:buChar", "a:buNone", "a:buAutoNum"):
        node = pPr.find(qn(tag))
        if node is not None:
            pPr.remove(node)
    indent = int(BULLET_INDENT * 914400 * deck.scale)
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


def block_height(lines: list[dict], box_w: float, size: float,
                 line_spacing: float = LINE_SPACING_EST) -> float:
    """估算整块多段文本的高度（设计英寸），含项目符号缩进与段前距。"""
    total = 0.0
    for item in lines:
        w = box_w - (BULLET_INDENT if item.get("bullet") else 0.0)
        n = wrapped_lines(plain_text(item), max(w, 0.05), size)
        total += n * size * line_spacing / 72.0
        total += item.get("space_before", 0) / 72.0
    return total


def fit_block_size(lines: list[dict], box_w: float, box_h: float,
                   hi: float = BODY_MAX, lo: float = BODY_MIN,
                   step: float = 0.1) -> float:
    """求能装进 box_h 的最大字号。装不下也不低于 lo（再小就不该做成一页）。"""
    size = hi
    while size > lo:
        if block_height(lines, box_w, size) <= box_h:
            return round(size, 1)
        size -= step
    return lo


def extra_leading(lines: list[dict], box_w: float, size: float,
                  box_h: float, cap: float = 3.2) -> float:
    """把剩余高度摊到段间，短栏就不会在底部留一大条空白。返回每段追加的 pt。"""
    slack = box_h - block_height(lines, box_w, size)
    gaps = max(len(lines) - 1, 1)
    # slack 按估高口径算，回填时打七折，避免估算误差把文字又顶出去
    return max(0.0, min(cap, slack * 72.0 * 0.7 / gaps))


def rich_box(at: Rect, lines: list[dict], size: float,
             line_spacing: float = LINE_SPACING, lead: float = 0.0):
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
        p.line_spacing = line_spacing
        sb = item.get("space_before", 0) + (0 if is_first else lead)
        if sb:
            p.space_before = Pt(sb * deck.scale)
        if item.get("bullet"):
            set_bullet(p)
        runs = item.get("runs") or [{"text": item.get("text", ""),
                                     "bold": item.get("bold", False),
                                     "color": item.get("color", "ink")}]
        for seg in runs:
            run = p.add_run()
            run.text = seg.get("text", "")
            apply_font(run.font, T, size=size * deck.scale,
                       bold=seg.get("bold", False),
                       color=seg.get("color", "ink"))
    return box


# ==========================================================================
# 页标题（右上，与源页一致）
# ==========================================================================
s.text(Rect(6.00, 0.24, 6.52, 0.54), "工厂GMP合规管理",
       size=25, bold=True, color="primary_text", align="right",
       font_cn="微软雅黑")


# ==========================================================================
# 内容
# ==========================================================================
COLUMNS = [
    {
        "num": "01",
        "title": "生产车间常见合规性问题",
        "lines": [
            {"text": "一、人员管理类高频合规性问题：着装，培训，外来人员管理，复核，隐瞒不报"},
            {"text": "二、物料、中间产品管理类（混淆、交叉污染）"},
            {"text": "三、工艺操作合规问题"},
            {"text": "四、设备、设施相关问题（衔接设备安全巡检）：状态，标识，清洁，跑冒滴漏，交叉污染；"},
            {"text": "五、洁净区环境与公用系统（无菌车间重灾区）"},
            {"text": "六、记录与文件缺陷（飞检第一关注重点）"},
            {"text": "七、计算机化系统 & 数据完整性问题"},
            {"text": "八、质量体系工具落地问题（偏差、变更、CAPA、OOS瞒报，调查不充分）"},
            {"text": "九、现场 EHS 合规问题（GMP + 安全生产双重要求）"},
            {"text": "重大高风险缺陷（极易导致官方警告、停产）",
             "bold": True, "color": "bad", "space_before": 7},
            {"text": "伪造、篡改生产 / 检验原始数据；共享账号规避系统管控", "bullet": True},
            {"text": "未经 QA 放行物料投入生产", "bullet": True},
            {"text": "隐瞒重大异常，不报偏差，继续生产产品", "bullet": True},
            {"text": "拆除、屏蔽设备安全、工艺联锁装置", "bullet": True},
            {"text": "跨品种生产未有效清场，存在严重交叉污染风险", "bullet": True},
            {"text": "主要一般缺陷（日常巡检最常见）",
             "bold": True, "space_before": 7},
            {"text": "记录填写不及时、信息不全；记录修改不规范", "bullet": True},
            {"text": "物料标识不清、容器敞口存放", "bullet": True},
            {"text": "洁净区人员行为违规、更衣不规范", "bullet": True},
            {"text": "环境监测超标未调查；消毒管理不规范", "bullet": True},
            {"text": "维保、巡检记录漏记；计量器具临近 / 超校准周期", "bullet": True},
            {"text": "审计追踪仅开启，未实施周期性审核", "bullet": True},
        ],
    },
    {
        "num": "02",
        "title": "QC 实验室常见 GMP 违规行为",
        "lines": [
            {"text": "一、样品管理类（高频问题）；"},
            {"text": "二、检验操作行为违规"},
            {"text": "三、原始记录与数据完整性（GMP 红线，重大违规）"},
            {"text": "四、仪器设备管理违规；"},
            {"text": "五、试剂、试液、标准品、培养基管理；"},
            {"text": "六、实验室环境、清洁消毒、微生物专项违规"},
            {"text": "七、文件、报告、偏差 OOS 管理，调查"},
            {"text": "八、人员卫生与行为规范"},
            {"runs": [
                {"text": "重大高风险缺陷：", "bold": True, "color": "bad"},
                {"text": "伪造数据、删除图谱、共用账号、隐瞒 OOS 结果、擅自更改检验方法、挪用留样"},
             ], "space_before": 9},
            {"runs": [
                {"text": "一般违规：", "bold": True},
                {"text": "样品存放条件不达标、漏做仪器日常核查、试剂标识不全、未及时填写设备日志；"},
             ], "space_before": 7},
            {"runs": [
                {"text": "轻微违规：", "bold": True},
                {"text": "记录书写不规范、物品摆放杂乱、清洁不及时、标识粘贴不整齐；"},
             ], "space_before": 7},
        ],
    },
    {
        "num": "03",
        "title": "计算机化系统",
        "lines": [
            {"text": "一、用户访问与账号权限持续管理"},
            {"text": "二、审计追踪（Audit Trail）常态化管理（核查高频重点）"},
            {"text": "三、系统日常运维管理（所有运维操作不能脱离 GMP 体系）"},
            {"text": "四、备份管理与定期灾难恢复演练"},
            {"text": "五、电子记录、电子签名日常管控"},
            {"text": "六、系统异常事件、故障管理（IT 与质量体系打通）"},
            {"text": "七、变更控制—— 运行阶段所有系统变更必经流程"},
            {"text": "八、供应商持续管理（外购软件、SaaS 云系统重点）"},
            {"text": "十、人员持续培训与行为监督"},
            {"text": "十一、周期性系统合规回顾（System Periodic Review，SPR）"},
            {"text": "欧美药企标准要求：每 12 个月开展计算机化系统周期性回顾 SPR"},
            {"text": "十二、文档持续维护"},
        ],
    },
    {
        "num": "04",
        "title": "库房管理常见合规性问题",
        "lines": [
            {"text": "一、库房区域与硬件设施合规问题（高频）"},
            {"text": "二、物料存放与状态管理问题（GMP重点缺陷）"},
            {"text": "三、入库、取样、出库、退库流程合规问题"},
            {"text": "四、温湿度、冷链、储存条件合规问题"},
            {"text": "五、台账、记录、追溯管理问题（核查重灾区）"},
            {"text": "六、计算机化系统与数据完整性问题（WMS/ERP）"},
            {"text": "七、清洁、卫生、防虫鼠、环境管理问题"},
            {"text": "八、危化品、有机溶剂库房专项问题（EHS+GMP双重缺陷）"},
            {"text": "九、人员管理与操作合规问题"},
            {"text": "十、质量体系执行问题（偏差/变更/CAPA）"},
            {"text": "重大高风险缺陷（飞检直接判重缺陷）",
             "bold": True, "color": "bad", "space_before": 8},
            {"text": "未经放行物料入库、领用、投产", "bullet": True},
            {"text": "账物严重不符、虚假记录、补录数据", "bullet": True},
            {"text": "不合格品与合格品混放、失控管理", "bullet": True},
            {"text": "危化品违规混存、无防爆防静电措施", "bullet": True},
            {"text": "系统共享账号、删除/篡改物料数据", "bullet": True},
            {"text": "温湿度长期超标不处置、瞒报异常", "bullet": True},
        ],
    },
]


# ==========================================================================
# 版面：白底圆角大框 + 四栏（栏宽按内容量加权，单排铺满全高）
# ==========================================================================
# 底沿留出余量：模板页脚深蓝条从 y 7.18 起，内容压上去就会被色条吃掉最后一行
FRAME = Rect(0.52, 0.90, 12.30, 6.14)
frame_sh = s.rect(FRAME, fill=CARD_BG, line=BLUE, line_w=1.5, radius=0.03)
add_shadow(frame_sh, blur=10, dist=5, alpha=22)
# 顶部一条品牌色细条，替代原页的渐变装饰
s.rect(Rect(FRAME.x + 0.02, FRAME.y + 0.02, FRAME.w - 0.04, 0.05),
       fill=BLUE, radius=0)

INNER = FRAME.inset(0.24, 0.22)
GAP = 0.24
HEAD_H = 0.60                    # 序号 + 栏目标题 + 分隔线占用高度
PAD_X = 0.06                     # 正文相对栏左沿的内缩
FOOT_KEEP = 0.10                 # 正文底部留白，别顶到框线

# 用字符数的 0.6 次幂做阻尼加权：内容多的栏更宽，但不至于把窄栏挤瘪
weights = [sum(len(plain_text(it)) for it in c["lines"]) ** 0.6
           for c in COLUMNS]
total_w = INNER.w - GAP * (len(COLUMNS) - 1)
widths = [total_w * w / sum(weights) for w in weights]

cols = []
x = INNER.x
for w in widths:
    cols.append(Rect(x, INNER.y, w, INNER.h))
    x += w + GAP

# 正文字号：四栏取统一值（同一页内字号不一致会显得拼凑），取各栏可用值的最小者
body_h = INNER.h - HEAD_H - FOOT_KEEP
sizes = [fit_block_size(c["lines"], col.w - PAD_X, body_h)
         for c, col in zip(COLUMNS, cols)]
BODY_SIZE = min(sizes)

# 栏目标题：统一字号，取"最长标题也能排成一行"的档位，四栏基线才齐
NUM_W = 0.86
title_size = 13.0
while title_size > 9.5:
    if all(wrapped_lines(c["title"], col.w - NUM_W - 0.10, title_size) == 1
           for c, col in zip(COLUMNS, cols)):
        break
    title_size -= 0.25

for i, (col, data) in enumerate(zip(cols, COLUMNS)):
    if i > 0:
        s.line(col.x - GAP / 2, col.y + 0.04, col.x - GAP / 2,
               col.bottom - 0.04, "hairline", 1.0)

    # 序号 + 栏目标题
    s.text(Rect(col.x, col.y - 0.02, NUM_W, 0.46), data["num"],
           size=26, bold=True, color="primary", font_en="Arial",
           anchor="middle")
    s.text(Rect(col.x + NUM_W + 0.10, col.y - 0.02,
                col.w - NUM_W - 0.10, 0.46),
           data["title"], size=title_size, bold=True, color="primary_text",
           font_cn="微软雅黑", anchor="middle", line_spacing=1.08)
    s.line(col.x, col.y + HEAD_H - 0.12, col.right, col.y + HEAD_H - 0.12,
           "hairline", 1.25)

    lead = extra_leading(data["lines"], col.w - PAD_X, BODY_SIZE, body_h)
    rich_box(Rect(col.x + PAD_X, col.y + HEAD_H, col.w - PAD_X, body_h),
             data["lines"], size=BODY_SIZE, lead=lead)

deck.save(OUT)
print("已生成：%s（%d 页，%.2f x %.2f 英寸）"
      % (OUT, len(deck.slides), deck.canvas_w, deck.canvas_h))
print("栏宽（设计英寸）：", ["%.2f" % w for w in widths])
print("各栏可用字号：", sizes, "→ 统一取", BODY_SIZE,
      "（实际 %.1fpt）" % (BODY_SIZE * deck.scale))
print("栏目标题字号：%.2f（实际 %.1fpt）" % (title_size, title_size * deck.scale))
for c, col in zip(COLUMNS, cols):
    used = block_height(c["lines"], col.w - PAD_X, BODY_SIZE)
    print("  %s 正文占高 %.2f\" / 可用 %.2f\"" % (c["num"], used, body_h))
check_overflow(OUT)
