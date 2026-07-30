"""将 合成一页.pptx 的两页合并为一页（四栏：生产 / QC / 计算机化系统 / 库房）。

- 不增删改文字内容，仅重排版式
- 沿用源文件凯莱英模板与视觉语言（大号序号、竖线、圆角内容框、高风险红字）

运行：python3 build_gmp_onepage.py [输出路径.pptx]
"""

from __future__ import annotations

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(REPO, ".cursor", "skills", "exec-deck-builder",
                               "scripts"))

from deckkit import LIGHT, Deck, Rect, add_shadow, apply_font, check_overflow  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402
from pptx.util import Pt  # noqa: E402
from lxml import etree  # noqa: E402

TEMPLATE = os.path.join(REPO, "合成一页.pptx")
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    HERE, "工厂GMP合规管理.pptx")

# 源文件配色
BLUE = "1D42E1"
TITLE_BLUE = "132556"
INK = "1A1A1A"
RED = "FF0000"
FRAME_LINE = "1D42E1"

T = LIGHT.variant(
    name="asymchem-gmp",
    bg="FFFFFF", bg_alt="F7FAFD",
    surface="FFFFFF", surface_alt="F3F6FC",
    hairline="C5D0E6",
    ink=INK, ink_muted="4E6584", ink_on_accent="FFFFFF",
    primary=BLUE, secondary="3263A7", accent=BLUE,
    good="1A7F50", warn="A34E0F", bad=RED, neutral="5E7288",
    accent_text=BLUE, bad_text=RED, primary_text=TITLE_BLUE,
    margin=0.60, margin_top=1.05, margin_bottom=0.55,
    size_title=22, size_h=9, size_body=6.2, size_small=5.8, size_note=5,
    font_cn="微软雅黑", font_en="Arial",
)

deck = Deck(theme=T, template=TEMPLATE, canvas="large", layout="2_自定义版式")
s = deck.slide()  # 不自动写标题，按源页右上角放置


def set_bullet(paragraph, char="●"):
    """给段落加实心圆项目符号（对应源文件 Wingdings Ø 的视觉效果）。"""
    pPr = paragraph._p.get_or_add_pPr()
    for tag in ("a:buFont", "a:buChar", "a:buNone", "a:buAutoNum"):
        node = pPr.find(qn(tag))
        if node is not None:
            pPr.remove(node)
    pPr.set("marL", "114300")       # ~0.125"
    pPr.set("indent", "-114300")
    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buFont.set("typeface", "Arial")
    buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", char)


def rich_box(at: Rect, lines: list[dict], size: float = 6.2,
             line_spacing: float = 1.18):
    """写入多段富文本。

    lines 项支持:
      text / runs([{text,bold?,color?}...]) / bold / color / bullet / space_before
    """
    box = s.raw.shapes.add_textbox(s._i(at.x), s._i(at.y),
                                   s._i(at.w), s._i(at.h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Pt(0)
    tf.margin_top = tf.margin_bottom = Pt(0)
    tf.vertical_anchor = MSO_ANCHOR.TOP

    first = True
    for item in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = item.get("line_spacing", line_spacing)
        if item.get("space_before"):
            p.space_before = Pt(item["space_before"] * deck.scale)
        if item.get("space_after"):
            p.space_after = Pt(item["space_after"] * deck.scale)
        if item.get("bullet"):
            set_bullet(p, item.get("bullet_char", "●"))
        runs = item.get("runs")
        if not runs:
            runs = [{"text": item.get("text", ""),
                     "bold": item.get("bold", False),
                     "color": item.get("color", "ink")}]
        for seg in runs:
            run = p.add_run()
            run.text = seg.get("text", "")
            apply_font(run.font, T,
                       size=seg.get("size", item.get("size", size)) * deck.scale,
                       bold=seg.get("bold", False),
                       color=seg.get("color", "ink"),
                       font_en=seg.get("font_en"),
                       font_cn=seg.get("font_cn"))
    return box


def draw_frame(at: Rect):
    """源页风格：白底圆角框 + 蓝色描边 + 轻阴影。"""
    sh = s.rect(at, fill="FFFFFF", line=FRAME_LINE, line_w=1.5, radius=0.035)
    add_shadow(sh, blur=8, dist=4, alpha=25)
    return sh


# ==========================================================================
# 页标题（与源页一致：右上）
# ==========================================================================
s.text(Rect(6.40, 0.12, 6.30, 0.85), "工厂GMP合规管理",
       size=24, bold=True, color="primary_text", align="right",
       font_cn="微软雅黑")


# ==========================================================================
# 内容框与四栏
# ==========================================================================
FRAME = Rect(0.69, 1.14, 11.96, 5.96)
draw_frame(FRAME)

INNER = FRAME.inset(0.20, 0.18)
cols = INNER.split_h(1, 1, 1, 1, gap=0.18)

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
             "bold": True, "color": "bad", "space_before": 4},
            {"text": "伪造、篡改生产 / 检验原始数据；共享账号规避系统管控", "bullet": True},
            {"text": "未经 QA 放行物料投入生产", "bullet": True},
            {"text": "隐瞒重大异常，不报偏差，继续生产产品", "bullet": True},
            {"text": "拆除、屏蔽设备安全、工艺联锁装置", "bullet": True},
            {"text": "跨品种生产未有效清场，存在严重交叉污染风险", "bullet": True},
            {"text": "主要一般缺陷（日常巡检最常见）",
             "bold": True, "space_before": 3},
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
             ], "space_before": 5},
            {"runs": [
                {"text": "一般违规：", "bold": True},
                {"text": "样品存放条件不达标、漏做仪器日常核查、试剂标识不全、未及时填写设备日志；"},
             ], "space_before": 3},
            {"runs": [
                {"text": "轻微违规：", "bold": True},
                {"text": "记录书写不规范、物品摆放杂乱、清洁不及时、标识粘贴不整齐；"},
             ], "space_before": 3},
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
             "bold": True, "color": "bad", "space_before": 5},
            {"text": "未经放行物料入库、领用、投产", "bullet": True},
            {"text": "账物严重不符、虚假记录、补录数据", "bullet": True},
            {"text": "不合格品与合格品混放、失控管理", "bullet": True},
            {"text": "危化品违规混存、无防爆防静电措施", "bullet": True},
            {"text": "系统共享账号、删除/篡改物料数据", "bullet": True},
            {"text": "温湿度长期超标不处置、瞒报异常", "bullet": True},
        ],
    },
]

for i, (col, data) in enumerate(zip(cols, COLUMNS)):
    # 栏间细竖线（首栏不加）
    if i > 0:
        s.line(col.x - 0.08, col.y + 0.10, col.x - 0.08, col.bottom - 0.10,
               "hairline", 1.0)

    # 大号序号 + 栏目标题（沿用源页左号右题节奏）
    s.text(Rect(col.x, col.y - 0.04, 0.78, 0.58), data["num"],
           size=24, bold=True, color="primary", font_en="Arial",
           anchor="middle")
    s.text(Rect(col.x + 0.78, col.y + 0.06, col.w - 0.82, 0.46),
           data["title"], size=8.8, bold=True, color="primary_text",
           font_cn="微软雅黑", anchor="middle", line_spacing=1.12)

    # 标题下竖强调线（源页风格）
    line_x = col.x + 0.08
    s.line(line_x, col.y + 0.56, line_x, col.bottom - 0.06, BLUE, 1.35)

    body = Rect(col.x + 0.22, col.y + 0.60, col.w - 0.28, col.h - 0.68)
    rich_box(body, data["lines"], size=6.1, line_spacing=1.15)

deck.save(OUT)
print("已生成：%s（%d 页，%.2f x %.2f 英寸）"
      % (OUT, len(deck.slides), deck.canvas_w, deck.canvas_h))
check_overflow(OUT)
