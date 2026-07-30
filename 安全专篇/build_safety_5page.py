"""优化 安全(5页).pptx 的排版，输出 安全专篇/安全管理汇报(5页).pptx。

做法：以源文件为模板、保留全部 5 页，把第 1-4 页的形状清掉后按原内容重绘
（内容逐字照抄、风格沿用：白底 + 深蓝 + Arial 大号序号 + 竖分隔线 + 圆角框），
第 5 页只做两处改动 —— 中间蓝竖线变细、右下空白处补一块隐患台账方案总结。

为什么重绘而不是原地改字号：源文件的栏宽与各栏内容量不匹配（第 2 页 01 栏最宽
却也最长、03 栏最窄），只调字号治不了挤压；重绘后栏宽按字数加权、字号自动适配，
同页只有标题/正文两级字号。

修掉的具体缺陷：
  - 第 1 页正文色 E7E6E6 压白底，对比度 1.25:1，投屏几乎看不见 → 改深蓝 132556
  - 第 1 页 6 条事故只占上半页 → 排成 3x2 编号卡片，铺满内容区
  - 第 2/3 页正文 9pt 贴框边 → 栏宽重分 + 内边距统一 + 字号自动放大
  - 第 4 页中间主题标签框仅 0.37-0.45" 宽，中文被挤成竖排 → 加宽到 0.78" 横排
  - 中文只设了 latin 字体，东亚字会回落 → 统一写入 ea/cs

坐标统一用"实际英寸"书写（画布 10 x 5.625），A() 负责换算成 deckkit 的设计单位。
模板安全区按版式背景图实测：页眉止于 y 0.749"，页脚色条自 y 5.409" 起。

运行：python3 build_safety_5page.py [输出路径.pptx]
"""

from __future__ import annotations

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(REPO, ".cursor", "skills", "exec-deck-builder",
                               "scripts"))

from deckkit import (LIGHT, Deck, Rect, Slide, add_shadow,  # noqa: E402
                     apply_font, check_overflow, text_width, wrapped_lines)
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402
from pptx.util import Pt  # noqa: E402
from lxml import etree  # noqa: E402

SRC = os.path.join(REPO, "安全(5页).pptx")
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    HERE, "安全管理汇报(5页).pptx")

# 源文件配色，全部保留
NAVY = "132556"        # 页标题 / 小节标题，白底 14.7:1
BLUE = "1D42E1"        # 大号序号与强调，白底 7.2:1
INK = "252D30"         # 正文，白底 14.0:1
MUTED = "55677D"
LINE = "C7D3E6"
CARD = "F6F9FC"
BAND = "EDF2FA"

LS = 1.20              # 行距
LS_EST = LS * 1.22     # 估高留余量，末行才不会顶出框

T = LIGHT.variant(
    name="asymchem-safety5",
    bg="FFFFFF", bg_alt="F7FAFD", surface=CARD, surface_alt=BAND,
    hairline=LINE, ink=INK, ink_muted=MUTED, ink_on_accent="FFFFFF",
    primary=BLUE, secondary="2E5E9E", accent=BLUE,
    good="1A7F50", warn="A34E0F", bad="C0392B", neutral="5E7288",
    accent_text=BLUE, primary_text=NAVY, secondary_text="1F4E86",
    margin=0.37, margin_top=1.07, margin_bottom=0.35,
    font_cn="微软雅黑", font_en="Arial",
)

deck = Deck(theme=T, template=SRC, keep_slides=True)
SCALE = deck.scale                      # 0.75
PW, PH = 10.0, 5.625                    # 实际英寸

# 安全内容区（实际英寸）
SAFE_TOP, SAFE_BOTTOM = 0.80, 5.34
SAFE_L, SAFE_R = 0.30, 9.70


def A(v: float) -> float:
    """实际英寸 -> deckkit 设计英寸。"""
    return v / SCALE


def R(x, y, w, h) -> Rect:
    """按实际英寸建一个 Rect。"""
    return Rect(A(x), A(y), A(w), A(h))


# --------------------------------------------------------------------------
# 文本工具
# --------------------------------------------------------------------------

def clear_shapes(raw, keep_names=()):
    for sh in list(raw.shapes):
        if sh.name in keep_names:
            continue
        try:
            if "SLIDE_NUMBER" in str(sh.placeholder_format.type):
                continue
        except Exception:
            pass
        sh._element.getparent().remove(sh._element)


def set_bullet(paragraph, char="●", indent=0.11):
    """indent 单位为实际英寸。"""
    pPr = paragraph._p.get_or_add_pPr()
    for tag in ("a:buFont", "a:buChar", "a:buNone", "a:buAutoNum"):
        node = pPr.find(qn(tag))
        if node is not None:
            pPr.remove(node)
    emu = int(indent * 914400)
    pPr.set("marL", str(emu))
    pPr.set("indent", str(-emu))
    bf = etree.SubElement(pPr, qn("a:buFont"))
    bf.set("typeface", "Arial")
    bc = etree.SubElement(pPr, qn("a:buChar"))
    bc.set("char", char)


def plain(item) -> str:
    if isinstance(item, str):
        return item
    if item.get("runs"):
        return "".join(seg.get("text", "") for seg in item["runs"])
    return item.get("text", "")


def norm_lines(lines):
    return [{"text": x} if isinstance(x, str) else x for x in lines]


def block_height(lines, box_w, size, bullet_indent=0.11,
                 ls_est=LS_EST) -> float:
    """估算多段文本高度（实际英寸）。box_w 也是实际英寸。"""
    total = 0.0
    for item in norm_lines(lines):
        w = box_w - (bullet_indent if item.get("bullet") else 0.0)
        sz = size * item.get("scale", 1.0)
        total += wrapped_lines(plain(item), max(A(w), 0.05), A(sz)) \
            * sz * ls_est / 72.0
        total += item.get("space_before", 0) / 72.0
    return total


def fit_size(groups, hi, lo, step=0.1, bullet_indent=0.11,
             ls_est=LS_EST) -> float:
    """groups: [(lines, box_w, box_h)]，单位实际英寸。返回都装得下的最大字号。"""
    size = hi
    while size > lo:
        if all(block_height(ls, w, size, bullet_indent, ls_est) <= h
               for ls, w, h in groups):
            return round(size, 1)
        size -= step
    return lo


def text(s: Slide, x, y, w, h, content, size, bold=False, color="ink",
         align="left", anchor="top", ls=LS, font_cn=None, font_en=None):
    return s.text(R(x, y, w, h), content, size=A(size), bold=bold,
                  color=color, align=align, anchor=anchor, line_spacing=ls,
                  font_cn=font_cn, font_en=font_en)


def rich(s: Slide, x, y, w, h, lines, size, lead=0.0, anchor="top",
         bullet_indent=0.11, ls=LS):
    box = s.raw.shapes.add_textbox(s._i(A(x)), s._i(A(y)),
                                   s._i(A(w)), s._i(A(h)))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Pt(0)
    tf.margin_top = tf.margin_bottom = Pt(0)
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP,
                          "middle": MSO_ANCHOR.MIDDLE}[anchor]
    first = True
    for item in norm_lines(lines):
        is_first = first
        p = tf.paragraphs[0] if is_first else tf.add_paragraph()
        first = False
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = ls
        sb = item.get("space_before", 0) + (0 if is_first else lead)
        if sb:
            p.space_before = Pt(sb)
        if item.get("bullet"):
            set_bullet(p, indent=bullet_indent)
        runs = item.get("runs") or [{"text": item.get("text", "")}]
        for seg in runs:
            run = p.add_run()
            run.text = seg.get("text", "")
            # apply_font 直接写最终 pt；画布就是实际尺寸，所以这里用实际 pt，
            # 不能像 s.text() 那样再传设计 pt（deckkit 内部会替它乘 scale）
            apply_font(run.font, T,
                       size=size * item.get("scale", 1.0),
                       bold=seg.get("bold", item.get("bold", False)),
                       color=seg.get("color", item.get("color", "ink")))
    return box


def lead_for(lines, box_w, size, box_h, cap=4.0, bullet_indent=0.11,
             ls_est=LS_EST):
    slack = box_h - block_height(lines, box_w, size, bullet_indent, ls_est)
    gaps = max(len(lines) - 1, 1)
    return max(0.0, min(cap, slack * 72.0 * 0.6 / gaps))


def page_title(s: Slide, title: str):
    """页标题：右上，与源文件一致。"""
    text(s, 4.86, 0.10, 4.84, 0.42, title, 20, bold=True, color="primary_text",
         align="right", anchor="middle", font_cn="微软雅黑")


def numbered_head(s: Slide, x, y, w, num, head, num_size=26, head_size=11.5,
                  rule=True):
    """源文件的栏头样式：Arial 大号蓝序号 + 深蓝粗体标题 + 下方细线。"""
    text(s, x, y - 0.04, 0.62, 0.42, num, num_size, bold=True, color="primary",
         align="left", anchor="middle", font_en="Arial")
    text(s, x + 0.66, y, w - 0.66, 0.40, head, head_size, bold=True,
         color="primary_text", anchor="middle", font_cn="微软雅黑")
    if rule:
        s.line(A(x + 0.66), A(y + 0.40), A(x + w), A(y + 0.40), LINE, 1.0)
    return y + 0.48


def weighted_widths(groups, total, gap, floors=None):
    """按字数的 0.6 次幂分宽，并可施加每栏最小宽度。"""
    usable = total - gap * (len(groups) - 1)
    w8 = [sum(len(plain(i)) for i in norm_lines(g)) ** 0.6 for g in groups]
    out = [usable * v / sum(w8) for v in w8]
    if floors:
        out = [max(a, b) for a, b in zip(out, floors)]
        over = sum(out) - usable
        if over > 0:
            slack = [a - b for a, b in zip(out, floors)]
            pool = sum(slack)
            if pool > 0:
                out = [a - over * sl / pool for a, sl in zip(out, slack)]
    return out


# ==========================================================================
# 第 1 页：上半年安全管理反思 —— 6 起事故排成 3x2 编号卡片
# ==========================================================================
ACCIDENTS = ["200L桶静电事故", "应急响应事故", "硼烷四氢呋喃事故",
             "离心机氮封失效事故", "电缆事故", "电加热事故"]

s1 = Slide(deck, deck.prs.slides[0])
clear_shapes(s1.raw)
page_title(s1, "上半年安全管理反思")

grid = Rect(A(SAFE_L + 0.30), A(SAFE_TOP + 0.32),
            A(SAFE_R - SAFE_L - 0.60), A(SAFE_BOTTOM - SAFE_TOP - 0.52))
cells = grid.grid(3, 2, gap=A(0.26))
for i, (cell, name) in enumerate(zip(cells, ACCIDENTS), 1):
    card = s1.rect(cell, fill=CARD, line=LINE, line_w=1.0, radius=0.05)
    add_shadow(card, blur=7, dist=3, alpha=16)
    s1.rect(Rect(cell.x, cell.y, A(0.05), cell.h), fill=BLUE, radius=0)
    s1.text(Rect(cell.x + A(0.22), cell.y + A(0.16), A(0.70), A(0.42)),
            "%02d" % i, size=A(24), bold=True, color="primary",
            font_en="Arial", anchor="middle")
    s1.text(Rect(cell.x + A(0.22), cell.y + A(0.62),
                 cell.w - A(0.44), A(0.46)), name,
            size=A(14), bold=True, color="primary_text",
            font_cn="微软雅黑", line_spacing=1.16)


# ==========================================================================
# 第 2 页：重塑安全管理体系 01-03
# ==========================================================================
BASIS = ("依据：全员安全生产责任制、双重预防机制（风险分级管控+隐患排查治理）、"
         "安全生产标准化、安全绩效考核")

COLS2 = [
    ("01", "安全隐患排查治理体系", [
        {"text": "管控目标：实现隐患全流程闭环管控，杜绝隐患诱发安全事故",
         "bold": True, "color": "accent"},
        {"text": "落地措施：", "bold": True, "space_before": 3},
        {"text": "建立全员隐患排查、上报、奖励机制，调动全员安全积极性；",
         "bullet": True},
        {"text": "落实多维度排查：结合现有巡查，隐患排查体系和安全隐患台账系统，"
                 "包括：驻厂巡查，值班巡查，EHS安全隐患排查，结合班组自查，"
                 "车间主任隐患排查，部门隐患排查，厂房隐患排查，厂长全厂区隐患排查，"
                 "专项检查，节假日安全检查", "bullet": True},
        {"text": "隐患分级管控：划分⼀般、较大、重大隐患，分类建档、限期整改；"
                 "设立专人专岗管理巡查系统（隐患台账系统），及时汇总，分析隐患趋势，"
                 "追踪落地情况", "bullet": True},
        {"text": "“常态化排查、台账化管理、闭环式整改”隐患治理机制；"
                 "整改完成后必须复查复检，确保隐患彻底清零； "
                 "集团EHS监督复查属地隐患清零情况", "bullet": True},
        {"text": "风险预判机制：工艺变更，操作方式变更，季节变化（如湿度，温度），"
                 "人员变更（包括新人增加），高产能利用率，提前进行风险防控",
         "bullet": True},
        {"text": "设立生产安全委员会，专注隐患排查，整改方案审定，形成标准化整改方案；"
                 "在隐患解决措施上，更多考虑自动化手段；各厂长轮值安委会负责人，"
                 "暂定每季度轮换", "bullet": True},
    ]),
    ("02", "危险化学品全生命周期管理体系", [
        {"text": "管控目标：实现危险化学品采购、储存、使用、输送、装卸、"
                 "废弃处置全过程风险可控", "bold": True, "color": "accent"},
        {"text": "核心落地措施，在现有管理基础上，增加如下措施：", "bold": True,
         "space_before": 3},
        {"text": "1.危险化学品准入管理，配齐MSDS，危害，历史事故资料，"
                 "提前开展风险辨识与安全交底，并对全链条人员培训；"
                 "公司MSDS管理升级，有来源，有审核；"},
        {"text": "2.新类型危险化学品，包括包装，大小，规格变更等，采购前，"
                 "由集团危险化学品管理负责人审批； 建立集团危险化学品负责人岗责；"},
        {"text": "3.工厂危险化学品使用，到货计划，包括研发使用危险化学品，"
                 "由总工负责制定，工厂危险化学品管理负责人审批，"
                 "夏季库存量满足1周使用，降低库存量；建立工厂危险化学品负责人岗责；"},
        {"text": "4.危险化学品仓库增加消防级别，增加仓库及单一危险化学品专项应急预案，"
                 "并演练；方案由EHS制定；"},
        {"text": "5.  项目结束后危险化学品立即报废处理；"},
    ]),
    ("03", "设备设施安全管理体系", [
        {"text": "管控目标：保障生产设备、安全设施完好有效，"
                 "杜绝设备故障和失效引发事故；", "bold": True, "color": "accent"},
        {"text": "核心落地措施，在现有安全管理措施下，增加如下：", "bold": True,
         "space_before": 3},
        {"text": "1.设立工厂关键设备设施MSE，熟悉设备构造，维护，常见事故，日常巡查；"},
        {"text": "2.  设立工厂设备安全巡检员，进一步完善岗责；"
                 "负责全厂生产设备、配套辅助设施日常巡回安全检查，"
                 "排查设备机械、电气、液压、气动、防护装置隐患，跟踪隐患整改，"
                 "预防设备故障、机械伤害、设备起火等安全事故，"
                 "落实设备安全常态化管控；"},
        {"text": "3.各厂区尽快上线设备管理系统，将设备故障，安全隐患台账，"
                 "隐患解决纳入设备安全常态化管控。"},
    ]),
]

s2 = Slide(deck, deck.prs.slides[1])
clear_shapes(s2.raw)
page_title(s2, "重塑安全管理体系")

# 依据条：源文件是一行居中粗体，这里给它一条浅底，作为全页前提
band = R(SAFE_L, SAFE_TOP, SAFE_R - SAFE_L, 0.34)
s2.rect(band, fill=BAND, line=LINE, line_w=1.0, radius=0.06)
s2.rect(Rect(band.x, band.y, A(0.05), band.h), fill=BLUE, radius=0)
text(s2, SAFE_L + 0.18, SAFE_TOP, SAFE_R - SAFE_L - 0.30, 0.34, BASIS,
     10.5, bold=True, color="primary_text", anchor="middle",
     font_cn="微软雅黑")

BODY_TOP = SAFE_TOP + 0.34 + 0.14
GAP2, PAD2 = 0.26, 0.16
HEAD2 = 0.48
widths2 = weighted_widths([c[2] for c in COLS2], SAFE_R - SAFE_L, GAP2)
size2 = fit_size(
    [(c[2], w - 2 * PAD2, SAFE_BOTTOM - BODY_TOP - HEAD2 - 0.16)
     for c, w in zip(COLS2, widths2)], hi=10.5, lo=8.0)

x = SAFE_L
for (num, head, lines), w in zip(COLS2, widths2):
    col = R(x, BODY_TOP, w, SAFE_BOTTOM - BODY_TOP)
    card = s2.rect(col, fill=CARD, line=LINE, line_w=1.0, radius=0.04)
    add_shadow(card, blur=6, dist=3, alpha=14)
    by = numbered_head(s2, x + PAD2, BODY_TOP + 0.08, w - 2 * PAD2, num, head)
    bh = SAFE_BOTTOM - by - 0.14
    rich(s2, x + PAD2, by, w - 2 * PAD2, bh, lines, size2,
         lead=lead_for(lines, w - 2 * PAD2, size2, bh))
    x += w + GAP2


# ==========================================================================
# 第 3 页：重塑安全管理体系 04-06 + 07-09 待展开
# ==========================================================================
COLS3 = [
    ("04", "人员安全教育培训体系", [
        {"text": "管控目标：杜绝人的不安全行为，提升全员岗位安全履职能力；",
         "bold": True, "color": "accent"},
        {"text": "核心落地措施：", "bold": True, "space_before": 3},
        {"text": "新员工严格落实厂级、车间级、班组级三级安全教育培训；",
         "bullet": True},
        {"text": "高风险设备，高风险操作提升自动化水平，并提升保护层级，"
                 "减少人员直接参与操作；", "bullet": True},
        {"text": "增加专项培训，增加常态化复训频率，包括一线员工，车间管理人员，"
                 "EHS安全相关岗位人员；", "bullet": True},
        {"text": "车间一线员工，必须是化工，制药相关专业背景，大专及以上学历，"
                 "现有不满足学历和专业要求员工，在半年时间内进行培训，考核，"
                 "不通过进行调岗；", "bullet": True},
        {"text": "针对安全管理人员开展专项安全管理能力提升培训；", "bullet": True},
        {"text": "在工厂增加安全培训专岗，统筹和组织安全培训实施，"
                 "支持工厂安全培训体系工作开展", "bullet": True},
        {"text": "车间生产组织形式变更，讨论；在小分子生产车间，"
                 "实施按工段定岗为主，建立工段定岗责任制，"
                 "正常生产状态操作人员固定所属工段，承担本工段设备操作、巡检、"
                 "属地安全管理；车间项目变化很快，工艺变化很快，将实施危险反应，"
                 "危险工艺，危险化学品的反应釜，设备相对固定，相应人员固定；",
         "bullet": True},
    ]),
    ("05", "应急管理体系", [
        {"text": "管控目标：快速处置突发事故，最大限度降低人员伤亡和财产损失；",
         "bold": True, "color": "accent"},
        {"text": "健全综合应急预案、专项应急预案、岗位现场处置方案三级预案体系",
         "bullet": True},
        {"text": "应急物资与专业应急救援物资对齐，对消防救援材料，"
                 "与当地消防部门联动，进行储备；", "bullet": True},
        {"text": "组建工厂专职、兼职应急救援队伍，明确岗位职责，"
                 "车间义务消防员比例大于40%，确保每班至少两名义务消防员配置；",
         "bullet": True},
        {"text": "每年至少开展两次泄漏、火灾、爆炸、中毒等实战化应急演练，"
                 "该演练按车间，班组，厂内救援队伍为单位实施；"
                 "每人均需实操相应设备，器具；", "bullet": True},
        {"text": "每次演练后开展复盘评估，优化预案及处置流程；", "bullet": True},
        {"text": "建立与消防、医院、应急管理部门的外部应急联动机制；",
         "bullet": True},
        {"text": "规范事故上报、现场警戒、人员疏散、现场处置全流程。",
         "bullet": True},
        {"text": "工厂按班组配备足够对讲机，降低防爆手机使用比例；",
         "bullet": True},
    ]),
    ("06", "安全管理制度落地和执行监督体系", [
        {"text": "管控目标：破除制度纸面化，实现安全制度刚性落地、有效执行",
         "bold": True, "color": "accent"},
        {"text": "全员安全责任制：各岗位具体安全职责、工作标准和追责机制，"
                 "签订安全生产责任书，实现“人人有责、人人尽责、失职追责”。",
         "bullet": True},
        {"text": "EHS，车间管理人员增加现场巡检频次和力度，对违章、违规行为零容忍，"
                 "发现，制止，处罚，尽快全员通报，震慑，"
                 "确保各项安全制度、操作规程落地执行;", "bullet": True},
        {"text": "建立从班组，车间，厂房，厂长，EHS的多级监督机制；", "bullet": True},
        {"text": "制定工厂安全管理人员，EHS安全巡查人员比例要求，工厂按该要求实施；",
         "bullet": True},
        {"text": "对车间管理人员，EHS安全巡查人员增加明确隐患，违章排查每月KPI，"
                 "必须发现多少安全隐患和违章行为；", "bullet": True},
        {"text": "畅通员工安全监督、隐患举报渠道，鼓励全员参与安全管理；",
         "bullet": True},
    ]),
]

PENDING3 = [("07", "特殊作业安全管控体系"),
            ("08", "职业卫生与劳动保护"),
            ("09", "工艺安全")]

s3 = Slide(deck, deck.prs.slides[2])
clear_shapes(s3.raw)
page_title(s3, "重塑安全管理体系")

GAP3, PAD3 = 0.24, 0.16
HEAD3 = 0.48
PEND_H = 0.86          # 右下"07-09 待展开"区高度
COL3_TOP = SAFE_TOP + 0.06

# 06 栏下方要让出 07-09 区，所以三栏宽度按"可用高度不同"分别估
widths3 = weighted_widths([c[2] for c in COLS3], SAFE_R - SAFE_L, GAP3)
heights3 = [SAFE_BOTTOM - COL3_TOP] * 3
heights3[2] = SAFE_BOTTOM - COL3_TOP - PEND_H - 0.16
size3 = fit_size(
    [(c[2], w - 2 * PAD3, h - HEAD3 - 0.16)
     for c, w, h in zip(COLS3, widths3, heights3)], hi=10.5, lo=8.0)

x = SAFE_L
col3_rects = []
for (num, head, lines), w, h in zip(COLS3, widths3, heights3):
    col = R(x, COL3_TOP, w, h)
    col3_rects.append((x, w, h))
    card = s3.rect(col, fill=CARD, line=LINE, line_w=1.0, radius=0.04)
    add_shadow(card, blur=6, dist=3, alpha=14)
    by = numbered_head(s3, x + PAD3, COL3_TOP + 0.08, w - 2 * PAD3, num, head)
    bh = COL3_TOP + h - by - 0.14
    rich(s3, x + PAD3, by, w - 2 * PAD3, bh, lines, size3,
         lead=lead_for(lines, w - 2 * PAD3, size3, bh))
    x += w + GAP3

# 07-09：源文件只有标题、没有正文，做成一个"待展开"区块，与三栏同一视觉语言
px, pw, ph = col3_rects[2]
pend = R(px, SAFE_BOTTOM - PEND_H, pw, PEND_H)
s3.rect(pend, fill=BAND, line=LINE, line_w=1.0, radius=0.04)
rows = pend.split_v(1, 1, 1, gap=A(0.02))
for (num, head), row in zip(PENDING3, rows):
    s3.text(Rect(row.x + A(0.16), row.y, A(0.60), row.h), num,
            size=A(15), bold=True, color="primary", font_en="Arial",
            anchor="middle")
    s3.text(Rect(row.x + A(0.78), row.y, row.w - A(0.94), row.h), head,
            size=A(10.5), bold=True, color="primary_text",
            font_cn="微软雅黑", anchor="middle")


# ==========================================================================
# 第 4 页：200L 桶事故整改 —— 保留中轴鱼骨结构，加宽主题标签
# ==========================================================================
LEFT4 = [
    ("车间200L桶&方槽使用管理", [
        "TJ1、TJ4、公斤级实验室禁止使用可燃溶剂200L包装进入车间",
        "全面取缔使用桶或槽接收动态的离心或压力操作的母液",
        "杜绝使用塑料，四氟材料桶装酸性废液；不具备条件厂区，升级塑料桶放桶操作硬件",
        "尽最大可能减少周转桶进入车间转移废水，废水，废溶剂；",
        "明确废液间的标准配置",
        "建立自动化充氮及自动化灌装设备；配备氧含量监测，压力表",
        "200L桶加料标准化：氮气流速、泵的选型、尾气连接方式、插底管安装标准等。",
        "公司制度要求：酸性溶剂使用塑料/四氟桶需厂长审批",
    ]),
    ("衬塑设备管理", [
        "TJ1、TJ4全部更换为铪氏合金离心机/压滤罐；其它厂区新采购设备更换为铪氏合金设备。"
        "已安装衬塑设备工厂评估是否更换，提供整改概算和整改时间表；如果不更换，"
        "保护措施是否足够，集团EHS审批。",
        "工艺中杜绝酸性体系离心、过滤后严禁使用纯溶剂直接淋洗。如果必须使用纯溶剂淋洗，"
        "所有设备、管路及接收设备必须使用铪氏合金材质。",
        "取缔衬塑插底管，更换为铪氏合金插底管，连接静电接地；",
        "GMP车间减少塑料管使用；酸性体系转移，使用内衬halo材质管路；",
        "反应釜配置滴加罐",
    ]),
    ("氮气使用管理", [
        "储罐，设备上必须使用白钢氮气管路，禁止使用软管连接",
        "车间内氮气接口不足问题，工厂整改，增加氮气源接口，或者使用氮气排解决；"
        "每条氮气管路上，必须有流量计；",
        "使用氮气的设备，储罐，桶及其它工况（有一定耐压能力），必须有压力表；",
        "使用氮气房间，必须有氧含量监测和报警设备；",
        "TJ1&TJ4评估柔性隔离器氮气窒息和周围环境氮气窒息风险；",
        "每年至少两次使用惰性气体的安全排查",
    ]),
    ("设备管理", [
        "工厂内关键设备，高风险设备增加SME负责人，清楚掌握设备工作原理，安全，"
        "质量控制关键环节，影响因素，跟踪设备维保，故障情况，参与变更审批等；",
        "增加厂房内白班和夜班设备安全巡查员",
    ]),
    ("物料和生产禁忌要求", [
        "车间内还原剂，氧化剂，剧毒，易爆，自燃化学品使用，均应受EHS专人监督；"
        "车间使用后相关尾料，废液必须及时处理，禁止未破坏暂存；",
        "对于酸，碱，氧化剂化学品，与所使用设备，配件材质兼容性检查",
    ]),
]

RIGHT4 = [
    ("车间生产秩序", [
        "公司内所有监控，不能被遮挡；推进厂房内人员定位系统安装；",
        "项目步骤生产启动，开工前检查：设备准备、工艺风险、管道、配套附属设备、"
        "安全措施是否按公司流程执行到位，建立Checklist；EHS副主管，"
        "车间主任为最低级别要求；",
        "车间现场，一个工序完成后，进行现场清理，杜绝物料，包装桶堆积在现场；",
        "备料环节，减少车间物料存放，特别是可燃溶剂，易燃，自燃，易爆化学品",
        "高风险操作，必须是双岗操作，相互监督，复核；",
        "操作空间要求，对于狭小房间内操作，必须进行完一个操作，清场后，进入下一个操作",
        "TJ1整改一号楼，二号楼，三号楼操作空间风险，制定整改时间表；",
    ]),
    ("消防应急", [
        "溶剂、可燃物料（非水敏感）在车间内的使用、存放区域包括走廊和连廊、"
        "卸料平台必须有喷淋及其他有效消防设备；",
        "在高风险区域增加悬挂灭火器，但不能替代消防喷淋；",
        "检查应急处置，相应演练是否有效，包括车间内各类设备如果发生泄漏、火灾；"
        "各厂区隔离器内发生着火；实验室通风处内着火等",
        "检查相关风险隐患整改是否到位，措施是否有效处置措施。",
        "建立义务消防员管理制度，增加车间内义务消防员工作服标志；",
        "可燃溶剂，易燃易爆化学物使用区域，义务消防员比例不低于40%",
        "每年开展两次全员灭火实操考核，包括灭火器，消火栓，灭火毯；",
        "各工厂建立专门的消防培训，实操，考核场地，用于人员培训，不仅限消防相关，"
        "包括化学品泄漏，中毒演练；工厂设置专门培训岗位；",
        "推进防静电阻燃工作服，连体服替换，推进全覆盖安全帽替换工作；",
    ]),
    ("培训", [
        "建立年度安全培训专篇，对于各类安全培训，设备培训，操作培训等，"
        "进行明确频率，考核要求；涵盖生产、EHS不同级别管理人员考核清单",
        "生产管理人员，每年至少1次各类设备实操考核；重要安全管理要求考核；",
        "EHS现场巡查人员，安全相关人员，每年至少一次设备实操考核；",
        "未经培训和实操考核的人员，禁止进行操作，特别是项目人员；",
        "所有人员考核视频存档",
        "探讨EHS人员和生产人员轮岗办法",
    ]),
    ("巡查管理", [
        "巡查条例更新，增加巡查案例，并对巡查人员进行培训，"
        "并重新考察车间巡查人员范围，必须是对于生产操作，车间风险熟悉的人员",
        "对于巡查，检查走过场的人员，首次警告，第二次扣分，第三次撤销管理职务",
    ]),
]

BOTTOM4 = ("完善属地安全管理分工，压实安全管理责任", [
    "风险越高，安全管理责任更得压实；属地安全管理责任，压实到一线，"
    "同时EHS对于高风险区域，设备，操作等，也进行责任分区，压实巡查，监督责任；",
    "EHS牵头，建立分级巡查制度，完善安全管理体系",
])

s4 = Slide(deck, deck.prs.slides[3])
clear_shapes(s4.raw)
page_title(s4, "200L 桶事故整改")

# 这页字最多（约 2600 字）。标签列要比源文件宽才能横排中文，宽度只能从内容区
# 借；为了不把字号压到比源稿（7pt）还小，正文行距与内边距都收紧。
TAG_W = 0.58            # 源文件只有 0.37-0.45，中文被挤成竖排
AXIS_GAP = 0.06
BOT_H = 0.38
TOP4 = SAFE_TOP - 0.02
BODY4_BOTTOM = SAFE_BOTTOM - BOT_H - 0.05
LS4, LS4_EST = 1.10, 1.14
BUL4 = 0.07

content_w = (SAFE_R - SAFE_L - 2 * TAG_W - 2 * AXIS_GAP - 0.06) / 2
LX = SAFE_L
LTAG = LX + content_w + AXIS_GAP
AXIS = LTAG + TAG_W + 0.03
RTAG = AXIS + 0.03
RX = RTAG + TAG_W + AXIS_GAP

# 中轴
s4.line(A(AXIS), A(TOP4 + 0.06), A(AXIS), A(BODY4_BOTTOM - 0.02), BLUE, 1.25)

col_h = BODY4_BOTTOM - TOP4
PAD4 = 0.07
BLOCK_GAP4 = 0.04


def side_needs(side, size):
    """每个主题块按内容实际所需高度（含内边距）。"""
    return [block_height([{"text": t, "bullet": True} for t in items],
                         content_w - 2 * PAD4, size,
                         bullet_indent=BUL4, ls_est=LS4_EST)
            + 2 * PAD4 for _, items in side]


def side_fits(size):
    for side in (LEFT4, RIGHT4):
        need = sum(side_needs(side, size)) + BLOCK_GAP4 * (len(side) - 1)
        if need > col_h:
            return False
    return True


# 块高按"内容需要多少"分配，而不是按条数平均 —— 长条目多的块才拿到更多高度
size4 = 8.6
while size4 > 6.0 and not side_fits(size4):
    size4 = round(size4 - 0.1, 1)


def draw_side(side, cx, tagx):
    """一侧的若干主题块：内容块 + 标签块 + 标签到中轴的引出线。"""
    needs = side_needs(side, size4)
    total = col_h - BLOCK_GAP4 * (len(side) - 1)
    slack = max(0.0, total - sum(needs))
    heights = [n + slack * n / sum(needs) for n in needs]
    y = TOP4
    for (tag, items), h in zip(side, heights):
        box = R(cx, y, content_w, h)
        s4.rect(box, fill="FFFFFF", line=LINE, line_w=0.9, radius=0.03)
        bh = h - 2 * PAD4
        lines = [{"text": t, "bullet": True} for t in items]
        rich(s4, cx + PAD4, y + PAD4, content_w - 2 * PAD4, bh, lines, size4,
             lead=lead_for(lines, content_w - 2 * PAD4, size4, bh, cap=2.0,
                           bullet_indent=BUL4, ls_est=LS4_EST),
             bullet_indent=BUL4, ls=LS4)

        th = min(0.70, h - 0.06)
        trect = R(tagx, y + (h - th) / 2, TAG_W, th)
        s4.rect(trect, fill=BAND, line=BLUE, line_w=0.9, radius=0.05)
        s4.text(trect.inset(A(0.05), A(0.02)), tag, size=A(size4 + 0.4),
                bold=True, color="primary_text", align="center",
                anchor="middle", line_spacing=1.12, font_cn="微软雅黑")
        # 标签到中轴的引出线
        if tagx < AXIS:
            s4.line(A(tagx + TAG_W), A(y + h / 2), A(AXIS),
                    A(y + h / 2), LINE, 0.9)
        else:
            s4.line(A(AXIS), A(y + h / 2), A(tagx), A(y + h / 2), LINE, 0.9)
        y += h + BLOCK_GAP4


draw_side(LEFT4, LX, LTAG)
draw_side(RIGHT4, RX, RTAG)

# 底部横跨块
bot = R(SAFE_L, SAFE_BOTTOM - BOT_H, SAFE_R - SAFE_L, BOT_H)
s4.rect(bot, fill=BAND, line=BLUE, line_w=0.9, radius=0.04)
btag_w = 2.30
s4.text(R(SAFE_L + 0.14, SAFE_BOTTOM - BOT_H, btag_w, BOT_H), BOTTOM4[0],
        size=A(size4 + 0.6), bold=True, color="primary_text",
        anchor="middle", line_spacing=1.12, font_cn="微软雅黑")
s4.line(A(SAFE_L + 0.14 + btag_w + 0.06), A(SAFE_BOTTOM - BOT_H + 0.06),
        A(SAFE_L + 0.14 + btag_w + 0.06), A(SAFE_BOTTOM - 0.06), LINE, 0.9)
# 两条要点并排，宽度按字数线性分配：等分会把长的那条折成 3 行、超出这条底栏
bx = SAFE_L + btag_w + 0.30
bw_total = SAFE_R - bx - 0.14 - 0.18
chars = [len(t) for t in BOTTOM4[1]]
bxs = bx
for t, c in zip(BOTTOM4[1], chars):
    bw = bw_total * c / sum(chars)
    rich(s4, bxs, SAFE_BOTTOM - BOT_H + 0.04, bw, BOT_H - 0.08,
         [{"text": t, "bullet": True}], size4,
         anchor="middle", bullet_indent=BUL4, ls=LS4)
    bxs += bw + 0.18


# ==========================================================================
# 第 5 页：只改两处 —— 竖线变细、右下补台账方案总结
# ==========================================================================
s5 = Slide(deck, deck.prs.slides[4])

for sh in list(s5.raw.shapes):
    if sh.name == "Rectangle 54":
        # 源竖线 0.088" 宽、纯色，偏粗。收到 0.026" 并改用蓝灰细线，上下各收 0.06
        sh.left = s5._i(A(4.63))
        sh.top = s5._i(A(1.17))
        sh.width = s5._i(A(0.026))
        sh.height = s5._i(A(4.05))
        sh.fill.solid()
        from pptx.dml.color import RGBColor
        sh.fill.fore_color.rgb = RGBColor.from_string("9FB6D4")
        sh.line.fill.background()

LEDGER_HEAD = "安全隐患台账标准化管理方案"
LEDGER_SUB = "隐患全生命周期数字化载体：源头可查 · 过程可控 · 趋势可分析 · 责任可追溯"
LEDGER_POINTS = [
    ("统一台账", "巡检、自查、督查隐患同库登记，19 项必填字段，每条绑定整改前后照片"),
    ("闭环管控", "未整改→整改中→复查待验收→已闭环→延期整改；到期前 3 天预警、"
                 "超期推送、无佐证不予销号"),
    ("分级督办", "重大/较大/一般由安全部门统一定级，重大隐患挂牌督办并落实临时管控"),
    ("数据运用", "月度输出区域、类型与重复性隐患分析，支撑考核、专项行动与风险清单更新"),
    ("硬性规则", "禁造假、禁拆分重大隐患、权限分离留痕；电子台账永久保存，纸质≥3 年"),
]

# 上方图注到 y 4.10，页脚数据来源在 y 5.18，所以这块只能占 4.14 - 5.12
LED_TOP, LED_H = 4.14, 0.98
LED = R(4.82, LED_TOP, 4.83, LED_H)
s5.rect(LED, fill=BAND, line=LINE, line_w=1.0, radius=0.05)
s5.rect(Rect(LED.x, LED.y, A(0.045), LED.h), fill=BLUE, radius=0)

lx, lw = 4.82 + 0.15, 4.83 - 0.28
text(s5, lx, LED_TOP + 0.04, lw, 0.16, LEDGER_HEAD, 8.6, bold=True,
     color="primary_text", font_cn="微软雅黑")
text(s5, lx, LED_TOP + 0.18, lw, 0.15, LEDGER_SUB, 6.4, color="ink_muted",
     font_cn="微软雅黑")

led_lines = [{"runs": [{"text": k + "：", "bold": True, "color": "accent"},
                       {"text": v, "color": "ink"}], "bullet": True}
             for k, v in LEDGER_POINTS]
LED_BODY_TOP = LED_TOP + 0.29
LED_BODY_H = LED_TOP + LED_H - LED_BODY_TOP - 0.03
led_size = fit_size([(led_lines, lw, LED_BODY_H)], hi=7.2, lo=5.4, step=0.05,
                    bullet_indent=0.07, ls_est=1.14)
rich(s5, lx, LED_BODY_TOP, lw, LED_BODY_H, led_lines, led_size,
     lead=lead_for(led_lines, lw, led_size, LED_BODY_H, cap=1.0,
                   bullet_indent=0.07, ls_est=1.14),
     bullet_indent=0.07, ls=1.12)


deck.prs.save(OUT)
print("已生成：%s（%d 页，%.2f x %.2f 英寸）"
      % (OUT, len(deck.prs.slides._sldIdLst), PW, PH))
print("字号（实际 pt）：第2页 %.1f，第3页 %.1f，第4页 %.1f，台账块 %.2f"
      % (size2, size3, size4, led_size))
check_overflow(OUT)
