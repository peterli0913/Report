"""安全生产专题汇报（凯莱英模板）。

内容来自 安全专篇1页.pptx：本年度一起二级事故、由该事故触发的隐患倒查结果、
海因里希法则、以及安全生产的五大保障基石。

叙事主线：海因里希法则说 1 起严重事故背后有 300 起隐患信号；集团巡检系统里恰好
躺着约 300 条同类隐患记录且未被根本整改。所以这起事故不是偶然，而是 300 次预警
被忽略后的必然结果 —— 五大保障基石都需要被真正落实。

页面清单：
  1  封面        一起事故，背后是 300 次未被听见的预警
  2  核心结论    这起事故不是偶然，预警信号早已齐备
  3  事故回顾    本次事故基本情况
  4  海因里希法则 事故是概率的必然，不是运气的偶然
  5  倒查结果    约 300 条同类隐患，半年未根治
  6  问题本质    闭环断在整改与验证之间
  7  保障基石    五大基石均未被充分保障
  8  行动要求    请各层级立即落实
  9  结尾        下一个 300 条隐患里，可能藏着下一起事故

标有【待填】的位置需要按实际情况补充，原始素材里就是占位符。
中文内嵌强调统一用全角引号「」，避免与 Python 字符串定界冲突。

运行：
    python3 build_safety_deck.py [输出路径.pptx]
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
SK = os.path.join(REPO, ".cursor", "skills")
sys.path.insert(0, os.path.join(SK, "exec-deck-builder", "scripts"))
sys.path.insert(0, os.path.join(SK, "deck-imagery", "scripts"))

from deckkit import (BASE_H, BASE_W, LIGHT, Deck, Rect,  # noqa: E402
                     audit_theme, check_overflow, text_height)
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402

TEMPLATE = os.path.join(REPO, "安全专篇1页.pptx")
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "安全生产专题汇报.pptx")
ICON_DIR = os.path.join(HERE, "icons")

# --------------------------------------------------------------------------
# 主题：沿用凯莱英白底模板的配色，警示红取自原 PPT 强调数字所用的 C00000
# --------------------------------------------------------------------------
# 边距按模板背景图 image7.jpg 实测的安全区设定：页眉 logo 占到 y 2.13 英寸
# （设计单位 1.06），页脚深蓝条从 y 14.37 起（设计 7.18），左边缘对齐 logo。
T = LIGHT.variant(
    name="asymchem-safety",
    bg="FFFFFF", bg_alt="F7FAFD",
    surface="EDF3FA", surface_alt="DCE7F4",
    hairline="C2D3E6",
    ink="16283F", ink_muted="4E6584", ink_on_accent="FFFFFF",
    primary="003669", secondary="3263A7",
    # 安全汇报里红色就是唯一的警示语言，accent 与 bad 同色，语义不分叉
    accent="C00000", bad="C00000",
    good="1A7F50", warn="B45309", neutral="5E7288",
    accent_text="C00000", bad_text="C00000",
    good_text="167348", warn_text="A04C0D", neutral_text="4E6584",
    primary_text="003669", secondary_text="1F4E86",
    series=("003669", "C00000", "3263A7", "B45309", "1A7F50", "5E7288"),
    margin=0.58, margin_top=1.30, margin_bottom=0.52,
    size_title=24, size_h=16.5, size_body=14.5, size_small=12.5,
    size_kpi=46, size_note=10,
)

for _p in audit_theme(T):
    print("配色警告：", _p)

# 模板里的两个内容页版式：image7 干净、image8 右下带淡灰几何装饰
LAYOUT_PLAIN = "2_自定义版式"
LAYOUT_DECOR = "3_自定义版式"

deck = Deck(theme=T, template=TEMPLATE, layout=LAYOUT_PLAIN)

# 图标：离线几何绘制，不依赖外部素材
ICONS = False
try:
    from make_icon import draw_icon
    for _n in ("people", "alert", "flask", "search", "equipment"):
        draw_icon(_n, os.path.join(ICON_DIR, _n + ".png"), color="003669", size=300)
    ICONS = True
except Exception as _e:                       # 缺 Pillow 时退回色块
    print("图标生成跳过（%s），改用色块" % _e)


def icon(name):
    p = os.path.join(ICON_DIR, name + ".png")
    return p if ICONS and os.path.exists(p) else None


TOTAL = 9
# 页脚放在模板深蓝色条（实际 y 14.37 起）上方的白底区域，用灰字。
# 压在深蓝条上虽然也能看，但那条本身带渐变，叠色块很难对齐得不露痕迹。
FOOT_Y = BASE_H - T.mb - 0.22
# 模板右下角固定有 Confidential 标（实际 x 21.70 起 = 设计 10.85），页脚要让开
FOOT_W = 9.30


def footer(s, n, note=None):
    """页脚：模板右下已有 Confidential 保密标，这里只补数据来源与页码。"""
    if note:
        s.text(Rect(T.margin, FOOT_Y, FOOT_W, 0.24), note, size=T.size_note,
               color="ink_muted")
    s.text(Rect(T.margin + FOOT_W + 0.10, FOOT_Y, 0.85, 0.24),
           "%d / %d" % (n, TOTAL), size=T.size_note, color="ink_muted",
           align="right")


# --------------------------------------------------------------------------
# 1 封面
# --------------------------------------------------------------------------
s = deck.raw_slide(layout=LAYOUT_DECOR)
body = s.safe
s.text(Rect(body.x, body.y + 0.55, body.w, 0.34), "安全生产专题汇报",
       size=13, bold=True, color="accent")
s.rect(Rect(body.x, body.y + 1.05, 0.075, 1.30), fill="accent", radius=0)
s.text(Rect(body.x + 0.30, body.y + 1.02, body.w - 0.30, 0.72),
       "一起事故，背后是 300 次未被听见的预警", size=34, bold=True,
       color="primary", line_spacing=1.16)
s.text(Rect(body.x + 0.30, body.y + 1.80, body.w * 0.74, 0.56),
       "从本年度一起二级事故，看隐患整改闭环的失效", size=17,
       color="ink_muted", line_spacing=1.34)
s.line(body.x, 6.50, body.right, 6.50, "hairline", 0.8)
s.text(Rect(body.x, 6.62, body.w, 0.30),
       "汇报单位：【待填】    |    汇报人：【待填】    |    汇报日期：【待填】",
       size=12, color="ink_muted")

# --------------------------------------------------------------------------
# 2 核心结论：法则与实际数据的量级对照
# --------------------------------------------------------------------------
s = deck.slide(title="这起事故不是偶然 —— 预警信号早已齐备，缺的是整改闭环",
               eyebrow="核心结论")
top, bot = s.body.split_v(0.72, 0.28, gap=0.24)
left, right = top.split_h(1, 1, gap=0.34)

# 左：海因里希法则
s.rect(left, fill="surface")
inner = left.inset(0.30, 0.26)
s.text(Rect(inner.x, inner.y, inner.w, 0.30), "海因里希法则告诉我们",
       size=T.size_h, bold=True, color="primary")
# 三条等宽。曾经试过用宽度递增表示 1 : 29 : 300，但递增比例既装不下描述文字，
# 又无法真实反映 300 倍的量级差，反而误导 —— 比例关系交给第 4 页的金字塔表达。
rows = [("1", "起", "严重伤害事故（死亡或重伤）", "accent"),
        ("29", "起", "轻微伤害事故", "warn"),
        ("300", "起", "无伤害的未遂事件或事故隐患", "primary")]
y = inner.y + 0.52
NUM_W = 1.20
for val, unit, desc, col in rows:
    s.rect(Rect(inner.x, y, inner.w, 0.62), fill=col, radius=0.03)
    s.text(Rect(inner.x + 0.20, y, NUM_W, 0.62),
           [(val, {"size": 25, "bold": True, "color": "FFFFFF"}),
            (" " + unit, {"size": 11.5, "color": "FFFFFF"})], anchor="middle")
    dx = inner.x + 0.20 + NUM_W + 0.16
    s.text(Rect(dx, y, inner.right - dx - 0.20, 0.62), desc,
           size=12.5, color="FFFFFF", anchor="middle", line_spacing=1.3)
    y += 0.74
s.text(Rect(inner.x, y - 0.06, inner.w, 0.28),
       "每一起严重事故下面，都垫着约 300 个被忽略的信号",
       size=12, color="ink_muted", align="center")

# 右：我们的实际数据
s.rect(right, fill="surface")
inner = right.inset(0.30, 0.26)
s.text(Rect(inner.x, inner.y, inner.w, 0.30), "我们的实际情况",
       size=T.size_h, bold=True, color="primary")
s.text(Rect(inner.x, inner.y + 0.48, inner.w, 0.84),
       [("~300", {"size": 46, "bold": True, "color": "accent"}),
        ("  条", {"size": 16, "color": "ink_muted"})], anchor="middle")
s.text(Rect(inner.x, inner.y + 1.42, inner.w, 0.38),
       "集团巡检管理系统中 1~6 月关于 200L 桶操作的隐患记录",
       size=13, color="ink", line_spacing=1.35)
s.rect(Rect(inner.x, inner.y + 1.94, inner.w, 0.62), fill="accent", radius=0.03)
s.text(Rect(inner.x + 0.20, inner.y + 1.94, inner.w - 0.40, 0.62),
       "这些隐患未被从根本上整改解决", size=14.5, bold=True,
       color="FFFFFF", anchor="middle")

s.callout("两个 300 量级高度吻合：预警从不缺席，事故仍然发生 —— "
          "问题不在于没看见，而在于看见之后没有根治",
          at=bot, kind="accent", label="核心判断")
footer(s, 2, "法则数据为通用安全管理理论；隐患条数来自集团巡检管理系统 1~6 月记录")

# --------------------------------------------------------------------------
# 3 事故回顾
# --------------------------------------------------------------------------
s = deck.slide(title="本次事故基本情况", eyebrow="事故回顾",
               sub="本年度发生的一起二级事故，是本次专题倒查的起点")
top, bot = s.body.split_v(0.40, 0.60, gap=0.28)
cards = top.grid(3, 1, gap=0.26)
for cell, (label, value, hint) in zip(cards, [
        ("厂区", "【待填】厂区", "事故发生的厂区名称"),
        ("事故级别", "二级事故", "按集团事故分级标准"),
        ("发生位置", "3 号楼 1 层 303 房间", "具体作业地点")]):
    s.rect(cell, fill="surface")
    ii = cell.inset(0.26, 0.22)
    # 标签向下累加、注释固定在底部、主值填充中间：卡片高度变化时不会撞车
    s.text(Rect(ii.x, ii.y, ii.w, 0.26), label, size=12.5, bold=True,
           color="ink_muted")
    hint_top = ii.bottom - 0.26
    vs = 20 if len(value) <= 8 else 17
    s.text(Rect(ii.x, ii.y + 0.32, ii.w, hint_top - ii.y - 0.40), value,
           size=vs, bold=True, color="primary", line_spacing=1.2,
           anchor="middle", fit=True, min_pt=14)
    s.text(Rect(ii.x, hint_top, ii.w, 0.26), hint, size=11,
           color="ink_muted")

s.rect(bot, fill="surface")
ii = bot.inset(0.30, 0.26)
s.text(Rect(ii.x, ii.y, ii.w, 0.28), "事故简介", size=T.size_h, bold=True,
       color="primary")
s.text(Rect(ii.x, ii.y + 0.44, ii.w, 0.62),
       "【待填】请补充事故发生的时间、作业内容、直接原因、"
       "人员伤害情况与财产损失情况。", size=T.size_body, color="ink",
       line_spacing=1.45)
s.line(ii.x, ii.y + 1.22, ii.right, ii.y + 1.22, "hairline", 0.8)
s.bullets([
    ("这起事故的意义不止于事故本身",
     "它暴露出同类操作风险长期存在、此前已被多次记录却未被消除，"
     "因此本次同步开展了隐患倒查，结果见下一页"),
], at=Rect(ii.x, ii.y + 1.40, ii.w, ii.bottom - ii.y - 1.40))
footer(s, 3, "事故信息以正式事故调查报告为准；标注【待填】处需按实际情况补充")

# --------------------------------------------------------------------------
# 4 海因里希法则
# --------------------------------------------------------------------------
s = deck.slide(title="海因里希法则：事故是概率的必然，不是运气的偶然",
               eyebrow="理论依据")
left, right = s.body.split_h(0.56, 0.44, gap=0.36)

# 金字塔：上窄下宽，从上到下 1 / 29 / 300。
# 梯形内只放数字，名称和说明放到右侧统一起点 —— 塞进最窄那层会挤成多行。
pyr, note = left.split_h(0.48, 0.52, gap=0.22)
py = pyr.y + 0.16
tiers = [("1", "严重伤害事故", "死亡或重伤", "accent", 0.40),
         ("29", "轻微伤害事故", "需医疗处理但未致残", "warn", 0.70),
         ("300", "未遂事件与事故隐患", "尚未造成伤害的信号", "primary", 1.0)]
th = (pyr.h - 0.55) / 3
for val, name, desc, col, wr in tiers:
    bw = pyr.w * wr
    bx = pyr.x + (pyr.w - bw) / 2
    s.rect(Rect(bx, py, bw, th - 0.16), fill=col,
           shape=MSO_SHAPE.TRAPEZOID if wr < 1.0 else None, radius=0.02)
    s.text(Rect(bx, py, bw, th - 0.16), val, size=30, bold=True,
           color="FFFFFF", align="center", anchor="middle")
    # 右侧标注与该层垂直居中对齐
    cy = py + (th - 0.16) / 2
    s.text(Rect(note.x, cy - 0.32, note.w, 0.30), name, size=14.5, bold=True,
           color="primary", anchor="bottom")
    s.text(Rect(note.x, cy + 0.02, note.w, 0.30), desc, size=12,
           color="ink_muted", anchor="top")
    py += th
s.text(Rect(pyr.x, py - 0.06, pyr.w + note.w, 0.28),
       "比例关系为统计规律，不同行业数值略有差异",
       size=10.5, color="ink_muted")

s.text(Rect(right.x, right.y, right.w, 0.30), "这条法则对我们意味着什么",
       size=T.size_h, bold=True, color="primary")
s.bullets([
    ("严重事故从不孤立出现",
     "它站在数百个未被处理的隐患之上，是长期积累的结果"),
    ("隐患的数量决定事故的概率",
     "隐患基数不减少，事故只是时间问题，与运气无关"),
    ("能被消除的只有底层的 300",
     "无法直接消除事故，只能通过消除隐患降低概率"),
    ("已记录但未整改，等于没有发现",
     "记录本身不降低风险，只有闭环整改才降低风险"),
], at=Rect(right.x, right.y + 0.44, right.w, right.h - 0.44), gap=0.20)
footer(s, 4, "海因里希法则为通用安全管理理论，用于说明隐患数量与事故概率的统计关系")

# --------------------------------------------------------------------------
# 5 倒查结果
# --------------------------------------------------------------------------
s = deck.slide(title="巡检系统里躺着约 300 条同类隐患，半年未被根治",
               eyebrow="隐患倒查")
top, bot = s.body.split_v(0.40, 0.60, gap=0.28)
s.kpi_row([
    {"value": "~300", "unit": "条", "label": "200L 桶操作相关隐患记录",
     "color": "accent", "note": "集团巡检管理系统，1~6 月累计"},
    {"value": "6", "unit": "个月", "label": "隐患反复出现的时间跨度",
     "color": "primary", "note": "1 月至 6 月持续被记录"},
    {"value": "未", "unit": "根治", "label": "隐患的实际处理状态",
     "color": "accent", "note": "隐患未被从根本上整改解决"},
], at=top)

left, right = bot.split_h(0.52, 0.48, gap=0.34)
s.text(Rect(left.x, left.y, left.w, 0.30), "倒查是怎么做的",
       size=T.size_h, bold=True, color="primary")
s.bullets([
    ("以本次事故为起点，回溯半年隐患台账",
     "在集团巡检管理系统中检索同类操作风险的历史记录"),
    ("锁定 200L 桶操作这一类风险",
     "1~6 月共检索到约 300 条相关隐患记录"),
    ("核查这些隐患的处理状态",
     "结论是：隐患未被从根本上整改解决"),
], at=Rect(left.x, left.y + 0.44, left.w, left.h - 0.44), gap=0.18)

s.rect(right, fill="surface")
ii = right.inset(0.28, 0.24)
s.text(Rect(ii.x, ii.y, ii.w, 0.30), "这个结果说明什么",
       size=T.size_h, bold=True, color="primary")
s.bullets([
    "风险早已被识别，不存在没有想到的问题",
    "同类隐患在半年内被反复记录，说明整改未触及根源",
    "巡检发挥了发现作用，但闭环没有走完",
    "同样的风险仍在产线上，随时可能再次转化为事故",
], at=Rect(ii.x, ii.y + 0.46, ii.w, ii.bottom - ii.y - 0.46),
    size=13, bold_head=False, gap=0.16)
footer(s, 5, "数据来源：集团巡检管理系统隐患台账，统计区间为本年度 1 月至 6 月")

# --------------------------------------------------------------------------
# 6 问题本质：闭环断点
# --------------------------------------------------------------------------
s = deck.slide(title="隐患管理的闭环，断在「整改」与「验证」之间",
               eyebrow="根因分析",
               sub="以下为基于倒查结果的分析判断，具体断点需结合各厂区实际核实")
top, bot = s.body.split_v(0.40, 0.60, gap=0.28)

steps = [("01", "发现", "巡检、观察、上报", False),
         ("02", "记录", "录入巡检管理系统", False),
         ("03", "整改", "消除隐患的根本原因", True),
         ("04", "验证", "确认风险不再重现", True),
         ("05", "固化", "纳入标准与考核", False)]
cells = top.grid(5, 1, gap=0.16)
for cell, (no, name, desc, broken) in zip(cells, steps):
    s.rect(cell, fill="accent" if broken else "surface")
    ii = cell.inset(0.20, 0.18)
    fg = "FFFFFF" if broken else "ink"
    s.text(Rect(ii.x, ii.y, ii.w, 0.26), no, size=12.5, bold=True,
           color="FFFFFF" if broken else "ink_muted")
    s.text(Rect(ii.x, ii.y + 0.32, ii.w, 0.34), name, size=18, bold=True,
           color=fg)
    s.text(Rect(ii.x, ii.y + 0.72, ii.w, 0.52), desc, size=11.5,
           color=fg if broken else "ink_muted", line_spacing=1.35)
    if broken:
        s.text(Rect(ii.x, ii.bottom - 0.26, ii.w, 0.26), "断点",
               size=11, bold=True, color="FFFFFF", align="right")

left, right = bot.split_h(1, 1, gap=0.34)
s.text(Rect(left.x, left.y, left.w, 0.30), "断在哪里",
       size=T.size_h, bold=True, color="primary")
s.bullets([
    ("整改停在表面处置",
     "清理了现场、补了标识，但导致隐患反复出现的作业方式没有改变"),
    ("验证环节缺失或流于形式",
     "隐患关闭以已处理为准，而不是以同类风险不再出现为准"),
    ("重复出现未触发升级",
     "同类隐患被记录数百次，仍按单条处理，未升级为专项整治"),
], at=Rect(left.x, left.y + 0.44, left.w, left.h - 0.44), gap=0.18)

s.rect(right, fill="surface")
ii = right.inset(0.28, 0.24)
s.text(Rect(ii.x, ii.y, ii.w, 0.30), "闭环应当怎样才算走完",
       size=T.size_h, bold=True, color="primary")
s.bullets([
    ("整改要指向原因，不是指向现象",
     "改的是作业方式、工装或流程，不只是清理现场"),
    ("关闭要以不再重现为标准",
     "设定观察期，期内同类隐患零复现才允许关闭"),
    ("重复隐患要自动升级",
     "同类隐患累计到阈值即触发专项整治，不再按单条处理"),
], at=Rect(ii.x, ii.y + 0.46, ii.w, ii.bottom - ii.y - 0.46),
    size=13, gap=0.18)
footer(s, 6, "本页为基于隐患倒查结果的分析建议，具体措施需结合各厂区实际情况确定")

# --------------------------------------------------------------------------
# 7 五大保障基石
# --------------------------------------------------------------------------
s = deck.slide(title="安全生产的五大保障基石，目前均未被充分保障",
               eyebrow="保障基石",
               sub="五大体系共同支撑安全生产，任何一块不牢，事故就会从那里发生")
top, bot = s.body.split_v(0.70, 0.30, gap=0.26)
stones = [
    ("people", "人员培训及\n安全教育体系", "人是否真的会做、愿意做"),
    ("alert", "应急管理体系", "出事时能否有效控制损失"),
    ("flask", "危险化学品\n管理体系", "危化品全流程是否受控"),
    ("search", "安全风险管控及\n隐患排查治理", "风险是否被识别并消除"),
    ("equipment", "设备管理体系", "设备是否处于安全状态"),
]
cells = top.grid(5, 1, gap=0.20)
for cell, (ico, name, question) in zip(cells, stones):
    s.rect(cell, fill="surface")
    ii = cell.inset(0.22, 0.24)
    # 图标与名称向下累加，状态带固定在底部，设问填充中间剩余空间
    p = icon(ico)
    iy = ii.y
    if p:
        s.image(p, at=Rect(ii.x + (ii.w - 0.56) / 2, iy, 0.56, 0.56), mode="fit")
    else:
        s.rect(Rect(ii.x + (ii.w - 0.50) / 2, iy, 0.50, 0.50), fill="primary",
               shape=MSO_SHAPE.OVAL, radius=0)
    iy += 0.70
    nh = text_height(name, ii.w, 13.5, 1.28)
    s.text(Rect(ii.x, iy, ii.w, nh), name, size=13.5, bold=True,
           color="primary", align="center", line_spacing=1.28)
    iy += nh + 0.10
    band_top = ii.bottom - 0.34
    s.text(Rect(ii.x, iy, ii.w, max(band_top - iy - 0.10, 0.24)), question,
           size=11.5, color="ink_muted", align="center", line_spacing=1.32,
           anchor="top", fit=True, min_pt=10)
    # 状态带：汇报时按各体系实际评估结果填写
    s.rect(Rect(ii.x, band_top, ii.w, 0.34), fill="surface_alt", radius=0.06)
    s.text(Rect(ii.x, band_top, ii.w, 0.34), "现状：【待评估】",
           size=11, color="ink_muted", align="center", anchor="middle")

# 下方只留一条通栏结论：五块卡片已经把信息铺开，再分两块会挤且冲淡重点
s.callout("同类隐患半年内被记录约 300 次仍未根治 —— 风险管控与隐患治理这块基石"
          "首先没有立住。五个体系不是文件，要落到每一次实际操作上。",
          at=bot, kind="accent", label="请重视")
footer(s, 7, "五大保障基石划分沿用集团安全管理体系口径；各体系现状评级需按实际填写")

# --------------------------------------------------------------------------
# 8 行动要求
# --------------------------------------------------------------------------
s = deck.slide(title="请各层级立即落实的三件事", eyebrow="行动要求",
               sub="下列为建议框架，具体责任人与完成时限请按实际管理要求确定")
rowsp = s.body.split_v(1, 1, 1, gap=0.24)
actions = [
    ("管理层", "primary", [
        ("把重复隐患纳入管理视野",
         "对同类隐患累计超过阈值的风险点，升级为专项整治并跟踪到关闭"),
        ("为整改配资源",
         "涉及工装、设备、作业方式变更的整改，明确预算与责任部门"),
    ]),
    ("车间与班组", "secondary", [
        ("隐患关闭改为复现验证制",
         "设定观察期，期内同类隐患零复现才允许关闭，不以已处理为准"),
        ("对 200L 桶操作开展专项排查",
         "覆盖全部相关岗位，逐条核对作业方式是否已实质改变"),
    ]),
    ("每一位员工", "good", [
        ("继续上报，不要因为报了没用而停止",
         "上报量下降不等于更安全，隐瞒才是最大的风险"),
        ("按标准作业，不省步骤",
         "绝大多数事故的直接原因，是一次以为这次没事的省略"),
    ]),
]
for row, (who, col, items) in zip(rowsp, actions):
    s.rect(Rect(row.x, row.y, 1.72, row.h), fill=col, radius=0.04)
    s.text(Rect(row.x + 0.16, row.y, 1.40, row.h), who, size=15.5, bold=True,
           color="FFFFFF", align="center", anchor="middle")
    bx = row.x + 1.90
    s.bullets(items, at=Rect(bx, row.y + 0.08, row.right - bx, row.h - 0.16),
              size=13.5, gap=0.14)
footer(s, 8, "本页为建议框架，请按各厂区实际情况确定责任人、完成时限与验证方式")

# --------------------------------------------------------------------------
# 9 结尾
# --------------------------------------------------------------------------
s = deck.raw_slide(layout=LAYOUT_DECOR)
body = s.safe
s.text(Rect(body.x, body.y + 1.05, body.w, 0.34), "请全员重视",
       size=13, bold=True, color="accent")
s.rect(Rect(body.x, body.y + 1.53, 0.075, 1.52), fill="accent", radius=0)
s.text(Rect(body.x + 0.30, body.y + 1.49, body.w * 0.88, 0.68),
       "下一个 300 条隐患里，可能就藏着下一起事故",
       size=31, bold=True, color="primary", line_spacing=1.18)
s.text(Rect(body.x + 0.30, body.y + 2.31, body.w * 0.72, 0.60),
       "我们无法消除事故本身，但可以消除通向事故的每一条隐患。",
       size=16, color="ink_muted", line_spacing=1.4)
row = Rect(body.x + 0.30, 5.20, body.w - 0.30, 0.86)
for cell, (k, v) in zip(row.grid(3, 1, gap=0.26), [
        ("发现隐患", "立即上报，不判断大小"),
        ("整改隐患", "改到原因，不停在现象"),
        ("关闭隐患", "以不再重现为标准")]):
    s.rect(cell, fill="surface")
    ii = cell.inset(0.22, 0.16)
    s.text(Rect(ii.x, ii.y, ii.w, 0.30), k, size=14.5, bold=True,
           color="primary")
    s.text(Rect(ii.x, ii.y + 0.36, ii.w, 0.30), v, size=12, color="ink_muted")
s.line(body.x, 6.38, body.right, 6.38, "hairline", 0.8)
s.text(Rect(body.x, 6.50, body.w, 0.28),
       "本材料为内部安全警示用途，请勿外传", size=11, color="ink_muted")

deck.save(OUT)
print("已生成：%s（%d 页，画布 %.2f x %.2f 英寸）"
      % (OUT, len(deck.slides), deck.canvas_w, deck.canvas_h))
check_overflow(OUT)
