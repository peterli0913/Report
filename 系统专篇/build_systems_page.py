"""两页：设备信息管理系统（EIMS）与巡检管理系统，各占一页。

第 1 页：EIMS —— 文字内容原样取自 设备管理系统(1).pptx（里程碑恢复为横向燕尾形）。
第 2 页：巡检管理系统 —— 左截图 + 右系统能力说明。

沿用 设备管理系统(1).pptx 的模板：2_自定义版式。
安全内容区按其背景图 image6.jpg 实测：页眉占到 y 1.06 英寸，底部条从 y 7.18 起。

运行：python3 build_systems_page.py [输出路径.pptx]
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(REPO, ".cursor", "skills", "exec-deck-builder",
                               "scripts"))

from deckkit import LIGHT, Deck, Rect, check_overflow, text_height  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402

TEMPLATE = os.path.join(REPO, "设备管理系统(1).pptx")
SHOT = os.path.join(REPO, "巡检管理系统.png")
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "两系统建设进展.pptx")

T = LIGHT.variant(
    name="asymchem-systems",
    bg="FFFFFF", bg_alt="F7FAFD",
    surface="EEF3FA", surface_alt="DDE7F3",
    hairline="C6D5E6",
    ink="16283F", ink_muted="4E6584", ink_on_accent="FFFFFF",
    primary="003669", secondary="3263A7", accent="C8880C",
    good="1A7F50", warn="A34E0F", bad="C62828", neutral="5E7288",
    accent_text="8A5D06", good_text="167348", warn_text="A04C0D",
    bad_text="C62828", neutral_text="4E6584", primary_text="003669",
    margin=0.56, margin_top=1.20, margin_bottom=0.38,
    size_title=22, size_h=14, size_body=12, size_small=11, size_note=9,
)

deck = Deck(theme=T, template=TEMPLATE, layout="2_自定义版式")


def block_label(s, area, y, text):
    s.rect(Rect(area.x, y + 0.02, 0.055, 0.22), fill="accent", radius=0)
    s.text(Rect(area.x + 0.16, y, area.w - 0.16, 0.26), text,
           size=T.size_h, bold=True, color="primary")
    return y + 0.32


def footer(s, n, note):
    s.text(Rect(T.margin, s._foot_y, 10.8, 0.20), note,
           size=T.size_note, color="ink_muted")
    s.text(Rect(11.50, s._foot_y, 0.80, 0.20), "%d / 2" % n,
           size=T.size_note, color="ink_muted", align="right")


# ==========================================================================
# 第 1 页：设备信息管理系统（EIMS）—— 文字内容原样取自源 PPT
# ==========================================================================
s = deck.slide(
    title="设备信息管理系统（EIMS）完成系统验证，预计 8 月上线")
body = s.body
y = body.y

s.text(Rect(body.x, y, body.w, 0.22),
       "Equipment Information Management System", size=11, bold=True,
       color="ink_muted")
y += 0.30

# ---- 上半年工作回顾 ----
y = block_label(s, body, y, "2026 年上半年工作回顾")
cards = Rect(body.x, y, body.w, 1.00).split_h(1, 1, gap=0.20)
for cell, (head, body_txt) in zip(cards, [
    ("设备信息管理系统 (EIMS）完成验证 - 设备预防性维护",
     "按照年度计划在 DH1 工厂完成 EIMS 系统验证，具备上线运行使用条件，"
     "预计 8 月份正式上线使用（其他工厂提交变更启用）。"),
    ("设备故障报修模块覆盖使用",
     "设备故障报修模块各工厂已启用，自 1 月份启动以来，累计报修 4,804 次，"
     "已闭环处理 2,539 单（TJ3/DH1 报修量最高）。"),
]):
    s.rect(cell, fill="surface")
    ii = cell.inset(0.20, 0.12)
    hh = text_height(head, ii.w, 12.5, 1.20)
    s.text(Rect(ii.x, ii.y, ii.w, hh), head, size=12.5, bold=True,
           color="primary", line_spacing=1.20)
    s.text(Rect(ii.x, ii.y + hh + 0.06, ii.w, ii.bottom - ii.y - hh - 0.06),
           body_txt, size=11, color="ink", line_spacing=1.30)
y += 1.12

# ---- 系统实施里程碑（恢复横向，与原 PPT 一致）----
y = block_label(s, body, y, "系统实施里程碑")
MILESTONES = [
    ("2 月", "系统功能开发",
     "1、EIMS 系统功能开发；\n2、完成变更控制文件与 URS 签批；"),
    ("3 月", "UAT 用户测试",
     "1、EIMS 系统环境测试；\n2、UAT 及 VP、RAR、FS 签批，UAT 用户测试；"),
    ("4 月", "文件撰写审批", "1、CS+DS+代码审核和 UTITPR；"),
    ("5 月", "系统验证执行", "1、3Q 验证执行和验证报告签批；"),
    ("6 月", "系统放行与培训",
     "1、EIMS 系统 SOP 生效培训；\n2、EIMS 系统验证后放行；"),
]
cells = Rect(body.x, y, body.w, 1.42).grid(5, 1, gap=0.12)
for cell, (mon, name, detail) in zip(cells, MILESTONES):
    s.rect(Rect(cell.x, cell.y, cell.w, 0.36), fill="surface_alt",
           shape=MSO_SHAPE.CHEVRON, radius=0)
    s.text(Rect(cell.x + 0.08, cell.y, cell.w - 0.20, 0.36),
           "%s - %s" % (mon, name), size=10.5, bold=True, color="ink",
           align="center", anchor="middle")
    s.text(Rect(cell.x + 0.06, cell.y + 0.44, cell.w - 0.12, cell.h - 0.48),
           detail, size=10, color="ink_muted", line_spacing=1.30,
           fit=True, min_pt=9)
y += 1.52

# ---- 下半年工作计划（三列，避免纵向挤爆）----
y = block_label(s, body, y, "2026 年下半年工作计划")
PLANS = [
    ("设备全生命周期管理",
     "基于设备全生命周期管理 (ASM-037)：内控平台功能模块开发（备品备件管理），"
     "进一步完善设备全生命周期管理"),
    ("设备故障看板与报表",
     "优化设备故障分析仪表盘原型设计，实现设备故障分析报表和数据可视化。"),
    ("内控平台数据互通",
     "内控平台功能升级、历史数据迁移及验证环境数据共享，实现与 EIMS 的数据互通，"
     "预计 8 月份完成。"),
]
avail = body.bottom - y
cells = Rect(body.x, y, body.w, avail).split_h(1, 1, 1, gap=0.16)
for cell, (head, body_txt) in zip(cells, PLANS):
    s.rect(cell, fill="surface")
    ii = cell.inset(0.16, 0.12)
    s.text(Rect(ii.x, ii.y, ii.w, 0.42), head, size=12.5, bold=True,
           color="primary", line_spacing=1.20)
    s.text(Rect(ii.x, ii.y + 0.46, ii.w, ii.h - 0.46), body_txt,
           size=11, color="ink", line_spacing=1.30, fit=True, min_pt=9.5)

footer(s, 1, "数据来自设备管理系统汇报材料")


# ==========================================================================
# 第 2 页：巡检管理系统 —— 左截图 / 右能力
# ==========================================================================
s = deck.slide(title="巡检管理系统：隐患台账已归集 3.5 万余条")
body = s.body
y = body.y

s.text(Rect(body.x, y, body.w, 0.22),
       "Inspection Management System · EHS 隐患排查", size=11, bold=True,
       color="ink_muted")
y += 0.30

left, right = Rect(body.x, y, body.w, body.bottom - y).split_h(58, 42, gap=0.28)

# 左：系统界面截图
s.rect(Rect(left.x, left.y + 0.02, 0.055, 0.22), fill="accent", radius=0)
s.text(Rect(left.x + 0.16, left.y, left.w - 0.16, 0.26), "系统界面与隐患台账",
       size=T.size_h, bold=True, color="primary")
shot_top = left.y + 0.34
try:
    from PIL import Image
    with Image.open(SHOT) as im:
        ar = im.width / im.height
except Exception:
    ar = 1024 / 442

shot_w = left.w
shot_h = min(shot_w / ar, left.bottom - shot_top - 0.36)
frame = Rect(left.x, shot_top, shot_w, shot_h + 0.14)
s.rect(frame, fill="surface_alt", line="hairline", line_w=0.75, radius=0.02)
s.image(SHOT, at=Rect(left.x + 0.07, shot_top + 0.07, shot_w - 0.14, shot_h),
        mode="fit")
s.text(Rect(left.x, frame.bottom + 0.06, left.w, 0.28),
       "EHS 隐患排查台账（查询区间 2026-01-01 至 2026-07-30；截图显示累计 35,684 条）",
       size=10, color="ink_muted", line_spacing=1.25)

# 右：系统能力列表
s.rect(Rect(right.x, right.y + 0.02, 0.055, 0.22), fill="accent", radius=0)
s.text(Rect(right.x + 0.16, right.y, right.w - 0.16, 0.26), "系统能力",
       size=T.size_h, bold=True, color="primary")
FEATURES = [
    ("隐患全过程留痕",
     "逐条记录巡查主题、组织部门、巡查人、隐患地点、巡查发现、隐患分类与隐患等级"),
    ("三类巡查统一入口",
     "EHS 隐患排查、QA 日常巡查、生产日常巡查在同一系统内管理"),
    ("多维度统计分析",
     "支持按厂区、部门、巡查人、整改责任人、第一责任人分别统计"),
    ("计划与任务跟踪",
     "巡检计划与巡检任务明细可查，台账截图显示已累计 35,684 条记录"),
]
feat_area = Rect(right.x, right.y + 0.36, right.w, right.bottom - right.y - 0.36)
cells = feat_area.split_v(*([1] * len(FEATURES)), gap=0.12)
for cell, (head, body_txt) in zip(cells, FEATURES):
    s.rect(cell, fill="surface")
    ii = cell.inset(0.16, 0.10)
    s.rect(Rect(ii.x, ii.y + 0.04, 0.05, 0.20), fill="secondary", radius=0)
    s.text(Rect(ii.x + 0.14, ii.y, ii.w - 0.14, 0.28), head,
           size=13, bold=True, color="primary")
    s.text(Rect(ii.x, ii.y + 0.32, ii.w, ii.h - 0.32), body_txt,
           size=11, color="ink", line_spacing=1.30, fit=True, min_pt=9.5)

footer(s, 2, "信息取自巡检管理系统界面截图")

deck.save(OUT)
print("已生成：%s（%d 页，%.2f x %.2f 英寸）"
      % (OUT, len(deck.slides), deck.canvas_w, deck.canvas_h))
check_overflow(OUT)
