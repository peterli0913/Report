"""一页：设备信息管理系统（EIMS）与巡检管理系统。

左侧内容原样取自 设备管理系统(1).pptx 的那一页（文字未改动，仅因半页宽度把原来
横向的 5 段里程碑改成竖向排列）；右侧是 巡检管理系统.png 界面截图与要点。

沿用设备管理系统(1).pptx 的模板：2_自定义版式（页眉 ASYMCHEM 标 + 底部深蓝条）。
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
    size_title=20, size_h=13, size_body=11, size_small=9.5, size_note=8,
)

deck = Deck(theme=T, template=TEMPLATE, layout="2_自定义版式")
s = deck.slide(
    title="EIMS 完成系统验证具备上线条件，巡检系统隐患台账已归集 3.5 万余条")

LEFT, RIGHT = s.body.split_h(1, 1, gap=0.32)


def block_label(area, y, text, w=None):
    """小节标签：左侧一小段强调色竖条 + 深蓝粗体标题。"""
    s.rect(Rect(area.x, y + 0.02, 0.055, 0.20), fill="accent", radius=0)
    s.text(Rect(area.x + 0.16, y, (w or area.w) - 0.16, 0.24), text,
           size=T.size_h, bold=True, color="primary")
    return y + 0.26


# ==========================================================================
# 左侧：设备信息管理系统（EIMS）—— 文字内容原样取自源 PPT
# ==========================================================================
y = LEFT.y
s.text(Rect(LEFT.x, y, LEFT.w, 0.26), "设备信息管理系统（EIMS）",
       size=15, bold=True, color="primary")
s.text(Rect(LEFT.x, y + 0.27, LEFT.w, 0.20),
       "Equipment Information Management System", size=9, bold=True,
       color="ink_muted")
y += 0.42

y = block_label(LEFT, y, "2026 年上半年工作回顾")

# 两个成果要点并排
cards = Rect(LEFT.x, y, LEFT.w, 0.78).split_h(1, 1, gap=0.16)
for cell, (head, body) in zip(cards, [
    ("EIMS 完成验证 · 设备预防性维护",
     "按照年度计划在 DH1 工厂完成 EIMS 系统验证，具备上线运行使用条件，"
     "预计 8 月份正式上线使用（其他工厂提交变更启用）。"),
    ("设备故障报修模块覆盖使用",
     "设备故障报修模块各工厂已启用，自 1 月份启动以来，累计报修 4,804 次，"
     "已闭环处理 2,539 单（TJ3/DH1 报修量最高）。"),
]):
    s.rect(cell, fill="surface")
    ii = cell.inset(0.16, 0.13)
    hh = text_height(head, ii.w, 10.5, 1.25)
    s.text(Rect(ii.x, ii.y, ii.w, hh), head, size=10.5, bold=True,
           color="primary", line_spacing=1.25)
    s.text(Rect(ii.x, ii.y + hh + 0.05, ii.w, ii.bottom - ii.y - hh - 0.05),
           body, size=8.5, color="ink", line_spacing=1.34, fit=True, min_pt=7.5)
y += 0.88

# 5 段实施里程碑：源 PPT 是横向燕尾形，半页宽放不下，改竖向时间轴，文字不变
y = block_label(LEFT, y, "系统实施里程碑")
MILESTONES = [
    ("2 月", "系统功能开发",
     "1、EIMS 系统功能开发；2、完成变更控制文件与 URS 签批；"),
    ("3 月", "UAT 用户测试",
     "1、EIMS 系统环境测试；2、UAT 及 VP、RAR、FS 签批，UAT 用户测试；"),
    ("4 月", "文件撰写审批", "1、CS+DS+代码审核和 UTITPR；"),
    ("5 月", "系统验证执行", "1、3Q 验证执行和验证报告签批；"),
    ("6 月", "系统放行与培训",
     "1、EIMS 系统 SOP 生效培训；2、EIMS 系统验证后放行；"),
]
axis_x = LEFT.x + 0.30
step = 0.335
s.line(axis_x, y + 0.10, axis_x, y + step * (len(MILESTONES) - 1) + 0.20,
       "hairline", 1.2)
for i, (mon, name, detail) in enumerate(MILESTONES):
    cy = y + step * i + 0.15
    s.rect(Rect(axis_x - 0.065, cy - 0.065, 0.13, 0.13), fill="primary",
           shape=None, radius=0.5)
    s.text(Rect(LEFT.x, cy - 0.11, 0.26, 0.22), mon, size=9, bold=True,
           color="accent", align="right", anchor="middle")
    tx = axis_x + 0.16
    s.text(Rect(tx, cy - 0.115, 1.16, 0.22), name, size=10, bold=True,
           color="primary", anchor="middle")
    s.text(Rect(tx + 1.20, cy - 0.14, LEFT.right - tx - 1.20, 0.28), detail,
           size=8.5, color="ink_muted", anchor="middle", fit=True, min_pt=7.5)
y += step * len(MILESTONES) + 0.12

# 下半年重点
y = block_label(LEFT, y, "2026 年下半年工作计划")
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
row_h = (LEFT.bottom - y) / len(PLANS)
for i, (head, body) in enumerate(PLANS):
    ry = y + row_h * i
    s.rect(Rect(LEFT.x, ry, LEFT.w, row_h - 0.05), fill="surface")
    ii = Rect(LEFT.x + 0.14, ry, LEFT.w - 0.28, row_h - 0.05)
    s.text(Rect(ii.x, ii.y, 1.55, ii.h), head, size=9.5, bold=True,
           color="primary", anchor="middle")
    s.text(Rect(ii.x + 1.55, ii.y, ii.right - ii.x - 1.55, ii.h), body,
           size=8.5, color="ink_muted", anchor="middle", line_spacing=1.24,
           fit=True, min_pt=8)

# ==========================================================================
# 右侧：巡检管理系统
# ==========================================================================
y = RIGHT.y
s.text(Rect(RIGHT.x, y, RIGHT.w, 0.26), "巡检管理系统",
       size=15, bold=True, color="primary")
s.text(Rect(RIGHT.x, y + 0.27, RIGHT.w, 0.20),
       "Inspection Management System · EHS 隐患排查", size=9, bold=True,
       color="ink_muted")
y += 0.48

y = block_label(RIGHT, y, "系统界面与隐患台账")

# 截图：外面套一层浅底卡片 + 细边框，避免白底截图和白底页面糊在一起
try:
    from PIL import Image
    with Image.open(SHOT) as im:
        ar = im.width / im.height
except Exception:
    ar = 1024 / 442
# 先给下方的说明与 4 条要点留出高度，剩下的才给截图，避免要点被压成一条线
NEED_BELOW = 0.26 + 0.26 + 4 * 0.30
shot_h = min(RIGHT.w / ar, RIGHT.bottom - y - NEED_BELOW - 0.24)
shot_w = shot_h * ar
frame = Rect(RIGHT.x + (RIGHT.w - shot_w) / 2, y, shot_w, shot_h + 0.14)
s.rect(frame, fill="surface_alt", line="hairline", line_w=0.75, radius=0.02)
s.image(SHOT, at=Rect(frame.x + 0.07, y + 0.07, shot_w - 0.14, shot_h),
        mode="fit")
y += frame.h + 0.10
s.text(Rect(RIGHT.x, y, RIGHT.w, 0.22),
       "系统界面：EHS 隐患排查台账（截图内查询区间 2026-01-01 至 2026-07-30）",
       size=8.5, color="ink_muted")
y += 0.30

y = block_label(RIGHT, y, "系统能力")
FEATURES = [
    ("隐患全过程留痕",
     "逐条记录巡查主题、部门、巡查人、地点、巡查发现、分类与等级"),
    ("三类巡查统一入口",
     "EHS 隐患排查、QA 日常巡查、生产日常巡查在同一系统内管理"),
    ("多维度统计分析",
     "支持按厂区、部门、巡查人、整改责任人、第一责任人分别统计"),
    ("计划与任务跟踪",
     "巡检计划与巡检任务明细可查，台账截图显示已累计 35,684 条记录"),
]
fh = (RIGHT.bottom - y) / len(FEATURES)
for i, (head, body) in enumerate(FEATURES):
    ry = y + fh * i
    s.rect(Rect(RIGHT.x, ry, 0.055, fh - 0.06), fill="secondary", radius=0)
    ii = Rect(RIGHT.x + 0.16, ry, RIGHT.w - 0.16, fh - 0.06)
    s.text(Rect(ii.x, ii.y, 1.45, ii.h), head, size=9.5, bold=True,
           color="primary", anchor="middle")
    s.text(Rect(ii.x + 1.52, ii.y, ii.right - ii.x - 1.52, ii.h), body,
           size=8.5, color="ink_muted", anchor="middle", line_spacing=1.26,
           fit=True, min_pt=8.5)

s.text(Rect(T.margin, 7.5 - T.mb - 0.20, 12.21, 0.20),
       "设备系统数据来自设备管理系统汇报材料；巡检系统信息取自系统界面截图",
       size=8.5, color="ink_muted")

deck.save(OUT)
print("已生成：%s（%d 页，%.2f x %.2f 英寸）"
      % (OUT, len(deck.slides), deck.canvas_w, deck.canvas_h))
check_overflow(OUT)
