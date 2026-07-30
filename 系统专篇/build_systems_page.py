"""一页：设备信息管理系统（EIMS）与巡检管理系统。

左侧内容原样取自 设备管理系统(1).pptx（半页宽度下里程碑改竖向）；
右侧为 巡检管理系统.png 大图展示（无总标题、无「系统能力」列表）。

沿用设备管理系统(1).pptx 的模板：2_自定义版式。
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
# 不设页级总标题，左右各自用系统名作栏目标题，腾出垂直空间放大截图
s = deck.slide()

# 右侧略宽，便于放大巡检截图
LEFT, RIGHT = s.body.split_h(46, 54, gap=0.28)


def block_label(area, y, text, w=None):
    s.rect(Rect(area.x, y + 0.02, 0.055, 0.20), fill="accent", radius=0)
    s.text(Rect(area.x + 0.16, y, (w or area.w) - 0.16, 0.24), text,
           size=T.size_h, bold=True, color="primary")
    return y + 0.28


# ==========================================================================
# 左侧：设备信息管理系统（EIMS）—— 文字内容原样取自源 PPT
# ==========================================================================
y = LEFT.y
s.text(Rect(LEFT.x, y, LEFT.w, 0.26), "设备信息管理系统（EIMS）",
       size=16, bold=True, color="primary")
s.text(Rect(LEFT.x, y + 0.28, LEFT.w, 0.20),
       "Equipment Information Management System", size=9.5, bold=True,
       color="ink_muted")
y += 0.52

y = block_label(LEFT, y, "2026 年上半年工作回顾")

cards = Rect(LEFT.x, y, LEFT.w, 0.92).split_h(1, 1, gap=0.14)
for cell, (head, body) in zip(cards, [
    ("EIMS 完成验证 · 设备预防性维护",
     "按照年度计划在 DH1 工厂完成 EIMS 系统验证，具备上线运行使用条件，"
     "预计 8 月份正式上线使用（其他工厂提交变更启用）。"),
    ("设备故障报修模块覆盖使用",
     "设备故障报修模块各工厂已启用，自 1 月份启动以来，累计报修 4,804 次，"
     "已闭环处理 2,539 单（TJ3/DH1 报修量最高）。"),
]):
    s.rect(cell, fill="surface")
    ii = cell.inset(0.14, 0.10)
    hh = text_height(head, ii.w, 10.5, 1.22)
    s.text(Rect(ii.x, ii.y, ii.w, hh), head, size=10.5, bold=True,
           color="primary", line_spacing=1.22)
    s.text(Rect(ii.x, ii.y + hh + 0.04, ii.w, ii.bottom - ii.y - hh - 0.04),
           body, size=8.5, color="ink", line_spacing=1.30, fit=True, min_pt=7.5)
y += 1.02

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
axis_x = LEFT.x + 0.32
step = 0.36
s.line(axis_x, y + 0.10, axis_x, y + step * (len(MILESTONES) - 1) + 0.22,
       "hairline", 1.2)
for i, (mon, name, detail) in enumerate(MILESTONES):
    cy = y + step * i + 0.16
    s.rect(Rect(axis_x - 0.065, cy - 0.065, 0.13, 0.13), fill="primary",
           shape=None, radius=0.5)
    s.text(Rect(LEFT.x, cy - 0.11, 0.28, 0.22), mon, size=9.5, bold=True,
           color="accent", align="right", anchor="middle")
    tx = axis_x + 0.16
    s.text(Rect(tx, cy - 0.115, 1.20, 0.22), name, size=10.5, bold=True,
           color="primary", anchor="middle")
    s.text(Rect(tx + 1.24, cy - 0.14, LEFT.right - tx - 1.24, 0.28), detail,
           size=8.5, color="ink_muted", anchor="middle", fit=True, min_pt=7.5)
y += step * len(MILESTONES) + 0.14

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
# 右侧：巡检管理系统 —— 标题 + 大图（无「系统能力」）
# ==========================================================================
y = RIGHT.y
s.text(Rect(RIGHT.x, y, RIGHT.w, 0.26), "巡检管理系统",
       size=16, bold=True, color="primary")
s.text(Rect(RIGHT.x, y + 0.28, RIGHT.w, 0.20),
       "Inspection Management System · EHS 隐患排查", size=9.5, bold=True,
       color="ink_muted")
y += 0.54

y = block_label(RIGHT, y, "系统界面与隐患台账")

try:
    from PIL import Image
    with Image.open(SHOT) as im:
        ar = im.width / im.height
except Exception:
    ar = 1024 / 442

# 下方仅留一行图注，其余高度全部给截图
CAPTION_H = 0.30
avail_h = RIGHT.bottom - y - CAPTION_H - 0.08
shot_w = RIGHT.w
shot_h = min(shot_w / ar, avail_h)
# 若按宽算出的高度仍有余量，再按高度反推加宽感——宽度已是满栏，直接居中贴满
frame = Rect(RIGHT.x, y, shot_w, shot_h + 0.14)
s.rect(frame, fill="surface_alt", line="hairline", line_w=0.75, radius=0.02)
s.image(SHOT, at=Rect(RIGHT.x + 0.08, y + 0.07, shot_w - 0.16, shot_h),
        mode="fit")
y = frame.bottom + 0.08
s.text(Rect(RIGHT.x, y, RIGHT.w, 0.26),
       "EHS 隐患排查台账（查询区间 2026-01-01 至 2026-07-30；截图显示累计 35,684 条）",
       size=9, color="ink_muted")

s.text(Rect(T.margin, s._foot_y, 12.21, 0.20),
       "设备系统数据来自设备管理系统汇报材料；巡检系统信息取自系统界面截图",
       size=8.5, color="ink_muted")

deck.save(OUT)
print("已生成：%s（%d 页，%.2f x %.2f 英寸）"
      % (OUT, len(deck.slides), deck.canvas_w, deck.canvas_h))
check_overflow(OUT)
