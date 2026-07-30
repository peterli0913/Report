"""示例：贴合企业现有模板的配色与画布（数据为虚构演示数据）。

演示三件事：
  1. 用从现有模板提取的品牌色派生主题，并用 with_readable_text() 修正文字对比度
  2. 用 26.67x15 大画布（与现有模板同尺寸），组件代码与标准画布完全一致
  3. 给封面加程序生成的背景图案，并用半透明遮罩保证标题可读

运行：
    python3 example_brand.py
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))
sys.path.insert(0, os.path.join(HERE, "..", "..", "deck-imagery", "scripts"))

from deckkit import DARK, Deck, Rect, check_overflow  # noqa: E402

OUT = os.path.join(HERE, "example_brand.pptx")
BG = os.path.join(HERE, "cover-bg.png")

# 色值取自 安全专篇1页.pptx 的 ppt/theme/theme2.xml，是现有模板的实际用色。
# 正式对外材料请向品牌部门确认官方 VI 色值。
BRAND = DARK.variant(
    name="asymchem-brand",
    bg="003669",           # dk1 主色深蓝
    bg_alt="002A52",
    surface="0B4880",
    surface_alt="3263A7",  # dk2
    hairline="4A79B8",
    ink="E7E6E6",          # lt2，模板正文实际用色
    ink_muted="A8C0DC",
    primary="4472C4",      # accent1
    secondary="2AA9DB",    # accent6
    accent="E5B620",       # accent2 强调金
    series=("4472C4", "E5B620", "2AA9DB", "27A567", "2649AB", "C93B33"),
).with_readable_text()     # 换了底色，语义文字色必须重算

# 封面背景：分子网络图案，中心留空放标题。缺 Pillow 时跳过，不影响其余页面。
bg_path = None
try:
    from make_bg import molecule
    molecule(BG, "002A52", "2AA9DB", size=(2400, 1350), nodes=24)
    bg_path = BG
except Exception as e:
    print("跳过背景图（%s）" % e)

deck = Deck(theme=BRAND, canvas="large")   # 26.67 x 15 英寸
T = deck.theme

deck.cover(title="2026 年上半年生产运营回顾",
           subtitle="产能达成、质量合规与产线瓶颈的系统性复盘",
           kicker="生产管理部 · 半年度经营汇报",
           meta="汇报人：张三    |    2026 年 7 月 30 日    |    内部资料 请勿外传",
           image=bg_path, scrim=52)

s = deck.slide(title="上半年产量超目标 5.2 个百分点，但 3 号线瓶颈已开始制约三季度交付",
               eyebrow="执行摘要")
top, bot = s.body.split_v(0.40, 0.60, gap=0.26)
s.kpi_row([
    {"value": "1,284", "unit": "吨", "label": "上半年总产量",
     "delta": "同比 +23.1%", "trend": "up"},
    {"value": "105.2", "unit": "%", "label": "年度目标达成率",
     "delta": "高于计划 5.2pp", "trend": "up", "color": "accent"},
    {"value": "99.2", "unit": "%", "label": "批次一次通过率",
     "delta": "同比 +1.4pp", "trend": "up"},
    # 偏差数下降是好事，必须显式指定绿色，不能让工具按箭头方向自动配色
    {"value": "6", "unit": "件", "label": "重大偏差",
     "delta": "同比 -45%", "trend": "down", "delta_color": "good"},
], at=top)

la, ra = bot.split_h(1, 1, gap=0.34)
s.text(Rect(la.x, la.y, la.w, 0.28), "三个结论", size=T.size_h, bold=True)
s.bullets([("产能已释放到设计上限", "1、2 号线利用率 94%，进一步增产依赖 3 号线扩能"),
           ("质量指标全面向好", "一次通过率 99.2%，重大偏差同比下降 45%"),
           ("瓶颈从原料转向后处理", "结晶工序 CT 占全流程 38%，是当前唯一瓶颈")],
          at=Rect(la.x, la.y + 0.38, la.w, la.h - 0.38))
s.text(Rect(ra.x, ra.y, ra.w, 0.28), "需要决策的事项", size=T.size_h, bold=True)
s.bullets(["3 号线结晶工序扩能投资 2,800 万元，请批准立项",
           "Q3 起 A 产品排产上限暂定 240 吨 / 月",
           "关键岗位增编 12 人，请人力资源部支持"],
          at=Rect(ra.x, ra.y + 0.38, ra.w, ra.h - 0.38), marker="num")
s.note("数据来源：MES 系统 + ERP 生产模块，统计区间 2026-01-01 至 2026-06-30")

s = deck.slide(title="产量逐月攀升，Q2 月均产量较 Q1 提升 28%", eyebrow="产量趋势")
ca, cb = s.body.split_v(0.78, 0.22, gap=0.18)
s.chart("bar", ["1月", "2月", "3月", "4月", "5月", "6月"],
        [("2025 年", [148, 132, 165, 171, 178, 182]),
         ("2026 年", [186, 171, 208, 232, 241, 246])], at=ca, number_format="0")
s.callout("6 月单月产量 246 吨创历史新高，已达 3 号线满负荷状态", at=cb,
          kind="accent", label="关键发现")
s.note("数据来源：MES 系统批次产出记录")

deck.add_page_numbers()
deck.save(OUT)
print("已生成：%s（%d 页，画布 %.2f x %.2f 英寸）"
      % (OUT, len(deck.slides), deck.canvas_w, deck.canvas_h))
check_overflow(OUT)
