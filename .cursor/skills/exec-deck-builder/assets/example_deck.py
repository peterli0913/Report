"""示例：一份完整的生产运营月度汇报（数据为虚构演示数据）。

这份脚本是 exec-deck-builder 的参照实现 —— 需要生成新汇报时，从这里挑
结构相近的页型改内容，比从零写坐标快得多，也不会漏掉 note/校验步骤。

运行：
    python3 example_deck.py                    # 生成到本脚本同目录
    python3 example_deck.py /tmp/out.pptx      # 生成到指定路径（不改动仓库文件）
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))

from deckkit import Deck, Rect, check_overflow  # noqa: E402

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "example_deck.pptx")

deck = Deck(theme="dark", canvas="wide")
T = deck.theme

AGENDA = ["整体表现", "质量与合规", "瓶颈与改善", "下一步计划"]

# ---------------------------------------------------------------- 封面
deck.cover(
    title="2026 年上半年生产运营回顾",
    subtitle="产能达成、质量合规与产线瓶颈的系统性复盘",
    kicker="生产管理部 · 半年度经营汇报",
    meta="汇报人：张三    |    2026 年 7 月 30 日    |    内部资料 请勿外传",
)

# ---------------------------------------------------------------- 摘要页
s = deck.slide(
    title="上半年产量超目标 5.2 个百分点，但 3 号线瓶颈已开始制约三季度交付",
    eyebrow="执行摘要",
)
top, bottom = s.body.split_v(0.44, 0.56, gap=0.26)
s.kpi_row([
    {"value": "1,284", "unit": "吨", "label": "上半年总产量",
     "delta": "同比 +23.1%", "trend": "up"},
    {"value": "105.2", "unit": "%", "label": "年度目标达成率",
     "delta": "高于计划 5.2pp", "trend": "up", "color": "accent"},
    {"value": "99.2", "unit": "%", "label": "批次一次通过率",
     "delta": "同比 +1.4pp", "trend": "up"},
    {"value": "6", "unit": "件", "label": "重大偏差",
     "delta": "同比 -45%", "trend": "down", "delta_color": "good"},
], at=top)

left, right = bottom.split_h(0.52, 0.48, gap=0.3)
s.text(Rect(left.x, left.y, left.w, 0.26), "三个结论", size=T.size_h, bold=True)
s.bullets([
    ("产能已释放到设计上限", "1、2 号线利用率 94%，进一步增产依赖 3 号线扩能"),
    ("质量指标全面向好", "一次通过率 99.2%，重大偏差同比下降 45%"),
    ("瓶颈从原料转向后处理", "结晶工序 CT 时间占全流程 38%，是当前唯一瓶颈"),
], at=Rect(left.x, left.y + 0.34, left.w, left.h - 0.34))

s.text(Rect(right.x, right.y, right.w, 0.26), "需要决策的事项",
       size=T.size_h, bold=True)
s.bullets([
    "3 号线结晶工序扩能投资 2,800 万元，请批准立项",
    "Q3 起 A 产品排产上限暂定 240 吨/月",
    "关键岗位增编 12 人，请人力资源部支持",
], at=Rect(right.x, right.y + 0.34, right.w, right.h - 0.34), marker="num")
s.note("数据来源：MES 系统 + ERP 生产模块，统计区间 2026-01-01 至 2026-06-30")

# ---------------------------------------------------------------- 章节 1
deck.section("01", "整体表现", "产量、达成率与产线负荷", agenda=AGENDA)

# 图表页：柱状对比
s = deck.slide(
    title="产量逐月攀升，Q2 月均产量较 Q1 提升 28%",
    eyebrow="产量趋势",
)
chart_area, callout_area = s.body.split_v(0.78, 0.22, gap=0.16)
s.chart(
    "bar",
    ["1月", "2月", "3月", "4月", "5月", "6月"],
    [("2025 年", [148, 132, 165, 171, 178, 182]),
     ("2026 年", [186, 171, 208, 232, 241, 246])],
    at=chart_area,
    number_format="0",
)
s.callout("6 月单月产量 246 吨创历史新高，已达 3 号线满负荷状态",
          at=callout_area, kind="accent", label="关键发现")
s.note("数据来源：MES 系统批次产出记录")

# 达成率页
s = deck.slide(
    title="四条产线中 3 号线是唯一未达标产线，缺口 8.4 个百分点",
    eyebrow="产线达成率",
)
left, right = s.body.split_h(0.52, 0.48, gap=0.36)
s.text(Rect(left.x, left.y, left.w, 0.26), "各产线目标达成率",
       size=T.size_h, bold=True)
s.progress_bars([
    {"label": "1 号线（原料预处理）", "value": 103, "text": "103%"},
    {"label": "2 号线（合成反应）", "value": 101, "text": "101%"},
    {"label": "3 号线（结晶后处理）", "value": 91.6, "text": "91.6%"},
    {"label": "4 号线（包装）", "value": 108, "text": "108%"},
], at=Rect(left.x, left.y + 0.62, left.w, left.h - 0.62), target=100)
s.text(Rect(right.x, right.y, right.w, 0.26), "3 号线缺口归因（影响占比）",
       size=T.size_h, bold=True)
s.chart(
    "hbar",
    ["结晶周期偏长", "设备非计划停机", "人员配置不足", "其他"],
    [("影响占比", [52, 26, 15, 7])],
    at=Rect(right.x, right.y + 0.44, right.w, right.h - 0.44),
    number_format='0"%"',
    colors=["bad", "warn", "neutral", "hairline"],
)
s.note("达成率 = 实际产出 / 年度计划折算至半年的目标值")

# ---------------------------------------------------------------- 章节 2
deck.section("02", "质量与合规", "一次通过率、偏差与审计", agenda=AGENDA)

s = deck.slide(
    title="质量指标全面改善，但供应商来料波动仍是最大风险源",
    eyebrow="质量表现",
)
top, bottom = s.body.split_v(0.52, 0.48, gap=0.26)
s.table(
    ["质量指标", "上半年", "去年同期", "同比", "年度目标", "评价"],
    [
        ["批次一次通过率", "99.2%", "97.8%", "+1.4pp", "≥98.5%", "达成"],
        ["重大偏差", "6 件", "11 件", "-45.5%", "≤10 件", "达成"],
        ["客户投诉", "2 起", "5 起", "-60.0%", "≤3 起", "达成"],
        ["来料检验合格率", "96.4%", "98.1%", "-1.7pp", "≥98.0%", "未达成"],
        ["OOS 调查关闭及时率", "94.0%", "89.0%", "+5.0pp", "≥95.0%", "接近"],
    ],
    at=top,
    col_widths=[2.3, 1.0, 1.0, 0.9, 1.0, 0.9],
    align=["left", "right", "right", "right", "right", "center"],
    cell_colors={(0, 5): "good", (1, 5): "good", (2, 5): "good",
                 (3, 5): "bad", (4, 5): "warn",
                 (0, 3): "good", (1, 3): "good", (2, 3): "good",
                 (3, 3): "bad", (4, 3): "good"},
)
s.icon_cards([
    {"icon": "✓", "title": "已见效的举措",
     "desc": "工艺参数标准化 + 关键岗位复训，一次通过率连续 5 个月上行",
     "color": "good"},
    {"icon": "!", "title": "待解决的问题",
     "desc": "两家原料供应商来料纯度波动，已启动第二供应商审计",
     "color": "warn"},
    {"icon": "→", "title": "下一步动作",
     "desc": "Q3 完成来料放行标准升版，并将纯度纳入供应商 KPI 考核",
     "color": "primary"},
], at=bottom, cols=3)
s.note("数据来源：QMS 系统偏差与投诉台账；来料检验合格率取批次加权值")

# ---------------------------------------------------------------- 章节 3
deck.section("03", "瓶颈与改善", "结晶工序专题", agenda=AGENDA)

s = deck.slide(
    title="结晶工序占全流程周期 38%，是唯一制约产能的瓶颈环节",
    eyebrow="瓶颈定位",
)
top, bottom = s.body.split_v(0.55, 0.45, gap=0.26)
s.chart(
    "hstacked",
    ["A 产品", "B 产品", "C 产品"],
    [("前处理", [12, 14, 11]), ("合成反应", [26, 24, 28]),
     ("结晶后处理", [38, 41, 36]), ("包装检验", [24, 21, 25])],
    at=top,
    number_format='0"%"',
)
s.compare(
    {"title": "改善前（2026 Q1）", "color": "neutral", "items": [
        "结晶批次平均耗时 41 小时",
        "降温曲线依赖人工经验判断",
        "批次间粒径分布波动 ±18%",
    ], "note": "月产能上限 210 吨"},
    {"title": "改善后（试点 3 批次）", "color": "good", "items": [
        "结晶批次平均耗时 29 小时（-29%）",
        "自动降温程序 + 在线粒径监测",
        "批次间粒径分布波动收窄至 ±6%",
    ], "note": "月产能上限可提升至 268 吨"},
    at=bottom,
)
s.note("周期占比按各工序标准工时占单批次总工时计算；试点数据来自 6 月 3 个验证批次")

s = deck.slide(title="扩能方案：优先投自动结晶控制系统，投资回收期 14 个月",
               eyebrow="方案选择")
top, bottom = s.body.split_v(0.56, 0.44, gap=0.26)
s.matrix([
    {"title": "自动结晶控制系统", "color": "good", "items": [
        "投资 2,800 万元", "产能 +27%", "回收期 14 个月"]},
    {"title": "新增结晶釜 2 台", "color": "surface", "items": [
        "投资 5,600 万元", "产能 +42%", "回收期 26 个月"]},
    {"title": "外协后处理", "color": "surface", "items": [
        "无资本支出", "产能 +15%", "单位成本 +22%"]},
    {"title": "维持现状", "color": "surface", "items": [
        "无投入", "Q4 起交付缺口约 90 吨", "存在客户违约风险"]},
], at=top, x_label=("投入强度低", "投入强度高"),
    y_label=("产能收益低", "产能收益高"))
s.timeline([
    {"when": "8 月", "title": "立项审批", "desc": "完成投资评审与预算下达", "status": "doing"},
    {"when": "9-10 月", "title": "设备采购", "desc": "招标与合同签订", "status": "todo"},
    {"when": "11 月", "title": "安装调试", "desc": "利用检修窗口施工", "status": "todo"},
    {"when": "12 月", "title": "验证放行", "desc": "3 批次工艺验证", "status": "todo"},
    {"when": "2027 Q1", "title": "产能释放", "desc": "月产能提升至 268 吨", "status": "todo"},
], at=bottom)
s.note("投资回收期按现有产品结构与 2026 年均价测算，未考虑价格波动")

# ---------------------------------------------------------------- 结尾
deck.closing(
    title="请领导决策的三项事项",
    items=[
        "批准 3 号线自动结晶控制系统立项，投资 2,800 万元",
        "确认 Q3 A 产品排产上限 240 吨/月，超出部分安排外协",
        "支持关键岗位增编 12 人，9 月底前到岗",
    ],
    meta="生产管理部    |    2026 年 7 月 30 日",
)

deck.add_page_numbers()
deck.save(OUT)
print("已生成：", OUT, "共 %d 页" % len(deck.slides))
check_overflow(OUT)
