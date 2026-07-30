#!/usr/bin/env python3
"""deckkit 自测：跑一遍全部组件 x 全部主题 x 全部画布，确认没有回归。

改动 deckkit.py 之后先跑这个，再去生成正式汇报 —— 排版类 bug 在单页上不明显，
换个画布或换个主题才会暴露（例如坐标写死、颜色对比度不足）。

用法：
    python3 selftest.py [输出目录]
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from deckkit import (Deck, Rect, THEMES, audit_theme, check_overflow, fit_size,
                     text_width, wrapped_lines)


def build(theme, canvas, path):
    deck = Deck(theme=theme, canvas=canvas)
    T = deck.theme

    deck.cover(title="组件自测：%s / %s 画布" % (theme, canvas),
               subtitle="覆盖 KPI、图表、表格、进度条、时间轴、对比、四象限",
               kicker="deckkit selftest", meta="自动生成 · 不含真实数据")

    # KPI + 要点
    s = deck.slide(title="KPI 卡片与要点列表在两种画布下应完全等比",
                   eyebrow="组件 01")
    # 按组件的固有高度切区域，而不是按比例平分：KPI 卡固定约 1.95"，
    # 底部结论条 0.86"。比例平分会在页面中间留下一条空带。
    body = s.body
    top = Rect(body.x, body.y, body.w, 1.95)
    bot = Rect(body.x, top.bottom + 0.26, body.w,
               body.bottom - top.bottom - 0.26 - 1.06)
    s.kpi_row([
        {"value": "1,284", "unit": "吨", "label": "总产量",
         "delta": "+23.1%", "trend": "up"},
        {"value": "105.2", "unit": "%", "label": "达成率", "delta": "+5.2pp",
         "trend": "up", "color": "accent"},
        # 偏差数下降是好事，必须显式覆盖颜色
        {"value": "6", "unit": "件", "label": "重大偏差", "delta": "-45%",
         "trend": "down", "delta_color": "good"},
    ], at=top)
    a, b = bot.split_h(1, 1, gap=0.3)
    s.bullets([("圆点要点", "带说明的两级结构"), ("第二条", "说明文字"),
               ("第三条", "说明文字")], at=a)
    s.bullets(["序号要点一", "序号要点二", "序号要点三"], at=b, marker="num")
    s.callout("全页唯一的结论条，用强调色承载一句话结论", kind="accent",
              label="结论")
    s.note("自测数据，无业务含义")

    # 图表矩阵
    s = deck.slide(title="五种图表的默认样式都应无边框、有数据标签", eyebrow="组件 02")
    cells = s.body.grid(3, 2, gap=0.26)
    cats = ["1月", "2月", "3月", "4月"]
    s.chart("bar", cats, [("A", [12, 18, 15, 22]), ("B", [8, 11, 14, 17])],
            at=cells[0])
    s.chart("line", cats, [("趋势", [12, 18, 15, 22])], at=cells[1])
    s.chart("hbar", ["甲", "乙", "丙"], [("占比", [52, 30, 18])], at=cells[2],
            number_format='0"%"')
    s.chart("stacked", cats, [("前段", [30, 32, 28, 30]),
                              ("中段", [40, 38, 42, 40]),
                              ("后段", [30, 30, 30, 30])], at=cells[3],
            number_format='0"%"')
    s.chart("doughnut", ["达标", "接近", "未达标"], [("分布", [62, 26, 12])],
            at=cells[4])
    s.chart("area", cats, [("累计", [12, 30, 45, 67])], at=cells[5])

    # 表格 + 进度条
    s = deck.slide(title="表格数字列右对齐，进度条按最大值缩放", eyebrow="组件 03")
    # 表格高度 = 表头 + 行高 x 行数。区域给大了表格不会跟着长高，只会在下方留白，
    # 所以先算出高度再切区域（这一点和其他组件相反）。
    row_h, header_h, rows = 0.52, 0.44, 3
    body = s.body
    top = Rect(body.x, body.y, body.w, header_h + row_h * rows)
    bot = Rect(body.x, top.bottom + 0.30, body.w,
               body.bottom - top.bottom - 0.30)
    s.table(["项目", "本期", "同期", "同比", "评价"],
            [["一次通过率", "99.2%", "97.8%", "+1.4pp", "达成"],
             ["重大偏差", "6 件", "11 件", "-45.5%", "达成"],
             ["来料合格率", "96.4%", "98.1%", "-1.7pp", "未达成"]],
            at=top, col_widths=[2.2, 1, 1, 1, 0.9], row_h=row_h,
            header_h=header_h,
            align=["left", "right", "right", "right", "center"],
            cell_colors={(0, 4): "good", (1, 4): "good", (2, 4): "bad"})
    s.progress_bars([
        {"label": "1 号线", "value": 103}, {"label": "2 号线", "value": 91.6},
        {"label": "3 号线", "value": 108}, {"label": "4 号线", "value": 78},
    ], at=bot, target=100, align="fill")

    # 时间轴 / 对比 / 四象限 / 图标卡
    s = deck.slide(title="横竖时间轴、左右对比、四象限、图标卡", eyebrow="组件 04")
    top, bot = s.body.split_v(0.44, 0.56, gap=0.26)
    ta, tb = top.split_h(0.62, 0.38, gap=0.3)
    s.timeline([{"when": "Q1", "title": "立项", "status": "done"},
                {"when": "Q2", "title": "建设", "status": "doing"},
                {"when": "Q3", "title": "验证", "status": "todo"}], at=ta)
    s.timeline([{"when": "8 月", "title": "审批"},
                {"when": "9 月", "title": "采购"}], at=tb, orient="v")
    ba, bb = bot.split_h(1, 1, gap=0.3)
    s.compare({"title": "改善前", "color": "neutral", "items": ["耗时 41 小时"]},
              {"title": "改善后", "color": "good", "items": ["耗时 29 小时"]},
              at=ba)
    s.matrix([{"title": "优先做", "color": "good", "items": ["高收益低投入"]},
              {"title": "评估", "items": ["高收益高投入"]},
              {"title": "顺手做", "items": ["低收益低投入"]},
              {"title": "不做", "items": ["低收益高投入"]}],
             at=bb, x_label=("投入低", "投入高"), y_label=("收益低", "收益高"))

    s = deck.slide(title="图标卡与结论条：每页只留一个重点", eyebrow="组件 05")
    # callout 不传 at 时贴页面底部（约 0.86" 高），卡片区要让出这段加间距
    top = Rect(s.body.x, s.body.y, s.body.w, s.body.h - 1.06)
    s.icon_cards([
        {"icon": "✓", "title": "已见效", "desc": "描述文字，用于说明这一项的具体内容",
         "color": "good"},
        {"icon": "!", "title": "待解决", "desc": "描述文字", "color": "warn"},
        {"icon": "→", "title": "下一步", "desc": "描述文字", "color": "primary"},
    ], at=top, cols=3)
    s.callout("这是全页唯一的结论条，用强调色承载一句话结论",
              kind="accent", label="结论")

    deck.closing(title="需要决策的事项",
                 items=["事项一", "事项二", "事项三"], meta="deckkit selftest")
    deck.add_page_numbers()
    deck.save(path)
    return deck


def main():
    outdir = sys.argv[1] if len(sys.argv) > 1 else "selftest-out"
    os.makedirs(outdir, exist_ok=True)

    # 度量函数的基本正确性：中文应按全角算宽，别退化成半角
    assert text_width("中文", 14) > text_width("ab", 14), "中文宽度估算异常"
    assert wrapped_lines("中" * 40, 2.0, 14) > 1, "折行估算异常"
    assert fit_size("中" * 200, 3.0, 0.6, 20) < 20, "自动降字号未生效"
    print("度量函数检查通过")

    # 主题配色：改过色值后必须全部达标，否则生成的每一页都会带着问题
    theme_bad = 0
    for tn in ("dark", "light", "slate"):
        probs = audit_theme(THEMES[tn])
        theme_bad += len(probs)
        print("主题 %-6s 配色检查 %s" % (tn, "通过" if not probs
                                     else "%d 个问题" % len(probs)))
        for p in probs:
            print("     " + p)

    # 单序列柱形图传 colors 时必须按数据点着色。PowerPoint 的着色单位是序列，
    # 按序列上色会让所有柱子同色 —— 传了四个颜色却全渲染成第一个，是静默错误。
    probe = Deck(theme="dark")
    ps = probe.slide(title="着色探针")
    ch = ps.chart("hbar", ["甲", "乙", "丙"], [("占比", [50, 30, 20])],
                  at=ps.body, colors=["bad", "warn", "neutral"])
    pt_colors = [str(p.format.fill.fore_color.rgb)
                 for p in ch.plots[0].series[0].points]
    assert len(set(pt_colors)) == 3, "单序列 colors 未按数据点生效：%s" % pt_colors
    # hbar 内部反转了类别，颜色要跟着反转，第一项仍应拿到传入的第一个颜色
    assert pt_colors[-1] == probe.theme.color("bad"), \
        "hbar 反转后颜色未对齐类别：%s" % pt_colors
    print("图表着色检查通过")

    total = theme_bad
    for theme in ("dark", "light", "slate"):
        for canvas in ("wide", "large"):
            path = os.path.join(outdir, "selftest-%s-%s.pptx" % (theme, canvas))
            deck = build(theme, canvas, path)
            issues = check_overflow(path, verbose=False)
            total += len(issues)
            flag = "OK" if not issues else "%d 个问题" % len(issues)
            print("%-28s %d 页  %s" % (os.path.basename(path),
                                       len(deck.slides), flag))
            for it in issues:
                print("     第%d页 [%s] %s" % (it["slide"], it["kind"], it["msg"]))
    print("\n合计问题数：%d" % total)
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
