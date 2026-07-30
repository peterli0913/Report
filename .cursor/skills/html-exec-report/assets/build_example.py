"""生成示例汇报页 example.html（数据为虚构演示数据）。

这个脚本本身就是用法示例：HTML 骨架直接写，图表用 svg_chart 生成后嵌入。
需要做新汇报时，从这里挑结构相近的页型改内容。

运行：
    python3 build_example.py            # 生成 example.html
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))

from svg_chart import bar, hbar, stacked, legend, progress  # noqa: E402

AGENDA = ["整体表现", "质量与合规", "瓶颈与改善", "下一步计划"]


def agenda(active):
    return '<div class="agenda">%s</div>' % "".join(
        '<span class="%s">%02d %s</span>' % ("on" if i == active else "", i, t)
        for i, t in enumerate(AGENDA, 1))


def foot(note, n, total):
    return ('<div class="note">%s</div><div class="page-no">%d / %d</div>'
            % (note, n, total))


P = []      # 各页 HTML
TOTAL = 10


# ---------------------------------------------------------------- 1 封面
P.append("""
<section class="slide slide-cover">
  <div class="kicker">生产管理部 · 半年度经营汇报</div>
  <h1>2026 年上半年<br>生产运营回顾</h1>
  <p class="sub">产能达成、质量合规与产线瓶颈的系统性复盘</p>
  <div class="meta">汇报人：张三 &nbsp;&nbsp;|&nbsp;&nbsp; 2026 年 7 月 30 日
    &nbsp;&nbsp;|&nbsp;&nbsp; 内部资料 请勿外传</div>
</section>
""")

# ---------------------------------------------------------------- 2 执行摘要
P.append("""
<section class="slide">
  <div class="slide-head">
    <div class="eyebrow">执行摘要</div>
    <h3 class="slide-title">上半年产量超目标 5.2 个百分点，但 3 号线瓶颈已开始制约三季度交付</h3>
  </div>
  <div class="body">
    <div class="kpi-row" style="flex:0 0 168px">
      <div class="kpi"><div class="v">1,284<span class="u">吨</span></div>
        <div class="l">上半年总产量</div><div class="d t-good">▲ 同比 +23.1%</div></div>
      <div class="kpi hl"><div class="v">105.2<span class="u">%</span></div>
        <div class="l">年度目标达成率</div><div class="d t-good">▲ 高于计划 5.2pp</div></div>
      <div class="kpi"><div class="v">99.2<span class="u">%</span></div>
        <div class="l">批次一次通过率</div><div class="d t-good">▲ 同比 +1.4pp</div></div>
      <div class="kpi"><div class="v">6<span class="u">件</span></div>
        <div class="l">重大偏差</div><div class="d t-good">▼ 同比 -45%</div></div>
    </div>
    <div class="row grow">
      <div class="col" style="flex:1.05">
        <div class="section-label">三个结论</div>
        <ul class="bullets">
          <li><span class="h">产能已释放到设计上限</span>
            <span class="d">1、2 号线利用率 94%，进一步增产依赖 3 号线扩能</span></li>
          <li><span class="h">质量指标全面向好</span>
            <span class="d">一次通过率 99.2%，重大偏差同比下降 45%</span></li>
          <li><span class="h">瓶颈从原料转向后处理</span>
            <span class="d">结晶工序占全流程周期 38%，是当前唯一瓶颈</span></li>
        </ul>
      </div>
      <div class="col" style="flex:1">
        <div class="section-label">需要决策的事项</div>
        <ul class="bullets num">
          <li>3 号线结晶工序扩能投资 2,800 万元，请批准立项</li>
          <li>Q3 起 A 产品排产上限暂定 240 吨 / 月</li>
          <li>关键岗位增编 12 人，请人力资源部支持</li>
        </ul>
      </div>
    </div>
  </div>
""" + foot("数据来源：MES 系统 + ERP 生产模块，统计区间 2026-01-01 至 2026-06-30",
            2, TOTAL) + "</section>")

# ---------------------------------------------------------------- 3 章节 01
P.append("""
<section class="slide slide-section">
  <div class="num">01</div>
  <h2>整体表现</h2>
  <p class="sub">产量、达成率与产线负荷</p>
""" + agenda(1) + "</section>")

# ---------------------------------------------------------------- 4 产量趋势
cats = ["1月", "2月", "3月", "4月", "5月", "6月"]
series2 = [("2025 年", [148, 132, 165, 171, 178, 182]),
           ("2026 年", [186, 171, 208, 232, 241, 246])]
P.append("""
<section class="slide">
  <div class="slide-head">
    <div class="eyebrow">产量趋势</div>
    <h3 class="slide-title">产量逐月攀升，Q2 月均产量较 Q1 提升 28%%</h3>
  </div>
  <div class="body">
    <div style="flex:0 0 404px">%s%s</div>
    <div class="callout"><span class="tag">关键发现</span>
      6 月单月产量 246 吨创历史新高，已达 3 号线满负荷状态</div>
  </div>
%s</section>""" % (legend(series2), bar(cats, series2, height=380, unit=""),
                   foot("数据来源：MES 系统批次产出记录", 4, TOTAL)))

# ---------------------------------------------------------------- 5 达成率
P.append("""
<section class="slide">
  <div class="slide-head">
    <div class="eyebrow">产线达成率</div>
    <h3 class="slide-title">四条产线中 3 号线是唯一未达标产线，缺口 8.4 个百分点</h3>
  </div>
  <div class="body">
    <div class="row grow">
      <div class="col" style="flex:1.1">
        <div class="section-label">各产线目标达成率</div>
        %s
      </div>
      <div class="col" style="flex:1">
        <div class="section-label">3 号线缺口归因（影响占比）</div>
        <div class="grow" style="display:flex;align-items:center">%s</div>
      </div>
    </div>
  </div>
%s</section>""" % (
    progress([
        {"label": "1 号线（原料预处理）", "value": 103, "text": "103%"},
        {"label": "2 号线（合成反应）", "value": 101, "text": "101%"},
        {"label": "3 号线（结晶后处理）", "value": 91.6, "text": "91.6%"},
        {"label": "4 号线（包装）", "value": 108, "text": "108%"},
    ]),
    hbar(["结晶周期偏长", "设备非计划停机", "人员配置不足", "其他"],
         [52, 26, 15, 7], colors=["bad", "warn", "neutral", "hairline"],
         width=560, height=340, unit="%"),
    foot("达成率 = 实际产出 / 年度计划折算至半年的目标值", 5, TOTAL)))

# ---------------------------------------------------------------- 6 章节 02
P.append("""
<section class="slide slide-section">
  <div class="num">02</div>
  <h2>质量与合规</h2>
  <p class="sub">一次通过率、偏差与审计</p>
""" + agenda(2) + "</section>")

# ---------------------------------------------------------------- 7 质量
P.append("""
<section class="slide">
  <div class="slide-head">
    <div class="eyebrow">质量表现</div>
    <h3 class="slide-title">质量指标全面改善，但供应商来料波动仍是最大风险源</h3>
  </div>
  <div class="body">
    <table class="table">
      <thead><tr><th>质量指标</th><th class="num">上半年</th><th class="num">去年同期</th>
        <th class="num">同比</th><th class="num">年度目标</th><th class="mid">评价</th></tr></thead>
      <tbody>
        <tr><td>批次一次通过率</td><td class="num">99.2%</td><td class="num">97.8%</td>
          <td class="num t-good">+1.4pp</td><td class="num">≥98.5%</td>
          <td class="mid t-good">达成</td></tr>
        <tr><td>重大偏差</td><td class="num">6 件</td><td class="num">11 件</td>
          <td class="num t-good">-45.5%</td><td class="num">≤10 件</td>
          <td class="mid t-good">达成</td></tr>
        <tr><td>客户投诉</td><td class="num">2 起</td><td class="num">5 起</td>
          <td class="num t-good">-60.0%</td><td class="num">≤3 起</td>
          <td class="mid t-good">达成</td></tr>
        <tr><td>来料检验合格率</td><td class="num">96.4%</td><td class="num">98.1%</td>
          <td class="num t-bad">-1.7pp</td><td class="num">≥98.0%</td>
          <td class="mid t-bad">未达成</td></tr>
        <tr><td>OOS 调查关闭及时率</td><td class="num">94.0%</td><td class="num">89.0%</td>
          <td class="num t-good">+5.0pp</td><td class="num">≥95.0%</td>
          <td class="mid t-warn">接近</td></tr>
      </tbody>
    </table>
    <div class="card-grid cols-3 grow">
      <div class="card icon-card">
        <div class="icon-badge bg-good">✓</div>
        <h4 class="card-title">已见效的举措</h4>
        <p class="card-desc">工艺参数标准化 + 关键岗位复训，一次通过率连续 5 个月上行</p>
      </div>
      <div class="card icon-card">
        <div class="icon-badge bg-warn">!</div>
        <h4 class="card-title">待解决的问题</h4>
        <p class="card-desc">两家原料供应商来料纯度波动，已启动第二供应商审计</p>
      </div>
      <div class="card icon-card">
        <div class="icon-badge bg-primary">→</div>
        <h4 class="card-title">下一步动作</h4>
        <p class="card-desc">Q3 完成来料放行标准升版，并将纯度纳入供应商 KPI 考核</p>
      </div>
    </div>
  </div>
""" + foot("数据来源：QMS 系统偏差与投诉台账；来料检验合格率取批次加权值", 7, TOTAL)
         + "</section>")

# ---------------------------------------------------------------- 8 章节 03
P.append("""
<section class="slide slide-section">
  <div class="num">03</div>
  <h2>瓶颈与改善</h2>
  <p class="sub">结晶工序专题</p>
""" + agenda(3) + "</section>")

# ---------------------------------------------------------------- 9 瓶颈
stk = [("前处理", [12, 14, 11]), ("合成反应", [26, 24, 28]),
       ("结晶后处理", [38, 41, 36]), ("包装检验", [24, 21, 25])]
P.append("""
<section class="slide">
  <div class="slide-head">
    <div class="eyebrow">瓶颈定位</div>
    <h3 class="slide-title">结晶工序占全流程周期 38%%，是唯一制约产能的瓶颈环节</h3>
  </div>
  <div class="body">
    %s
    <div style="flex:0 0 216px">%s</div>
    <div class="compare grow">
      <div class="cmp">
        <h4 class="bg-neutral">改善前（2026 Q1）</h4>
        <div class="in"><ul class="bullets">
          <li>结晶批次平均耗时 41 小时</li>
          <li>降温曲线依赖人工经验判断</li>
          <li>批次间粒径分布波动 ±18%%</li>
        </ul></div>
        <div class="foot">月产能上限 210 吨</div>
      </div>
      <div class="cmp">
        <h4 class="bg-good">改善后（试点 3 批次）</h4>
        <div class="in"><ul class="bullets">
          <li>结晶批次平均耗时 29 小时（-29%%）</li>
          <li>自动降温程序 + 在线粒径监测</li>
          <li>批次间粒径分布波动收窄至 ±6%%</li>
        </ul></div>
        <div class="foot">月产能上限可提升至 268 吨</div>
      </div>
    </div>
  </div>
%s</section>""" % (
    legend(stk),
    stacked(["A 产品", "B 产品", "C 产品"], stk, width=1160, height=200),
    foot("周期占比按各工序标准工时占单批次总工时计算；试点数据来自 6 月 3 个验证批次",
         9, TOTAL)))

# ---------------------------------------------------------------- 10 结尾
P.append("""
<section class="slide slide-cover">
  <h1>请领导决策的三项事项</h1>
  <ul class="bullets num" style="margin-top:34px;padding-left:27px;
      font-size:var(--fs-cover-sub)">
    <li>批准 3 号线自动结晶控制系统立项，投资 2,800 万元</li>
    <li>确认 Q3 A 产品排产上限 240 吨 / 月，超出部分安排外协</li>
    <li>支持关键岗位增编 12 人，9 月底前到岗</li>
  </ul>
  <div class="meta">生产管理部 &nbsp;&nbsp;|&nbsp;&nbsp; 2026 年 7 月 30 日</div>
</section>
""")

HTML = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>2026 年上半年生产运营回顾</title>
<link rel="stylesheet" href="deck.css">
</head>
<body>
<div class="progress-top"></div>
<div class="deck">
%s
</div>
<div class="nav">
  <button data-go="prev" title="上一页">‹</button>
  <span class="cur"></span>
  <button data-go="next" title="下一页">›</button>
  <span style="margin-left:6px">← → 翻页 · F 全屏 · P 打印</span>
</div>
<script src="deck.js"></script>
</body>
</html>
""" % "\n".join(p.strip() for p in P)

out = os.path.join(HERE, "example.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(HTML)
print("已生成：%s（%d 页）" % (out, len(P)))
