---
name: html-exec-report
description: 当需要做网页形式的汇报材料时使用 —— HTML 汇报页、可翻页的网页版演示、单文件 html 汇报、能发链接或发附件的汇报、可在浏览器全屏演示并打印成 PDF 的材料。也用于把已有的 PPT 内容改成网页版、做数据看板式的汇报页、或者用户说"做成 html""做个网页版""不想用 PPT"。
---

# HTML 汇报页

产出可在浏览器全屏演示、键盘翻页、并能直接打印成 PDF 的汇报页。
零运行时依赖（不用 React、不用图表库、不连 CDN），内网和离线环境都能打开。

## 什么时候用 HTML，什么时候用 PPT

| 场景 | 选择 |
|---|---|
| 对方要拿去改内容、加页、走审批流 | **PPT**（`exec-deck-builder`） |
| 正式会议投屏、需要留档为 pptx | **PPT** |
| 发链接给多人看、手机上也要能读 | HTML |
| 有交互需求（悬停看明细、切换视图） | HTML |
| 要嵌进内部系统、门户、看板 | HTML |
| 只是要一份能打印的 PDF | HTML（打印更可控，字体不会跳） |

拿不准就做 PPT —— 中国企业的汇报流程基本都在 pptx 上流转。

## 快速开始

从示例改，不要从零写：

```bash
cd .cursor/skills/html-exec-report
python3 assets/build_example.py          # 生成 assets/example.html
python3 scripts/shoot.py assets/example.html preview/   # 逐页截图看效果
```

`assets/example.html` 是一份完整的 10 页汇报，涵盖全部页型。做新汇报时复制它，
挑结构最像的页改内容。`assets/build_example.py` 展示了如何用 Python 组装页面并
嵌入图表。

最小骨架：

```html
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>汇报标题</title>
  <link rel="stylesheet" href="deck.css">
</head>
<body>
  <div class="progress-top"></div>
  <div class="deck">
    <section class="slide slide-cover">
      <div class="kicker">生产管理部</div>
      <h1>2026 年上半年<br>生产运营回顾</h1>
      <p class="sub">产能、质量与瓶颈复盘</p>
      <div class="meta">汇报人：张三 | 2026-07-30 | 内部资料</div>
    </section>

    <section class="slide">
      <div class="slide-head">
        <div class="eyebrow">执行摘要</div>
        <h3 class="slide-title">产量超目标 5.2 个百分点，3 号线已成瓶颈</h3>
      </div>
      <div class="body"><!-- 内容 --></div>
      <div class="note">数据来源：MES 系统</div>
      <div class="page-no">2 / 10</div>
    </section>
  </div>
  <div class="nav">
    <button data-go="prev">‹</button><span class="cur"></span>
    <button data-go="next">›</button>
  </div>
  <script src="deck.js"></script>
</body>
</html>
```

`deck.css` 和 `deck.js` 从 `assets/` 复制到输出目录，或者最后用 `inline.py` 内联。

## 布局系统

页面是**固定 1280x720 px** 设计尺寸，`deck.js` 按视口等比缩放 —— 和 PPT 一样精确
可控，不会因为窗口大小把版面挤乱。字号按 96dpi 与 PPT 版一一对应，同一份内容做
两种载体观感一致。

布局用 flex 工具类，不要写绝对定位：

```html
<div class="body">                          <!-- 纵向排列，自动 gap -->
  <div class="row grow">                    <!-- 横向排列，占满剩余高度 -->
    <div class="col" style="flex:1.1">…</div>   <!-- 左列，权重 1.1 -->
    <div class="col" style="flex:1">…</div>     <!-- 右列 -->
  </div>
  <div class="callout">…</div>               <!-- 底部结论条，高度自适应 -->
</div>
```

| 类 | 作用 |
|---|---|
| `.body` | 内容区，纵向 flex，子元素间自动 gap |
| `.row` / `.col` | 横向 / 纵向 flex 容器 |
| `.grow` | 占满剩余空间（`flex:1; min-height:0`） |
| `.card-grid.cols-2/3/4` | 等宽网格 |

**图表容器要给明确高度**（`style="flex:0 0 380px"`），因为 SVG 用
`preserveAspectRatio` 保持比例，容器高度不确定时会留白。

## 组件速查

| 组件 | 类名 | 说明 |
|---|---|---|
| 封面 | `.slide.slide-cover` | 内含 `.kicker` `h1` `.sub` `.meta` |
| 章节页 | `.slide.slide-section` | 内含 `.num` `h2` `.sub` `.agenda` |
| 页头 | `.slide-head` > `.eyebrow` + `.slide-title` | 标题写结论，不写栏目名 |
| KPI 卡 | `.kpi-row` > `.kpi` > `.v`/`.u`/`.l`/`.d` | 加 `.hl` 让数字用强调色 |
| 卡片 | `.card` / `.card-grid.cols-3` | `.card-title` + `.card-desc` |
| 图标卡 | `.card.icon-card` > `.icon-badge` | badge 加 `.bg-good` 等着色 |
| 要点 | `ul.bullets` > `li` > `.h` + `.d` | 加 `.num` 变编号列表 |
| 表格 | `table.table`，数字列加 `.num` | 数字列一律右对齐 |
| 达成率条 | `.progress-list` > `.pr` | 用 `svg_chart.progress()` 生成 |
| 时间轴 | `.timeline` > `.tl` > `.when`/`.dot`/`.t`/`.d` | — |
| 左右对比 | `.compare` > `.cmp` > `h4` + `.in` + `.foot` | h4 加 `.bg-good` 等着色 |
| 四象限 | `.matrix` > `.quad` | 按左上/右上/左下/右下顺序 |
| 结论条 | `.callout`（`.good`/`.warn`/`.bad`/`.plain`） | 每页最多一个 |
| 页脚 | `.note` + `.page-no` | 有数字的页必须写 note |

语义色工具类：文字用 `.t-good` `.t-warn` `.t-bad` `.t-neutral` `.t-accent` `.t-muted`；
底色用 `.bg-good` `.bg-warn` `.bg-bad` `.bg-neutral` `.bg-primary` `.bg-accent`。

**文字色和底色是两组不同的值**（`--good` 和 `--good-text`）：饱和的绿做色块很好，
当成 16px 小字压在深底上就达不到 4.5:1 对比度。用工具类就不会用错。

## 图表

用 `scripts/svg_chart.py` 生成内联 SVG，不要引图表库 —— CDN 在内网拉不到会整页
空白，canvas 图表打印时经常输出空白。

```python
import sys; sys.path.insert(0, ".cursor/skills/html-exec-report/scripts")
from svg_chart import bar, hbar, line, stacked, donut, legend, progress

series = [("2025 年", [148, 132, 165]), ("2026 年", [186, 171, 208])]
svg = legend(series) + bar(["1月", "2月", "3月"], series, height=380, unit=" 吨")
```

| 函数 | 用途 |
|---|---|
| `bar(categories, series, ...)` | 纵向分组柱状，看趋势和分组比较 |
| `hbar(categories, values, ...)` | 横向条形，类目名长或做排名时用；传入顺序即显示顺序，第一项在最上 |
| `line(categories, series, area=False)` | 折线，看趋势 |
| `stacked(categories, series, horizontal=True)` | 堆积条形，看构成 |
| `donut(labels, values, ...)` | 环形，只用于单一构成且不超过 5 段 |
| `legend(series)` | 图例，多序列时必须给 |
| `progress(items, target)` | 达成率条（输出 HTML 不是 SVG） |

颜色默认取 CSS 变量（`var(--primary)` 等），所以切换浅色主题时图表自动跟着变。
也可以传语义名：`colors=["bad", "warn", "neutral"]`。

## 主题

深色是默认。要浅色版（打印、亮环境）在 `<body>` 上加属性：

```html
<body data-theme="light">
```

改品牌色覆盖 CSS 变量即可，全套组件和图表会自动跟随：

```html
<style>
  :root { --bg:#003669; --primary:#3263A7; --accent:#E5B620; }
</style>
```

改完颜色要重新截图检查对比度 —— 换了底色，原本达标的文字可能就不达标了。

## 工作流

1. **写页面清单**：每页一行"行动标题 + 版式"，确认后再动手。
2. **复制 `example.html` 改内容**，图表用 `svg_chart.py` 生成。
3. **逐页截图复核**：`python3 scripts/shoot.py 汇报.html preview/`。
   HTML 不看渲染结果等于没测 —— 溢出、错位只有截图才看得见。
4. **打包分发**：`python3 scripts/inline.py 汇报.html 汇报-单文件.html`。

## 演示与分发

| 需求 | 做法 |
|---|---|
| 键盘翻页 | `← →` `↑ ↓` `PageUp/Down` `空格` |
| 全屏 | 按 `F` |
| 跳到某页 | URL 加 `#5` |
| 打印 / 导出 PDF | 按 `P`，或浏览器打印。已配 A4 横向、每页一张 |
| 发给别人 | 用 `inline.py` 打包成单文件 html，双击即可看 |
| 隐藏翻页控件 | URL 加 `?hideui=1`（截图、嵌入时用） |

**打印深色底需要在打印设置里勾选"背景图形"**，否则底色不输出。要交付 PDF 时
建议直接用浅色主题（`data-theme="light"`），省掉这个坑。

## 常见错误

- **给图表容器用 `flex:1` 但不给高度**：SVG 保持比例居中，容器过高就会大片留白。
  给明确高度（`flex:0 0 380px`）或用 `align-items:center` 居中。
- **忘了 `min-height:0`**：flex 子项默认不能收缩到内容以下，嵌套 flex 时内容会
  溢出容器。`.grow` 已经带了这个属性，自己写 `flex:1` 时要记得加。
- **引 CDN 上的字体或图表库**：内网打不开，演示现场变白屏。所有资源都要本地化。
- **直接把 css/js 和 html 分开发给别人**：附件里只有 html，样式全丢。
  用 `inline.py` 打包。
- **正文字号小于 16px**：投屏后排看不清。`--fs-body` 是 19px，不要往下调。
- **用 `emoji` 当图标**：不同系统渲染成不同样式。用 `.icon-badge` 里的符号，
  或者 `deck-imagery` skill 生成的图标。
- **一页塞太多内容**：HTML 页面不会像 PPT 那样把文字裁掉，而是会溢出或压缩，
  更难发现。正文控制在 300 字以内。

## 相关 skill

- 内容结构、指标选择、结论提炼 → `production-report-narrative`
- 配色和版式的取舍依据 → `exec-visual-system`
- 配图与图标 → `deck-imagery`
- 同一份内容要交付 pptx → `exec-deck-builder`
