---
name: exec-deck-builder
description: 当需要制作汇报 PPT、演示文稿、幻灯片、deck，或者需要读取/修改已有 pptx 文件时使用。包括月度/季度/半年度经营汇报、生产运营汇报、专题汇报、述职报告、董事会与管理层汇报、投资立项汇报、复盘总结，以及"把这些数据做成 PPT""美化一下这个 PPT""做一页 slide"这类要求。也用于需要在幻灯片里放图表、KPI 大数字、对比、时间轴、四象限的场合。
---

# 高层汇报 PPT 生成

用 `scripts/deckkit.py`（基于 python-pptx）生成幻灯片。产出的是真正的 pptx：
文字可编辑、图表是原生图表，用户能在 PowerPoint 里继续改，不是一张图片。

**核心原则：先定每一页的结论，再选承载它的版式。** 版式服务于结论，不是反过来。

## 何时用 / 何时不用

- 用：要交付 `.pptx` 文件的任何场合。
- 不用：要交付网页/单文件 HTML 汇报 → 用 `html-exec-report` skill。
- 内容怎么写、指标怎么选、结论怎么提炼 → 配合 `production-report-narrative` skill。
- 配图和图标 → 配合 `deck-imagery` skill。
- 配色和版式规范的取舍依据 → 配合 `exec-visual-system` skill。

## 五步工作流（不要跳步）

1. **先写页面清单**：每页一行"行动标题 + 用什么版式"。如果能和用户对话，先给他
   确认；不能对话时把清单写进生成脚本的注释里，然后继续做。
   行动标题指把结论写进标题，而不是写栏目名：
   - 不写"上半年产量情况" → 写"产量超目标 5.2 个百分点，3 号线已成瓶颈"
   - 不写"质量指标分析" → 写"质量指标全面改善，来料波动是最大风险源"
2. **写生成脚本**，`import deckkit`，从 `assets/example_deck.py` 找结构最像的页型改。
   切区域前先看下面的「组件占位高度」表 —— 组件不会撑满你给的区域。
3. **跑 `check_deck.py --strict`**：error 必须清零；warn 要逐条看过再决定放不放过
   （加了 `--strict` 时 warn 也会让退出码变成 1，这是故意的，逼你看一眼）。
4. **渲染成图逐页看**（`render_deck.py`）。质检只保证"文字读得出来"，保证不了
   "版面好看"。留白是否均匀、重点是否突出，只有看图才知道。
5. **报告数据来源**：每页有数字就调 `s.note("数据来源：...")`。高层一定会问数字哪来的。

## 快速开始

```python
import sys; sys.path.insert(0, ".cursor/skills/exec-deck-builder/scripts")
from deckkit import Deck, Rect, check_overflow

deck = Deck(theme="dark", canvas="wide")   # theme: dark/light/slate；canvas: wide/large

deck.cover(title="2026 年上半年生产运营回顾", subtitle="产能、质量与瓶颈复盘",
           kicker="生产管理部", meta="汇报人：张三  |  2026-07-30")

deck.section("01", "整体表现", agenda=["整体表现", "质量合规", "改善计划"])

s = deck.slide(title="产量超目标 5.2 个百分点，但 3 号线已成瓶颈", eyebrow="整体表现")
top, bottom = s.body.split_v(0.42, 0.58, gap=0.26)      # 上下分区
s.kpi_row([
    {"value": "1,284", "unit": "吨", "label": "总产量", "delta": "+23.1%", "trend": "up"},
    {"value": "105.2", "unit": "%", "label": "达成率", "delta": "+5.2pp",
     "trend": "up", "color": "accent"},
], at=top)
s.chart("bar", ["1月", "2月", "3月"], [("2026", [186, 171, 208])], at=bottom)
s.note("数据来源：MES 系统，截至 2026-06-30")

deck.closing(items=["批准 3 号线扩能立项", "确认 Q3 排产上限"])
deck.add_page_numbers()
deck.save("汇报.pptx")
check_overflow("汇报.pptx")       # 生成后立刻自查
```

## 坐标系统

组件的坐标一律用**设计英寸**，基准画布 13.333 x 7.5。换画布时组件代码一个字都不用
改，库内部按 `deck.scale` 换算。

| `canvas` | 尺寸 | 什么时候用 |
|---|---|---|
| `"wide"` | 13.333 x 7.5 | **默认选这个**。PowerPoint 标准 16:9，对方打开、改内容都最省事 |
| `"large"` | 26.67 x 15 | 要和公司现有模板（双倍画布）混排、或者要把生成的页插进既有 pptx 时 |
| `"a4"` | 11.69 x 8.27 | 主要是打印用 |

**不要自己写死英寸数或 pt 数**，用区域切分拿位置：

```python
s.body                              # 标题以下、脚注以上的内容区
s.body.split_v(0.4, 0.6, gap=0.26)  # 上下两块，比例 4:6
s.body.split_h(1, 1, gap=0.3)       # 左右等分
s.body.grid(3, 2, gap=0.24)         # 3 列 2 行，行优先返回 6 块
area.inset(0.2)                     # 四周内缩
area.pad(top=0.4)                   # 只缩某一边
```

### 组件占位高度（切区域前先看这个）

**组件不会撑满你给它的区域。** 按比例平分 `body` 往往在页面中间留下一条空白带，
这是最常见的排版返工原因。有固有高度的组件，要先算高度再切区域：

| 组件 | 实际占用高度（设计英寸） |
|---|---|
| `kpi_row` | 固定 `1.95`（不传 `height` 时），不随区域变高 |
| `table` | `header_h + row_h × 行数`；自动行高**上限 0.40**，所以 4 行表最多约 1.9 高。要撑满就显式传 `row_h` |
| `progress_bars` | `n × (bar_h + gap) − gap`，开 `show_target` 时上方还要 `0.28` 放目标标签 |
| `callout` | 最小 `0.86`，不传 `at` 时贴页面底部通栏 |
| `bullets` | 每条约 `size×1.3/72 + 说明行高 + gap`，返回排完后的 y 坐标，可以接着往下放 |
| `icon_cards` / `compare` / `matrix` / `timeline` | 卡片会撑满区域，可以放心用比例切分 |
| `chart` | 撑满区域 |

典型写法（表格 + 达成率条，页面不留空带）：

```python
row_h, header_h, rows = 0.52, 0.44, 3
body = s.body
top = Rect(body.x, body.y, body.w, header_h + row_h * rows)   # 按表格实际高度切
bot = Rect(body.x, top.bottom + 0.30, body.w, body.bottom - top.bottom - 0.30)
s.table(headers, data, at=top, row_h=row_h, header_h=header_h)
s.progress_bars(items, at=bot, align="fill")
```

`check_deck.py` 的 `dead_space` 警告就是在报这个问题，看到就按上面的方式改。

## 组件速查

| 组件 | 用途 | 关键参数 |
|---|---|---|
| `deck.cover()` | 封面 | title, subtitle, kicker, meta, image |
| `deck.section()` | 章节过渡页 | number, title, subtitle, agenda |
| `deck.slide()` | 内容页 | title（写行动标题）, eyebrow, sub |
| `deck.closing()` | 结尾页 | title, items（写要对方决策什么，别写"谢谢"） |
| `s.kpi_row()` | KPI 大数字卡 | value, unit, label, delta, trend, note |
| `s.chart()` | 原生图表 | 见下方图表选型 |
| `s.table()` | 数据表 | headers, rows, col_widths, align, cell_colors |
| `s.bullets()` | 要点列表 | items 传 `(标题, 说明)` 得到两级结构 |
| `s.icon_cards()` | 图标卡片组 | icon（符号）, title, desc, color |
| `s.progress_bars()` | 达成率条 | label, value, target |
| `s.timeline()` | 里程碑 | when, title, desc, status(done/doing/todo) |
| `s.compare()` | 左右对比 | 改善前后、目标 vs 实际、方案 A/B |
| `s.matrix()` | 2x2 四象限 | 风险矩阵、优先级排序 |
| `s.callout()` | 结论条 | kind: accent/good/warn/bad，**每页最多一个** |
| `s.image()` | 配图 | mode='cover' 裁剪填满（默认），'fit' 完整放入 |
| `s.note()` | 数据来源 | 有数字的页必须写 |

完整参数见 `references/deckkit-api.md`。

## 换成企业品牌色

```python
from deckkit import DARK, Deck

T = DARK.variant(bg="003669", bg_alt="002A52", surface="0B4880",
                 primary="3263A7", accent="E5B620").with_readable_text()
deck = Deck(theme=T, canvas="large")     # large = 26.67x15，与现有模板同尺寸
```

**改过底色一定要接 `.with_readable_text()`**：底色一变，原本达标的语义文字色
（绿/红/灰）就可能掉到 4.5:1 以下。这类问题在显示器上不明显，投屏时直接消失。
改完还要重跑 `check_deck.py` 确认。

现有模板的实际用色见 `exec-visual-system` skill。

## 图表选型

| 要表达 | 用 | 不要用 |
|---|---|---|
| 时间趋势 | `line` / `bar` | 饼图 |
| 项目间大小比较 | `hbar`（类目名长时） | 饼图 |
| 构成占比 | `stacked` / `hstacked` | 多个饼图并排 |
| 单一构成且只有 2-4 项 | `doughnut` | 超过 5 项的饼图 |
| 多指标 vs 目标 | `progress_bars` | 饼图、雷达图 |
| 精确数值对照 | `table` | 图表 |

饼图只能表达"一个整体的构成"，不能比较、不能看趋势。超过 5 个扇区就没人读得出来。

## 语义颜色

按含义选色，不要按好看选色。传语义名即可，库会自动处理填充与文字的对比度：

| 名字 | 含义 |
|---|---|
| `accent` | 全页唯一重点（金色）。用多了就没有重点 |
| `good` / `warn` / `bad` | 达成 / 需关注 / 未达成或风险 |
| `neutral` | 中性、对照项、历史值 |
| `primary` / `secondary` | 品牌色，用于图表序列和普通色块 |

`fill=` 和 `color=` 可以传同一个名字，库分别取填充版和文字版。

## 必须避免的错误

这些都是实测踩过的，不是理论问题：

- **不要用 `font.name` 设中文字体。** python-pptx 只写 `<a:latin>`，中文会回落到主题
  字体。始终通过 deckkit 的组件或 `apply_font()` 设字体，它会同时写 latin/ea/cs。
- **不要假设文字装不下会自动缩小。** PowerPoint 会直接裁掉。长文本传 `fit=True`，
  或先用 `text_height()` 量高度再决定分几页。
- **横向条形图不要自己反转数据。** `hbar`/`hstacked` 库已经反转过，让第一项显示在
  顶部；单序列时 `colors` 也跟着反转过。你再反一次就又倒过来了。
- **堆积图的数据标签位置不能用 `outEnd`。** PowerPoint 会判定文件损坏。库已强制
  改成 `inEnd`，不要用 `label_position="outEnd"` 覆盖。
- **正文不要小于 14pt（设计单位）。** 投屏后排看不见。内容装不下就拆页或删内容，
  不要缩字号。表格（默认 12pt）、图表标签、页脚来源不受这条限制 —— 它们是辅助信息，
  9-12pt 可以接受；但表格是这一页的主体内容时，显式传 `size=14`。
- **一页只放一个 `callout`。** 全篇的强调色（`accent`）用量控制在 3 处以内 ——
  `callout(kind="accent")` 算一处，KPI 卡里 `color="accent"` 的大数字也算一处。
  某一页确实要提示风险，用 `kind="warn"` 或 `kind="bad"`，不要再叠一个 accent。
- **不要让组件默认撑满 `body`。** 先按上面的「组件占位高度」表算好再切区域。
- **深色底配图片要压一层半透明色块**，否则图上的文字读不清。用
  `s.rect(..., transparency=45)` 或 `deck.cover(image=..., scrim=52)`，
  **不透明的色块会把图整张盖住**。
- **表格数字列一律右对齐**（`align=["left","right",...]`），左对齐的数字没法比较位数。
- **改过主题底色一定要接 `.with_readable_text()`**，并重跑质检。

## 依赖

```bash
pip install python-pptx Pillow          # 生成 + 图片处理（必需）
# 渲染预览（可选，但强烈建议装，用于第 4 步视觉复核）
#   macOS   brew install --cask libreoffice && brew install poppler
#   Ubuntu  sudo apt-get install -y libreoffice-impress poppler-utils fonts-noto-cjk
```

字体：脚本写入的字体名由**打开文件的那台机器**渲染。中文默认用微软雅黑（Windows
Office 自带）。

### 看预览图时要知道的两件事

- **LibreOffice 会用本机字体替代微软雅黑**，字宽略有差异。判断是否溢出以
  `check_deck.py` 的估算为准，预览用来看版式和配色。
- **LibreOffice 不渲染堆积图的数据标签**。预览里 `stacked` / `hstacked` 的百分数
  一个都看不到，但文件本身是对的（在 PowerPoint 里能正常显示）。不要去"修"这个
  不存在的问题。

## 脚本

| 脚本 | 用途 |
|---|---|
| `python3 scripts/check_deck.py x.pptx --strict` | 交付前质检。error 必须清零；`--strict` 下 warn 也会让退出码非零，逐条看过再决定 |
| `python3 scripts/render_deck.py x.pptx out/ --grid` | 渲染逐页 PNG（统一命名 `slide-01.png`）+ 总览图 |
| `python3 scripts/selftest.py` | 改动 deckkit.py 后跑回归自测：度量函数、主题配色、图表着色、全组件 x 3 主题 x 2 画布 |
| `python3 assets/example_deck.py [输出路径]` | 11 页参照汇报（标准画布） |
| `python3 assets/example_brand.py [输出路径]` | 3 页企业品牌色示例（大画布 + 封面背景图） |

质检的分工要清楚：**`check_deck.py` 只能保证"文字读得出来"**（不溢出、不越界、
不被盖住、对比度够、字体对），**保证不了"版面好看"**。留白是否均匀、重点是否突出、
配色是否得体，只有看渲染图才知道。这就是第 4 步不能跳的原因。

## 读取或改造已有 pptx

`python-pptx` 可以读文本、改文字、加页，但有三个限制要知道：

- **不能复制幻灯片。** 只能 `add_slide(layout)`。要复用版式就基于母版版式重建。
- **`text_frame.text = "..."` 会清掉格式**（段落塌缩成一个无样式 run）。要保留格式
  就逐个改 `run.text`。
- **读不了 SVG/EMF**，模板里的矢量图标会让 `add_picture` 抛 `UnidentifiedImageError`。

要提取已有 PPT 的配色和字体作为新汇报的基线，可以读 `ppt/theme/theme*.xml` 里的
`clrScheme` 和 `fontScheme`，再用 `Theme(...)` 或 `theme.variant(...)` 建对应主题。
