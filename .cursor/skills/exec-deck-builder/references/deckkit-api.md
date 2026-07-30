# deckkit API 参考

按需查阅。常用写法在 `../SKILL.md`，可运行的完整示例在 `../assets/example_deck.py`。

坐标与尺寸一律是**设计英寸**（基准画布 13.333 x 7.5），字号是**设计 pt**。
换画布只改 `Deck(canvas=...)`，其余代码不动。

---

## Deck

```python
Deck(theme="dark", canvas="wide", template=None, layout=None, keep_slides=False)
```

| 参数 | 说明 |
|---|---|
| `theme` | `"dark"` / `"light"` / `"slate"`，或一个 `Theme` 实例 |
| `canvas` | `"wide"` 13.333x7.5（默认，优先选这个）· `"large"` 26.67x15（贴合公司现有模板）· `"a4"` 11.69x8.27 · 或 `(宽, 高)` |
| `template` | 传入 `.pptx` 路径则沿用该文件的母版、版式与画布尺寸 |
| `layout` | 默认版式名。套模板时指定带页眉 logo / 页脚的那个 |
| `keep_slides` | 默认 `False`，即清空模板自带的幻灯片只留母版和版式 |

属性：`deck.theme`、`deck.scale`（设计单位到实际画布的倍数）、`deck.slides`、
`deck.canvas_w` / `deck.canvas_h`（实际画布英寸）、`deck.prs`（原始 `Presentation`）。

套模板时的行为：画布尺寸以模板为准（改了背景图会被拉变形）；版式带来的空占位符
会被删掉（否则「单击此处编辑标题」跟着导出）；`bg=None` 时不填底色，版式的背景图
才不会被盖掉。

### 整页方法

```python
deck.cover(title, subtitle=None, meta=None, kicker=None, image=None,
           accent_block=True, scrim=52, layout=None, bg=None)
deck.section(number, title, subtitle=None, agenda=None, layout=None, bg=None)
deck.slide(title=None, eyebrow=None, sub=None, bg=None, rule=False,
           layout=None) -> Slide
deck.closing(title="讨论与决策事项", items=None, meta=None, layout=None, bg=None)
deck.raw_slide(layout=None, keep_placeholders=False) -> Slide
deck.layout(name)          # 按名称取版式，支持模糊匹配
deck.list_layouts()        # 列出模板全部版式名
deck.add_page_numbers(skip_first=True, total=True)
deck.save(path)
```

`raw_slide()` 建一页但不设背景、不加标题，用于封面这类要完全自己排的页面 ——
套模板时它能保住版式自带的背景图。每页都可以用 `layout=` 单独指定版式（封面用
一个、内容页用另一个）。

- `cover(image=..., scrim=52)` 在图上压一层 52% 的半透明底色保证标题可读，图案本身
  透出来。照片类背景通常要调到 60-70。**不要自己用不透明矩形压图**，那等于没放图。
- `section(agenda=[...])` 在页脚列出全部章节并高亮当前章，听众才知道讲到哪了。
- `closing(items=[...])` 写**要对方决策什么**，不要写"谢谢聆听"。

---

## Rect：区域切分

```python
Rect(x, y, w, h)
```

属性：`right` `bottom` `cx` `cy`

```python
r.inset(d, dy=None)                     # 四周内缩（dy 单独控制上下）
r.pad(top=0, right=0, bottom=0, left=0) # 只缩指定边
r.split_h(*ratios, gap=0)               # 横向按比例切，返回 list[Rect]
r.split_v(*ratios, gap=0)               # 纵向按比例切
r.grid(cols=1, rows=1, gap=0.24, gap_y=None)  # 行优先返回 cols*rows 块
```

比例自动归一化，`split_h(0.55, 0.45)` 和 `split_h(55, 45)` 等价。

Slide 上的现成区域：`s.page`（整页）、`s.safe`（去页边距）、`s.body`（标题以下、脚注以上，加过标题后自动下移）。

---

## Slide 组件

### kpi_row —— KPI 大数字卡

```python
s.kpi_row(items, at=None, cols=None, gap=None, card=True, height=None)
```

`items` 每项：

| 键 | 说明 |
|---|---|
| `value` | 大数字，字符串（自己格式化好千分位） |
| `unit` | 单位，小字跟在数字后 |
| `label` | 指标名 |
| `delta` | 同比/环比文字，如 `"+23.1%"` |
| `trend` | `up` / `down` / `flat`，决定箭头方向和默认颜色 |
| `delta_color` | 覆盖 delta 颜色。**下降是好事时必须传**（如能耗、偏差数传 `"good"`） |
| `color` | 大数字颜色，突出某一个指标时传 `"accent"` |
| `note` | 卡片底部补充小字 |

不传 `height` 时卡片高度**固定约 1.95"**，不会撑满 `body`。切区域时按这个数算，
否则下方会留出一条空白带。

### chart —— 原生图表

```python
s.chart(kind, categories, series, at=None, legend=None, data_labels=None,
        number_format=None, colors=None, gap_width=60, overlap=None, smooth=False,
        value_axis=True, category_axis=True, max_scale=None, min_scale=None,
        label_position=None, title=None, target=None, target_label="目标")
```

`series` 是 `[(名称, [数值...]), ...]`。

`kind` 取值：`bar` `hbar` `line` `line_markers` `area` `stacked` `hstacked`
`stacked100` `hstacked100` `pie` `doughnut` `radar`

库已处理好的默认行为，不要重复处理：

- `hbar` / `hstacked` **已反转类别顺序**，第一项显示在顶部；单序列时 `colors` 也
  跟着反转，所以你传入的颜色顺序始终对应你传入的类别顺序。
- **但序列的上下顺序没反转，也不该反转。** 横向条形图里 PowerPoint 把第一个序列画在
  每组的**下方**（这是它的固有约定，图例顺序也跟着）。想让某个序列显示在上面，
  把它放在 `series` 列表的后面。类别会反转、序列不会，这两条放在一起容易误判。
- 单序列自动隐藏图例，多序列自动显示在顶部。
- 柱/条/饼默认开数据标签，折线/面积默认关（折线开了容易糊）。
- 堆积图标签强制 `inEnd`（`outEnd` 会让 PowerPoint 报文件损坏），并按每段填充色
  自动选标签文字色（压在金色段上的白字读不出来）。
- 各类别合计为 100 的堆积图自动把值轴上限设为 100。
- 图表区与绘图区透明，融进页面底色。

`number_format` 用 Excel 格式串：`"0"` `"0.0"` `'0"%"'` `"#,##0"` `'0.0"吨"'`。

**`colors` 的含义随序列数变化**：

- 多序列时按**序列**取色：`colors[i]` 是第 i 个序列的颜色。
- 单序列的柱/条形图按**数据点**取色：`colors[j]` 是第 j 根柱子的颜色。这是特意做的
  —— PowerPoint 的着色单位是序列，只有一个序列时按序列上色会让所有柱子同色，
  传四个颜色却全渲染成第一个。
- 折线、面积、雷达图始终按序列取色（按点着色对折线没有意义）。

传语义名或 hex 都可以：`colors=["bad", "warn", "neutral", "hairline"]`。
（避免用 `hairline` 当数据颜色，它在深底上几乎看不见，用 `neutral`。）

**目标参考线**：`target=95` 给折线/面积/雷达图加一条灰色虚线，图例里显示为
`target_label`。柱状图和条形图不支持 —— PowerPoint 的组合图 python-pptx 画不出来，
传了会被静默忽略。柱图要表达目标，改用 `progress_bars`，或者把目标值写进标题和
结论条。

```python
s.chart("line", months, [("闭环率", vals)], at=s.body,
        target=95, target_label="目标 95%", min_scale=88, max_scale=98,
        number_format='0.0"%"')
```

纵轴截断（`min_scale`）在数据集中于高位时很有用（91→96 在 0-100 轴上是一条平线），
但刻度必须标出来，否则会夸大变化幅度。

### table —— 数据表

```python
s.table(headers, rows, at=None, col_widths=None, align=None, row_h=None,
        header_h=None, cell_colors=None, size=None, zebra=True)
```

- `col_widths=[2.3, 1, 1, 1]` 相对比例，自动归一化。
- `align=["left", "right", "right", "center"]` —— **数字列一律 `right`**。
- `cell_colors={(行, 列): "good"}` 给单元格文字染色，行列都从 0 起（不含表头）。
- 表头固定用 `surface_alt` 底色 + 粗体，已关掉 PowerPoint 自带的花哨样式。
- **表格不会撑满区域**：自动行高上限是 `0.40`，所以 4 行表最多约 1.9" 高，区域给
  大了只会在下方留白。要占满就显式传 `row_h` / `header_h`，并按
  `header_h + row_h × 行数` 反算区域高度。
- 字号默认 `size_small`（12pt）。表格是这一页的主体内容时传 `size=14`。

### bullets —— 要点列表

```python
s.bullets(items, at=None, size=None, marker="dot", gap=0.14, color="ink",
          bold_head=True, marker_color="accent", desc_color="ink_muted")
```

`items` 元素可以是 `"文字"`，也可以是 `("结论", "解释说明")` —— 后者形成"粗体结论 +
灰色解释"的两级结构，比一串平铺的 bullet 好读得多。

`marker`：`"dot"` 圆点 / `"num"` 两位序号 / `"none"` 无标记。

返回排完后的 y 坐标，可以接着往下放东西。

### icon_cards —— 图标卡片组

```python
s.icon_cards(items, at=None, cols=3, gap=None, icon_size=0.46)
```

`items` 每项 `{icon, title, desc, color}`。`icon` 用文字符号（`✓ ! → ● ▲ ■ 1 2 3`），
画在色圆里，不依赖外部图片文件，换机器不会缺图。要用真图标见 `deck-imagery` skill。

### progress_bars —— 达成率条

```python
s.progress_bars(items, at=None, size=None, bar_h=0.20, gap=0.30,
                show_target=True, target=100.0, align="middle")
```

`items` 每项 `{label, value, text?, color?}`。`value` 传数值（可超过 100）。

条长按 `max(target, 所有 value) × 1.04` 缩放，并在 `target` 处画虚线参考线 —— 不硬
截断到 100，否则 103% 和 108% 会画成一样长。轨道末端因此总会露出一小段底色，
这是刻意留的余量，不是画错了。不传 `color` 时按 `value` 与 `target` 的关系自动取
good（达标）/ warn（差 5% 以内）/ bad。

**值域集中在 90-100 时这个组件会失去分辨力**：四个值都在 94-98、缩放上限约 101，
条长差不到 8%，肉眼分不出来，"谁没达标"只能靠颜色和参考线位置读。这种情况要么
接受"靠颜色读"，要么改用表格把数值和缺口直接列出来。

`target_label` 覆盖默认的「目标 X%」文字，用于非百分比指标（默认写死了百分号）。

`align`：`middle` 垂直居中 / `top` 贴顶 / `fill` 撑满区域。开 `show_target` 时上方会
占用 `0.28"` 放目标标签，虚线和标签跟着条形组走，不会被第一根条压住。

占用高度：`n × (bar_h + gap) − gap`，加上目标标签的 0.28"。

### timeline —— 里程碑

```python
s.timeline(items, at=None, orient="h", size=None)
```

`items` 每项 `{when, title, desc?, status?}`。`status`：`done` 绿 / `doing` 金 / `todo` 灰。
`orient="v"` 竖向排列，适合放在窄区域；横竖两种方向都会渲染 `desc`。

### compare —— 左右对比

```python
s.compare(left, right, at=None, gap=None)
```

`left` / `right` 都是 `{title, items[], color?, note?}`。用于改善前后、目标 vs 实际、
方案 A vs 方案 B。标题栏用 `color` 填充，文字色自动取对比色。

### matrix —— 2x2 四象限

```python
s.matrix(quadrants, at=None, x_label=None, y_label=None, gap=0.18)
```

`quadrants` 按 **[左上, 右上, 左下, 右下]** 顺序，每项 `{title, items[]?, desc?, color?}`。

`x_label` / `y_label` 传字符串居中显示，传 `(低端, 高端)` 二元组则摆到轴两端并自动
加箭头。纵轴标签用真正的竖排文字方向，不是靠窄文本框逐字折行。

### callout —— 结论条

```python
s.callout(text, at=None, kind="accent", label=None, size=None, icon=None)
```

`kind`：`accent`（结论）/ `good` / `warn` / `bad` / `primary` / `surface`（中性）。
不传 `at` 时贴在页面底部通栏，占用约 `0.86"` 高（切区域时要给它让出这段 + 间距）。
**每页最多一个**，全篇 `accent` 用量控制在 3 处以内。

### image —— 配图

```python
s.image(path, at=None, mode="cover", radius=None, caption=None)
```

`mode="cover"` 按区域比例裁剪填满（不变形，默认）；`"fit"` 完整放入并居中留白。
变形的照片在汇报里非常显眼，所以默认走裁剪而不是拉伸。

### 基础图元

```python
s.rect(at, fill=None, line=None, line_w=1.0, radius=None, shadow=False,
       shape=None, transparency=0)
s.line(x1, y1, x2, y2, color="hairline", width=1.0, dash=None)
s.text(at, content, size=None, bold=False, color="ink", align="left", anchor="top",
       line_spacing=1.28, italic=False, font_en=None, font_cn=None, fit=False,
       min_pt=8.0, space_after=0.0, shrink_wrap=False, vert=None)
s.sparkline(values, at, color="accent", width=1.8)
s.set_bg(color)
s.bg_image(path)
s.title(text, eyebrow=None, sub=None, color="ink", rule=False)
s.note(text)
s.page_number(n, total=None)
```

`s.text()` 的 `content` 可以传富文本片段列表，在一行里混排不同字号/颜色：

```python
s.text(area, [("1,284", {"size": 40, "bold": True, "color": "accent"}),
              (" 吨", {"size": 16, "color": "ink_muted"})])
```

- `fit=True`：按区域高度自动降档字号（防裁字的主要手段）。
- `anchor`：`top` / `middle` / `bottom`。
- `vert="vert270"`：竖排，从下往上读，用于纵轴标签。
- `align`：`left` / `center` / `right` / `justify`。**正文永远左对齐**，只有标题居中。

---

## Theme

```python
Theme(name=..., bg=..., ink=..., primary=..., size_body=14, margin=0.62, ...)
theme.variant(accent="FF6B00", font_cn="思源黑体")   # 基于现有主题改几个字段
theme.with_readable_text()                          # 改过底色后重算语义文字色
```

**改过 `bg` / `surface` 之后一定要接 `.with_readable_text()`**：底色一变，原本达标的
文字色就可能掉到 4.5:1 以下，而这类问题在显示器上不明显、投屏时却直接消失。

```python
T = DARK.variant(bg="003669", surface="0B4880",
                 accent="E5B620").with_readable_text()
deck = Deck(theme=T, canvas="large")
```

### 颜色语义

| 分组 | 字段 |
|---|---|
| 底色层次 | `bg` `bg_alt` `surface` `surface_alt` `hairline` |
| 文字 | `ink` `ink_muted` `ink_on_accent` |
| 品牌 | `primary` `secondary` `accent` |
| 语义填充 | `good` `warn` `bad` `neutral` |
| 语义文字 | `good_text` `warn_text` `bad_text` `neutral_text` `accent_text` `primary_text` `secondary_text` |
| 图表 | `series`（按序列顺序取用的元组） |

**为什么填充色和文字色要分两组**：饱和的绿做色块很好看，但当成 12pt 小字压在深底上
达不到 WCAG 4.5:1。传语义名时库自动取对应版本 —— `fill="good"` 取填充版，
`color="good"` 取文字版，调用方不用记两套名字。

### 字号阶梯（设计 pt）

`size_cover_title` 40 · `size_cover_sub` 18 · `size_section_title` 34 ·
`size_title` 23 · `size_kpi` 40 · `size_h` 16 · `size_body` 14 · `size_small` 12 ·
`size_eyebrow` 11.5 · `size_kpi_label` 11.5 · `size_note` 9.5

**正文 14pt 是投屏下限**，再小后排看不清。表格默认走 `size_small`（12pt）、脚注
9.5pt —— 辅助信息可以到 9-12pt，但表格作为主体内容时显式传 `size=14`。

### 版面

`margin` 0.62（左右页边距）· `gap` 0.24（元素间距）· `radius` 0.045（圆角占短边比例）

`margin_top` / `margin_bottom` 默认跟随 `margin`，套企业模板时单独放大以让开版式里
固定的页眉 logo 和页脚色条。只读属性 `theme.mt` / `theme.mb` 取生效值，
`title()`、`body`、`note()`、`page_number()` 都按它们定位。

```python
T = LIGHT.variant(margin=0.58, margin_top=1.30, margin_bottom=0.52)
```

`body` 的下沿会再让出 0.36 英寸给 `note()` / `page_number()`，所以内容不会和页脚叠。

### 辅助方法

```python
theme.color(key)             # 语义名或 hex -> hex
theme.text_color(key)        # 语义名 -> 文字版 hex
theme.luminance(key)         # WCAG 相对亮度 0-1
theme.ink_on(bg)             # 返回压在 bg 上对比度更高的文字色
theme.block_ink(fill)        # 返回 (主文字色, 次文字色)
theme.variant(**kw)          # 派生新主题
theme.with_readable_text()   # 按当前底色重算全部语义文字色
```

---

## 模块级工具

```python
text_width(text, size_pt)                    # 单行宽度（英寸），中文按全角算
wrapped_lines(text, box_w, size_pt)          # 预测折行数
text_height(text, box_w, size_pt, ls=1.28)   # 预测文本块高度（英寸）
fit_size(text, box_w, box_h, size_pt, min_pt=8.0)   # 能装下的最大字号
check_overflow(path, verbose=True)           # 快速自查，返回问题列表
audit_theme(theme, min_ratio=4.5)            # 检查主题配色，返回问题列表

apply_font(font, theme, size=None, bold=None, italic=None, color=None,
           font_en=None, font_cn=None)       # 同时写 latin/ea/cs
add_shadow(shape, blur=10.0, dist=4.0, direction=5400000, color="000000", alpha=22)
clear_shadow(shape)
set_fill_alpha(shape, transparency)          # 给纯色填充加透明度（0-100）
```

`set_fill_alpha` 用于压在背景图上的遮罩。`s.rect(..., transparency=45)` 已经调它，
只有手写形状时才需要直接用 —— 不透明的遮罩会把背景图整张盖住，等于没放图。

排版前先量、后放，是避免文字被裁的根本办法：

```python
h = text_height(long_text, area.w, T.size_body)
if h > area.h:
    s.text(area, long_text, fit=True)     # 或者拆到下一页
else:
    s.text(Rect(area.x, area.y, area.w, h), long_text)
```

---

## 直接操作底层

`deck.prs` 是原始的 python-pptx `Presentation`，`s.raw` 是原始 `Slide`。
库没覆盖的需求可以直接用 python-pptx API，但记住：**自己加的文本一定要用
`apply_font()` 设字体**，否则中文会回落到主题字体。
