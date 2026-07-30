---
name: deck-imagery
description: 当汇报材料需要配图时使用 —— 给 PPT 或 HTML 汇报页配插图、封面背景图、章节配图、图标，或者用户提出"这页太空了加点图""帮我配张图""插入合适的图片""做个示意图""加个图标"。也用于判断某张图该不该用、怎么裁剪才不变形、深色底上的图怎么处理才不影响文字可读性，以及医药化工等强合规行业用图的边界。
---

# 汇报配图

配图的判断顺序：**先问这一页缺不缺视觉锚点，再问该用哪类图。** 为了填空而加图，
只会让页面更乱。

## 先判断要不要图

| 页面状态 | 处理 |
|---|---|
| 有图表 / KPI 大数字 / 表格 | 已经有视觉锚点，**不要再加装饰图** |
| 只有文字和 bullet | 需要视觉元素，但优先加图标卡或结论色块，其次才是插图 |
| 封面、章节过渡页 | 适合背景图案，撑起质感 |
| 讲设备、厂区、工艺流程 | 适合示意图或真实照片 |
| 讲趋势、对比、构成 | 用图表，不要用插图 |

一句话：**数据用图表，概念用图标，氛围用背景，实物用照片。**

## 四种图片来源

按可靠性和合规性排序：

### 1. 内置图标（首选，零依赖离线可用）

18 个生产管理场景的线性图标，颜色跟主题走：

```bash
python3 scripts/make_icon.py list                    # 看全部图标名
python3 scripts/make_icon.py 质量 icon.png --color E5B620 --size 256
python3 scripts/make_icon.py all icons/ --color F2F5F9
```

图标名支持中文别名：产量 产能 质量 安全 设备 人员 成本 交付 风险 改善 目标
周期 完成 关注 工艺 合规 审计 检查 厂区 产线 流程

```python
from make_icon import draw_icon, ICONS, ALIASES
draw_icon(name, path, color="F2F5F9", size=256, stroke=None, padding=0.10)

draw_icon("安全", "icons/safety.png", color="E5B620", size=256)
s.image("icons/safety.png", at=Rect(1.0, 2.0, 0.5, 0.5), mode="fit")
```

**只要三五个概念图标，别用文件。** `deckkit` 的 `icon_cards()` 用文字符号
（`✓ ! → ● ▲`）画在色圆里就够了，不产生外部文件依赖，换机器不会缺图。

### 2. 程序生成的背景与底纹（封面、章节页首选）

```bash
python3 scripts/make_bg.py gradient bg.png --from 0A1F3C --to 1E6FBF --angle 135
python3 scripts/make_bg.py glow    bg.png --base 0A1F3C --accent 1E6FBF
python3 scripts/make_bg.py mesh    bg.png --base 0A1F3C --line 2C4E7A
python3 scripts/make_bg.py dots    bg.png --base 0A1F3C --accent 2C4E7A
python3 scripts/make_bg.py molecule bg.png --base 0A1F3C --accent 2AA9DB
python3 scripts/make_bg.py scrim out.png --image 照片.jpg --opacity 55
```

| 模式 | 观感 | 适合 |
|---|---|---|
| `gradient` | 平滑渐变 | 通用封面 |
| `glow` | 一处柔和亮区 | 封面，文字放亮区对侧 |
| `mesh` | 细网格 | 技术、工程类专题 |
| `dots` | 极轻点阵 | 需要质感但不能抢戏的内容页 |
| `molecule` | 节点连线网络 | 医药化工类封面，中心已留空放标题 |
| `scrim` | 给照片压半透明遮罩 | 照片上要放文字时**必须**先做这一步 |

Python 调用的完整签名（`size` 都是 `(宽, 高)` 像素，`opacity` 是 0-100）：

```python
gradient(path, c_from, c_to, angle=135, size=(1920, 1080))
glow(path, base, accent, size=..., cx=0.72, cy=0.28, radius=0.85, opacity=55)
mesh(path, base, line, size=..., spacing=90, opacity=40, width=1)
dots(path, base, accent, size=..., spacing=46, radius=3, opacity=62)
molecule(path, base, accent, size=..., nodes=26, opacity=42, seed=7, link_dist=0.30)
scrim(path, image, base="000000", opacity=55, size=None)
```

颜色参数传 6 位 hex（带不带 `#` 都行）。`molecule` 的节点用分层采样铺满四周、
中间留空，所以标题压在中央不会被线条穿过；换 `seed` 得到不同图案。

### 3. AI 生成插图（用于抽象概念，不用于表示实物）

需要一张具体插图时可以生成。写 prompt 抓住四点：

- **抽象、概念化**，不要写实照片风格
- **限定配色**，跟汇报主色一致（如"deep navy blue and gold, dark background"）
- **明确不要文字**："no text, no words, no labels" —— 生成的文字几乎都是乱码
- **留出放字的空间**："negative space on the left", "centered composition with empty margins"

可用的 prompt 模板：

```
Abstract minimal illustration of <概念>, geometric line art style,
deep navy blue (#0A1F3C) background with gold (#E5B620) accents,
generous negative space on the left, no text, no words, no people,
flat vector aesthetic, professional corporate presentation background
```

生成后插入：

```python
s.image("生成的图.png", at=s.body.split_h(0.55, 0.45)[1], mode="cover")
```

### 4. 企业真实照片（讲实物时最有说服力）

厂区、产线、设备、团队的真实照片胜过任何插图。但要向用户索取，**不要用 AI 生成
的图片假装是实际设施**（见下方合规约束）。

不要从网上搜图：版权状态不明，而且很可能带着其他公司的设备标识或水印。

## 医药化工行业的用图约束

这一节是硬约束，不是建议。制药 CDMO 的材料经常进入审计、客户尽调、监管检查的
视野，图片用错的代价远高于做得好看的收益。

| 禁止 | 原因 |
|---|---|
| 用 AI 生成或网络下载的图片表示实际厂房、产线、设备、洁净区 | GMP 语境下等同于虚假陈述，审计和客户尽调时是严重问题 |
| 生成或引用具体分子结构、合成路线示意 | 可能涉及客户产品机密，也可能画错误导 |
| 在材料中出现客户名称、产品代号、专利信息 | 保密协议范围，对内汇报也应脱敏为"A 客户""项目 M" |
| 使用可辨识的员工人脸照片而未取得同意 | 肖像权 |
| 用图片美化未达标的指标 | 汇报诚信问题 |

安全做法：

- AI 生成图**只用于抽象装饰**（背景图案、概念示意），不承载事实陈述。
- 需要展示实际设施时，向用户索取企业自有照片，并确认可用范围（内部/对外）。
- 涉及客户的内容一律代号化，图片里的标识、批号、屏幕内容做遮挡处理。
- 材料页脚标注密级（如"内部资料 请勿外传"）。

## 插入的技术要点

- **按比例裁剪，绝不拉伸。** `deckkit` 的 `s.image(mode="cover")` 默认按区域比例
  裁剪填满，不变形。`mode="fit"` 完整放入并居中留白。变形的照片在汇报里极其显眼。
- **深色底上的照片必须压遮罩。** 照片的高光区会把白色标题吃掉，先用
  `make_bg.py scrim` 处理，或者用 `deck.cover(image=...)`（已内置压色）。
- **分辨率按输出尺寸的 1.5 倍准备。** 13.333 英寸宽的画布满幅图给 1920px 宽足够；
  26.67 英寸画布给 3000px 以上。图片过小放大会糊，过大只是让文件变臃肿。
- **满幅图不要压在文字下面。** 用半出血：图占一侧 40-50%，文字在另一侧。
  `s.body.split_h(0.5, 0.5)` 拿两块区域。
- **图片一律加说明。** `s.image(..., caption="3 号线结晶工序改造后现场")`，
  否则听众会猜这张图想说明什么。
- **控制文件体积。** 一份汇报里超过 8 张满幅照片就该考虑压缩，PPT 超过 20MB
  在 OA 系统里经常发不出去。

## 图标与色块的配色

图标压在色块上时，图标颜色必须和块底有足够对比。`deckkit` 的 `icon_cards()` 会自动
按块底亮度选图标色。自己画时用 `theme.ink_on(块底色)` 取颜色，不要凭感觉给白色 ——
白色图标压在金色圆上几乎看不见。

## 常见错误

| 错误 | 后果 |
|---|---|
| 已有图表的页面再加装饰插图 | 两个视觉焦点互相打架 |
| 用高对比、有明显主体的图当背景 | 压在上面的文字读不清 |
| 图片拉伸填充 | 人脸和设备变形，非常显眼 |
| 每页都配一张图 | 材料变成图册，结论被淹没 |
| AI 生成的图里带乱码文字 | 一眼就看出是 AI 生成的 |
| 图标风格混用（线性 + 立体 + emoji） | 显得拼凑 |
| 用 emoji 当正式汇报的图标 | 不同系统渲染成不同样式，且偏随意 |
