# 汇报材料制作 Skill 集

给凯莱英生产管理场景准备的一套 Cursor Agent Skills，用来把汇报 PPT 和网页版汇报
做得更专业、更适合给管理层看。

装好之后**不需要记任何命令**。在 Cursor 里直接用中文说需求，相关 skill 会自动
生效，例如：

> 把这份上半年产量数据做成给管理层看的汇报 PPT

> 帮我做一页 slide，讲 3 号线的瓶颈和扩能方案

> 这份汇报做成网页版，要能全屏演示还能打印成 PDF

> 帮这页配张合适的图

## 装了什么

### 自研（针对生产管理汇报场景）

| Skill | 做什么 | 什么时候自动触发 |
|---|---|---|
| `exec-deck-builder` | 生成汇报 PPT。内置组件库，产出真正的 pptx（文字可编辑、图表是原生图表） | 提到 PPT、幻灯片、汇报、deck、演示文稿 |
| `html-exec-report` | 生成网页版汇报，可键盘翻页、全屏演示、打印 PDF、打包成单文件 | 提到 HTML、网页版、发链接 |
| `exec-visual-system` | 视觉决策依据：配色、字号层级、中文排版、版式选择 | 需要判断怎么排好看、怎么贴合企业 VI |
| `production-report-narrative` | 内容组织：汇报结构、行动标题、CDMO 指标体系、归因分析 | 需要搭汇报结构、选指标、提炼结论 |
| `deck-imagery` | 配图：18 个生产管理图标、6 种背景图案、AI 配图与合规约束 | 需要配图、加图标、页面太空 |

### Anthropic 官方（Apache-2.0）

| Skill | 做什么 |
|---|---|
| `theme-factory` | 10 套预设配色主题，也能按需生成新主题 |
| `frontend-design` | 通用视觉设计方法论，避免做出模板感的东西 |
| `internal-comms` | 内部沟通文档写作（状态汇报、管理层更新、FAQ） |

## 环境准备

只需要装 Python 依赖就能生成文件：

```bash
pip install python-pptx Pillow
```

想在生成后**看渲染效果**（强烈建议，能提前发现文字被裁、排版错位），再装：

```bash
# macOS
brew install --cask libreoffice && brew install poppler
# Ubuntu / Debian
sudo apt-get install -y libreoffice-impress poppler-utils fonts-noto-cjk
# Windows：从 libreoffice.org 下载安装，把 program 目录加入 PATH
```

网页版汇报的截图复核需要 Chrome，一般机器上已经有了。

## 快速验证

跑一遍示例，确认环境没问题：

```bash
cd .cursor/skills/exec-deck-builder
python3 assets/example_deck.py                       # 生成 11 页示例 PPT
python3 assets/example_brand.py                      # 用现有模板配色 + 26.67x15 画布
python3 scripts/check_deck.py assets/example_deck.pptx --strict   # 质检
python3 scripts/render_deck.py assets/example_deck.pptx preview/ --grid  # 渲染成图

cd ../html-exec-report
python3 assets/build_example.py                      # 生成 10 页网页版汇报
python3 scripts/shoot.py assets/example.html preview/  # 逐页截图
```

三份示例可以直接打开看效果，也可以当模板改：

| 文件 | 内容 |
|---|---|
| `exec-deck-builder/assets/example_deck.pptx` | 11 页完整汇报，覆盖全部页型（标准 16:9 画布） |
| `exec-deck-builder/assets/example_brand.pptx` | 3 页，用现有模板配色 + 26.67x15 画布 + 封面背景图 |
| `html-exec-report/assets/example.html` | 10 页网页版，与 PPT 版同一份内容 |

## 这套 skill 解决了什么

做中文汇报材料时反复踩的几个坑，已经在库层面处理掉了：

- **中文字体设了不生效**：`python-pptx` 只写西文字体标签，中文会回落到主题字体。
  库对每个文字同时写西文、东亚、复杂文本三种字体设置。
- **文字被裁掉**：PowerPoint 不会自动缩小放不下的文字，直接裁。库按全角/半角
  估算宽度、预测折行、自动降档字号，交付前还有质检工具扫一遍。
- **投屏后看不清**：语义色（绿/红/灰）当小字用时对比度不够。库把填充色和文字色
  分成两组，按 WCAG 亮度自动选可读的那个。
- **图表细节错**：横向条形图 PowerPoint 默认把第一项排在最下面；堆积图的数据标签
  位置用错会让 PowerPoint 判定文件损坏。这些都已按正确方式处理。
- **换画布就乱**：组件用设计单位定位，13.3 英寸和 26.67 英寸（现有模板尺寸）两种
  画布共用一套代码。

## 交付前必做两件事

1. **跑质检**：`python3 scripts/check_deck.py 你的文件.pptx --strict`，error 清零。
2. **渲染成图逐页看**。代码里看不出溢出和拥挤，必须看图。

## 目录结构

```
.cursor/skills/
├── exec-deck-builder/          PPT 生成
│   ├── scripts/deckkit.py          组件库
│   ├── scripts/check_deck.py       质检
│   ├── scripts/render_deck.py      渲染预览
│   ├── scripts/selftest.py         回归自测
│   ├── references/deckkit-api.md   完整 API
│   └── assets/example_deck.py      11 页参照汇报
├── html-exec-report/           网页版汇报
│   ├── assets/deck.css / deck.js   设计系统与翻页
│   ├── scripts/svg_chart.py        内联 SVG 图表
│   ├── scripts/shoot.py            逐页截图
│   └── scripts/inline.py           打包单文件
├── deck-imagery/               配图
│   ├── scripts/make_icon.py        18 个场景图标
│   └── scripts/make_bg.py          6 种背景图案
├── exec-visual-system/         视觉规范
├── production-report-narrative/ 内容结构与指标体系
├── theme-factory/              官方：配色主题
├── frontend-design/            官方：设计方法论
└── internal-comms/             官方：内部沟通写作
```

## 关于品牌色

`exec-visual-system` 里的凯莱英色值是从本仓库 `安全专篇1页.pptx` 的主题配色中
提取的（深蓝 `#003669`、强调金 `#E5B620`、微软雅黑 + Arial），是现有模板的实际用色。

**正式对外材料请向品牌/市场部门确认官方 VI 规范色值**，模板反推出的颜色不一定
等于 VI 手册标准。

## 许可

- 自研的 5 个 skill：随本仓库。
- `theme-factory` / `frontend-design` / `internal-comms` 来自
  [anthropics/skills](https://github.com/anthropics/skills)，Apache-2.0，
  各目录内保留原始 `LICENSE.txt`，未作修改。
- Anthropic 的 `pptx` / `docx` / `xlsx` / `pdf` 四个 skill 是 source-available 而非
  开源，许可明确禁止复制和再分发，**因此没有收录**。本仓库的 PPT 生成能力是基于
  开源库 `python-pptx` 独立实现的。
