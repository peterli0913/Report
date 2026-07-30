# 汇报材料制作 Skill 集

给凯莱英生产管理场景准备的一套 Cursor Agent Skills，用来把汇报 PPT 和网页版汇报
做得更专业、更适合给管理层看。

装好之后**不需要记任何命令**。在 Cursor 里直接用中文说需求，相关 skill 会自动
生效，例如：

> 把这份上半年产量数据做成给管理层看的汇报 PPT

> 帮我做一页 slide，讲 3 号线的瓶颈和扩能方案

> 这份汇报做成网页版，要能全屏演示还能打印成 PDF

## 怎么让它长期可用

skill 本身是纯文件，提交进 git 就不会丢。真正会丢的是**依赖**（python-pptx、
LibreOffice 这些），因为 Cloud Agent 的虚拟机在会话结束后会被回收，临时装的东西
不保留。所以按使用场景分三种做法：

### 场景一：在本仓库用（已经配好，无需操作）

`.cursor/environment.json` 已提交，Cloud Agent 每次新建虚拟机时会自动执行
`.cursor/setup.sh` 把依赖装齐。这个脚本是幂等的，重复运行安全，已装齐时几秒跑完。

本地第一次用，手动跑一次：

```bash
bash .cursor/setup.sh
```

### 场景二：在你的所有项目里用（推荐）

装到用户级目录，之后任何项目、任何 agent 会话都能触发：

```bash
bash .cursor/skills/install.sh              # 复制到 ~/.cursor/skills
bash .cursor/skills/install.sh --link       # 改用符号链接，本仓库改动即时生效
bash .cursor/skills/install.sh --all-tools  # 顺带让 Claude Code / Codex 也能用
bash .cursor/skills/install.sh --list       # 先看会装什么
bash .cursor/skills/install.sh --uninstall  # 移除
```

装完在 Cursor 里执行 `Developer: Reload Window`（或重启）即可生效。

### 场景三：在另一个仓库的 Cloud Agent 里用

Cursor 官方文档没有承诺 Cloud Agent 会读取虚拟机上的用户级目录，所以给 Cloud
Agent 用必须**把 skill 放进那个仓库**：

```bash
# 在目标仓库根目录执行
bash /path/to/本仓库/.cursor/skills/install.sh --target .cursor/skills
cp /path/to/本仓库/.cursor/environment.json .cursor/
cp /path/to/本仓库/.cursor/setup.sh .cursor/
git add .cursor && git commit -m "引入汇报材料 skill 集"
```

提交后该仓库的所有 Cloud Agent 都会自动装齐依赖。

### 场景四：分发给团队

仓库根已有 `.cursor-plugin/plugin.json`，可以作为 Cursor Plugin 分发：

- **团队 Marketplace**：Cursor Dashboard → Plugins → Team Marketplaces →
  Add Marketplace → Import from Repo，填本仓库地址。可设为 Default Off / Default
  On / Required。
- **本地试装**：`ln -s $(pwd) ~/.cursor/plugins/local/asymchem-report-kit`，
  然后 `Developer: Reload Window`。

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

## 环境自检

不确定装没装齐、或者哪个脚本报 ImportError：

```bash
python3 .cursor/skills/exec-deck-builder/scripts/doctor.py
# 装过之后可以直接用：deck-doctor
```

它会逐项列出就绪/缺失，并区分必需项（缺了连 pptx 都生成不了）和可选项（只影响
渲染预览、HTML 截图），缺什么给出对应平台的安装命令。

| 依赖 | 必需 | 缺了会怎样 |
|---|---|---|
| python-pptx | 是 | 无法生成 pptx |
| Pillow | 是 | 无法处理图片、生成图标与背景 |
| LibreOffice + poppler | 否 | 无法把 pptx 渲染成图做视觉复核 |
| 中文字体（Noto CJK） | 否 | 渲染预览里的中文显示成方块 |
| Chrome / Chromium | 否 | 无法给 HTML 汇报逐页截图 |

## 快速验证

下面的命令都输出到 `/tmp/demo`，不会改动仓库里的文件：

```bash
D=/tmp/demo && mkdir -p $D
P=.cursor/skills/exec-deck-builder
H=.cursor/skills/html-exec-report

python3 $P/assets/example_deck.py  $D/示例.pptx        # 11 页示例 PPT
python3 $P/assets/example_brand.py $D/品牌色.pptx      # 现有模板配色 + 26.67x15 画布
python3 $P/scripts/check_deck.py   $D/示例.pptx --strict        # 质检
python3 $P/scripts/render_deck.py  $D/示例.pptx $D/preview --grid   # 渲染成图

python3 $H/assets/build_example.py $D/网页版.html      # 10 页网页版汇报
python3 $H/scripts/shoot.py        $D/网页版.html $D/web-preview   # 逐页截图
python3 $H/scripts/inline.py       $D/网页版.html $D/单文件.html    # 打包单文件
```

仓库里的成品可以直接打开看效果，也可以当模板改：

| 文件 | 内容 |
|---|---|
| `安全专篇/安全生产专题汇报.pptx` | 9 页安全专题汇报，沿用凯莱英模板 |
| `exec-deck-builder/assets/example_deck.pptx` | 11 页完整汇报，覆盖全部页型（标准 16:9） |
| `exec-deck-builder/assets/example_brand.pptx` | 3 页，现有模板配色 + 26.67x15 画布 + 封面背景图 |
| `html-exec-report/assets/example.html` | 10 页网页版，与 PPT 版同一份内容 |

## 这套 skill 解决了什么

做中文汇报材料时反复踩的几个坑，已经在库层面处理掉了：

- **中文字体设了不生效**：`python-pptx` 只写西文字体标签，中文会回落到主题字体。
  库对每个文字同时写西文、东亚、复杂文本三种字体设置。
- **文字被裁掉**：PowerPoint 不会自动缩小放不下的文字，直接裁。库按全角/半角
  估算宽度、预测折行、自动降档字号，交付前还有质检工具扫一遍。
- **投屏后看不清**：语义色（绿/红/灰）当小字用时对比度不够。库把填充色和文字色
  分成两组，按 WCAG 亮度自动选可读的那个。
- **换品牌色后语义失效**：只按对比度提亮会把绿橙红一起洗成难分辨的淡彩。
  提亮时同步补偿饱和度，`audit_theme()` 还会检查语义色之间的可分辨度。
- **图表细节错**：横向条形图 PowerPoint 默认把第一项排在最下面；堆积图的数据标签
  位置用错会让 PowerPoint 判定文件损坏；单序列传多个颜色会全渲染成第一个。
- **换画布就乱**：组件用设计单位定位，13.3 英寸和 26.67 英寸（现有模板尺寸）两种
  画布共用一套代码。
- **套模板时标题压在 logo 上**：`margin_top` / `margin_bottom` 分向边距让组件自动
  避开版式里固定的页眉页脚。

## 交付前必做两件事

1. **跑质检**：`deck-check 你的文件.pptx --strict`，error 清零。
2. **渲染成图逐页看**。代码里看不出溢出和拥挤，必须看图。

## 目录结构

```
.cursor/
├── environment.json            Cloud Agent 环境配置（自动装依赖）
├── setup.sh                    幂等的依赖安装脚本
└── skills/
    ├── install.sh                  装到用户级目录，跨项目复用
    ├── exec-deck-builder/          PPT 生成
    │   ├── pyproject.toml              可 pip 安装，import 不依赖工作目录
    │   ├── scripts/deckkit.py          组件库
    │   ├── scripts/doctor.py           环境体检
    │   ├── scripts/check_deck.py       质检
    │   ├── scripts/render_deck.py      渲染预览
    │   ├── scripts/selftest.py         回归自测
    │   ├── references/deckkit-api.md   完整 API
    │   └── assets/example_deck.py      11 页参照汇报
    ├── html-exec-report/           网页版汇报
    ├── deck-imagery/               配图
    ├── exec-visual-system/         视觉规范
    ├── production-report-narrative/ 内容结构与指标体系
    ├── theme-factory/              官方：配色主题
    ├── frontend-design/            官方：设计方法论
    └── internal-comms/             官方：内部沟通写作
.cursor-plugin/plugin.json       Plugin 清单，用于团队 Marketplace 分发
安全专篇/                        安全生产专题汇报（成品 + 生成脚本）
```

## 关于品牌色

`exec-visual-system` 里的凯莱英色值是从 `安全专篇1页.pptx` 的主题配色中提取的
（深蓝 `#003669`、强调金 `#E5B620`、微软雅黑 + Arial），是现有模板的实际用色。

**正式对外材料请向品牌/市场部门确认官方 VI 规范色值**，模板反推出的颜色不一定
等于 VI 手册标准。同理，`production-report-narrative` 里的指标体系是行业通用口径，
具体定义和目标值需按公司实际管理体系核对。

## 许可

- 自研的 5 个 skill：随本仓库。
- `theme-factory` / `frontend-design` / `internal-comms` 来自
  [anthropics/skills](https://github.com/anthropics/skills)，Apache-2.0，
  各目录内保留原始 `LICENSE.txt`，未作修改。
- Anthropic 的 `pptx` / `docx` / `xlsx` / `pdf` 四个 skill 是 source-available 而非
  开源，许可明确禁止复制和再分发，**因此没有收录**。本仓库的 PPT 生成能力是基于
  开源库 `python-pptx` 独立实现的。
