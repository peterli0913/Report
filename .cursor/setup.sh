#!/usr/bin/env bash
# 汇报材料 skill 集的依赖安装脚本。
#
# 由 .cursor/environment.json 的 install 字段在每次新建 Cloud Agent VM 时调用，
# 也可以手动运行：bash .cursor/setup.sh
#
# 必须幂等：Cursor 会在每次 VM 启动、以及快照失效回退到默认镜像时重复执行它。
# 所以每一步都先检测再安装，已装齐时几秒内跑完。
#
# 用法：
#   bash .cursor/setup.sh            # 装齐全部（含渲染预览工具，推荐）
#   bash .cursor/setup.sh --minimal  # 只装生成 pptx 所需的 Python 依赖
#   bash .cursor/setup.sh --check    # 只检查不安装

set -uo pipefail

MODE="full"
for a in "$@"; do
  case "$a" in
    --minimal) MODE="minimal" ;;
    --check)   MODE="check" ;;
  esac
done

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SKILLS="$REPO/.cursor/skills"
log() { printf '  %s\n' "$*"; }

have()     { command -v "$1" >/dev/null 2>&1; }
have_py()  { python3 -c "import $1" >/dev/null 2>&1; }

# apt 在容器里可能没有 sudo；两种情况都要能跑
if [ "$(id -u)" = "0" ]; then SUDO=""; elif have sudo; then SUDO="sudo"; else SUDO=""; fi

PIP_FLAGS=""
# Ubuntu 24.04 起 pip 默认拒绝写系统环境，需要显式放行。
# 这里刻意不用 `pip --help | grep -q`：grep -q 命中后立刻退出会让 pip 收到
# SIGPIPE，配合 pipefail 就变成"管道失败"，标志位会被误判成不支持。
PIP_HELP="$(python3 -m pip install --help 2>/dev/null || true)"
case "$PIP_HELP" in
  *break-system-packages*) PIP_FLAGS="--break-system-packages" ;;
esac

echo "==> 检查 Python 依赖"
PY_MISSING=()
have_py pptx || PY_MISSING+=("python-pptx")
have_py PIL  || PY_MISSING+=("Pillow")

if [ ${#PY_MISSING[@]} -eq 0 ]; then
  log "python-pptx、Pillow 已就绪"
elif [ "$MODE" = "check" ]; then
  log "缺少：${PY_MISSING[*]}"
else
  log "安装 ${PY_MISSING[*]}"
  python3 -m pip install --quiet --upgrade $PIP_FLAGS "${PY_MISSING[@]}" \
    || python3 -m pip install --quiet --upgrade "${PY_MISSING[@]}" \
    || { echo "  Python 依赖安装失败，请手动执行：pip install python-pptx Pillow"; exit 1; }
fi

# 把 deckkit 装成可导入的模块，这样脚本里 import deckkit 不依赖当前工作目录
echo "==> 注册 deckkit 模块"
if [ -f "$SKILLS/exec-deck-builder/pyproject.toml" ]; then
  if [ "$MODE" = "check" ]; then
    have_py deckkit && log "deckkit 可直接 import" || log "deckkit 尚未注册"
  else
    python3 -m pip install --quiet -e "$SKILLS/exec-deck-builder" $PIP_FLAGS 2>/dev/null \
      || python3 -m pip install --quiet -e "$SKILLS/exec-deck-builder" 2>/dev/null \
      || log "可编辑安装失败，脚本会退回按路径查找（功能不受影响）"
    have_py deckkit && log "import deckkit 已可用" || log "未注册，脚本将按路径查找"
  fi
else
  log "跳过（未找到 exec-deck-builder/pyproject.toml）"
fi

if [ "$MODE" = "minimal" ]; then
  echo "==> 已按 --minimal 完成：可以生成 pptx / html，但不能渲染预览"
  exit 0
fi

echo "==> 检查渲染预览工具"
# LibreOffice 用于把 pptx 转 PDF，pdftoppm 再转 PNG，中文字体保证预览不出方块
APT_MISSING=()
have soffice   || APT_MISSING+=("libreoffice-impress")
have pdftoppm  || APT_MISSING+=("poppler-utils")
# 同样避开管道：把结果存进变量再判断，绕开 SIGPIPE + pipefail 的误判
if have fc-list; then
  CJK_FONTS="$(fc-list :lang=zh 2>/dev/null || true)"
  [ -n "$CJK_FONTS" ] || APT_MISSING+=("fonts-noto-cjk")
fi

if [ ${#APT_MISSING[@]} -eq 0 ]; then
  log "LibreOffice、poppler、中文字体已就绪"
elif [ "$MODE" = "check" ]; then
  log "缺少：${APT_MISSING[*]}"
elif have apt-get; then
  log "安装 ${APT_MISSING[*]}（首次约 1-2 分钟）"
  $SUDO apt-get update -qq 2>/dev/null
  DEBIAN_FRONTEND=noninteractive $SUDO apt-get install -y -qq --no-install-recommends \
    "${APT_MISSING[@]}" >/dev/null 2>&1 \
    && log "完成" \
    || log "安装失败，渲染预览不可用（生成 pptx 不受影响）"
else
  log "非 apt 系统，请手动安装：${APT_MISSING[*]}"
  log "  macOS: brew install --cask libreoffice && brew install poppler"
fi

# HTML 汇报的逐页截图需要 Chrome/Chromium。缺了不装 —— 体积大，且多数机器已有。
have google-chrome || have chromium || have chromium-browser \
  || log "未检测到 Chrome/Chromium：HTML 汇报的截图复核不可用，生成 html 不受影响"

echo "==> 完成"
if [ "$MODE" != "check" ] && [ -f "$SKILLS/exec-deck-builder/scripts/doctor.py" ]; then
  python3 "$SKILLS/exec-deck-builder/scripts/doctor.py" --quiet
fi
