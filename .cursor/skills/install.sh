#!/usr/bin/env bash
# 把这套汇报材料 skill 装到用户级目录，让它在**所有项目**里都可用。
#
# 为什么需要这一步：仓库里的 .cursor/skills/ 只对本仓库生效。装到用户级目录后，
# 你在任何项目、任何 agent 会话里说"做个汇报 PPT"都能触发。
#
# 用法：
#   bash install.sh                  # 复制到 ~/.cursor/skills（推荐，独立于本仓库）
#   bash install.sh --link           # 改用符号链接，仓库里改动即时生效（开发用）
#   bash install.sh --all-tools      # 同时让 Claude Code / Codex 也能用
#   bash install.sh --target DIR     # 装到指定目录（例如另一个项目的 .cursor/skills）
#   bash install.sh --only exec-deck-builder,exec-visual-system
#   bash install.sh --uninstall      # 移除已安装的这几个 skill
#   bash install.sh --list           # 只列出会装什么，不动手
#
# Cloud Agent 注意：Cursor 文档没有承诺 Cloud Agent 会读取 VM 上的用户级目录，
# 所以给 Cloud Agent 用请把 skill 提交进仓库的 .cursor/skills/（本仓库已经是）。

set -euo pipefail

SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$SRC/../.." && pwd)"

# 本套件包含的 skill。官方 Apache-2.0 的三个一并带上，它们是自研 skill 的协作方。
ALL_SKILLS=(exec-deck-builder html-exec-report exec-visual-system
            production-report-narrative deck-imagery
            theme-factory frontend-design internal-comms)

MODE="copy"
TARGET="$HOME/.cursor/skills"
ALL_TOOLS=0
ACTION="install"
ONLY=""
CUSTOM_TARGET=0
WANT_PIP="auto"

while [ $# -gt 0 ]; do
  case "$1" in
    --link)      MODE="link" ;;
    --all-tools) ALL_TOOLS=1 ;;
    --target)    TARGET="${2:?--target 需要一个目录}"; CUSTOM_TARGET=1; shift ;;
    --only)      ONLY="${2:?--only 需要 skill 名，逗号分隔}"; shift ;;
    --pip)       WANT_PIP="yes" ;;
    --no-pip)    WANT_PIP="no" ;;
    --uninstall) ACTION="uninstall" ;;
    --list)      ACTION="list" ;;
    -h|--help)   sed -n '2,24p' "${BASH_SOURCE[0]}" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *)           echo "未知参数：$1（用 --help 看用法）" >&2; exit 2 ;;
  esac
  shift
done

# 装到别的项目目录时默认不动 pip 注册：那会把全局的 import deckkit 指到那个副本，
# 副本一删就坏。要显式指定才注册。
if [ "$WANT_PIP" = "auto" ]; then
  [ "$CUSTOM_TARGET" = "1" ] && WANT_PIP="no" || WANT_PIP="yes"
fi

# 确定要处理哪些 skill
SKILLS=()
if [ -n "$ONLY" ]; then
  IFS=',' read -ra want <<< "$ONLY"
  for w in "${want[@]}"; do
    w="$(echo "$w" | tr -d ' ')"
    [ -f "$SRC/$w/SKILL.md" ] || { echo "找不到 skill：$w" >&2; exit 2; }
    SKILLS+=("$w")
  done
else
  for s in "${ALL_SKILLS[@]}"; do
    [ -f "$SRC/$s/SKILL.md" ] && SKILLS+=("$s")
  done
fi

if [ "$ACTION" = "list" ]; then
  echo "源目录：$SRC"
  echo "将安装到：$TARGET"
  echo "方式：$([ "$MODE" = link ] && echo 符号链接 || echo 复制)"
  echo "包含 ${#SKILLS[@]} 个 skill："
  for s in "${SKILLS[@]}"; do
    desc="$(grep -m1 '^description:' "$SRC/$s/SKILL.md" | cut -c14- | cut -c1-56)"
    printf '  %-30s %s...\n' "$s" "$desc"
  done
  exit 0
fi

# 兼容目录：Cursor 会同时加载 .claude / .codex 下的 skills，所以装一份就够两边用
TARGETS=("$TARGET")
if [ "$ALL_TOOLS" = "1" ]; then
  TARGETS+=("$HOME/.claude/skills" "$HOME/.codex/skills")
fi

if [ "$ACTION" = "uninstall" ]; then
  for t in "${TARGETS[@]}"; do
    for s in "${SKILLS[@]}"; do
      if [ -e "$t/$s" ] || [ -L "$t/$s" ]; then
        rm -rf "$t/$s"
        echo "已移除 $t/$s"
      fi
    done
  done
  echo
  echo "如需一并注销 Python 模块：pip uninstall asymchem-deckkit"
  exit 0
fi

for t in "${TARGETS[@]}"; do
  mkdir -p "$t"
  echo "==> 安装到 $t"
  for s in "${SKILLS[@]}"; do
    dst="$t/$s"
    if [ -e "$dst" ] || [ -L "$dst" ]; then
      rm -rf "$dst"
    fi
    if [ "$MODE" = "link" ]; then
      ln -s "$SRC/$s" "$dst"
      printf '  %-30s 已链接\n' "$s"
    else
      cp -R "$SRC/$s" "$dst"
      # 预览图和 __pycache__ 不必跟着走
      rm -rf "$dst/preview" "$dst/assets/preview"
      find "$dst" -name __pycache__ -type d -prune -exec rm -rf {} + 2>/dev/null || true
      printf '  %-30s 已复制\n' "$s"
    fi
  done
done

if [ "$WANT_PIP" = "yes" ]; then
  echo
  echo "==> 注册 deckkit 模块（让任何目录都能 import deckkit）"
  PKG="${TARGETS[0]}/exec-deck-builder"
  [ "$MODE" = "link" ] && PKG="$SRC/exec-deck-builder"
  if [ -f "$PKG/pyproject.toml" ]; then
    PIP_HELP="$(python3 -m pip install --help 2>/dev/null || true)"
    FLAGS=""
    case "$PIP_HELP" in *break-system-packages*) FLAGS="--break-system-packages" ;; esac
    if python3 -m pip install --quiet -e "$PKG" $FLAGS 2>/dev/null \
       || python3 -m pip install --quiet -e "$PKG" 2>/dev/null; then
      echo "  完成，可用命令：deck-doctor / deck-check / deck-render"
    else
      echo "  可编辑安装失败，脚本仍可用（agent 会按路径查找）"
    fi
  fi
else
  echo
  echo "==> 跳过 deckkit 的 pip 注册（装到了自定义目录，避免改动全局 import 指向）"
  echo "    需要的话加 --pip"
fi

echo
echo "==> 体检"
python3 "${TARGETS[0]}/exec-deck-builder/scripts/doctor.py" --quiet || true

cat <<EOF

装好了。接下来：

  · 本地 Cursor：重启或执行 Developer: Reload Window，然后在任意项目里直接说
    「帮我做一份汇报 PPT」即可触发。
  · 缺渲染预览工具（LibreOffice 等）时跑：bash $REPO/.cursor/setup.sh
  · 给 Cloud Agent 用：把 skill 提交进目标仓库的 .cursor/skills/，
    并在该仓库放一份 .cursor/environment.json（可直接复制本仓库的）。
EOF
