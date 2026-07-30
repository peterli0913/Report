#!/usr/bin/env python3
"""检查汇报材料 skill 集的运行环境，缺什么给出准确的安装命令。

换机器、换环境、或者哪个脚本报 ImportError 时先跑这个：

    python3 doctor.py            # 完整体检
    python3 doctor.py --quiet    # 只在有问题时输出
    python3 doctor.py --json     # 机器可读

退出码：0 全部就绪或只缺可选项；1 缺必需项（连 pptx 都生成不了）。
"""

import importlib
import json
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_ROOT = os.path.dirname(HERE)
SKILLS_DIR = os.path.dirname(SKILL_ROOT)
REPO = os.path.dirname(os.path.dirname(SKILLS_DIR))

INSTALL_HINT = {
    "python-pptx": "pip install python-pptx",
    "Pillow": "pip install Pillow",
    "libreoffice": (
        "macOS: brew install --cask libreoffice\n"
        "      Ubuntu: sudo apt-get install -y libreoffice-impress\n"
        "      Windows: 从 libreoffice.org 安装并把 program 目录加入 PATH"
    ),
    "poppler": (
        "macOS: brew install poppler\n"
        "      Ubuntu: sudo apt-get install -y poppler-utils"
    ),
    "cjk-font": (
        "Ubuntu: sudo apt-get install -y fonts-noto-cjk\n"
        "      macOS/Windows 自带中文字体，通常无需安装"
    ),
    "chrome": "装 Chrome 或 Chromium 即可（多数机器已有）",
}


def _mod(name):
    try:
        importlib.import_module(name)
        return True
    except Exception:
        return False


def _registered_path(name):
    """返回该模块在本脚本目录之外解析到的文件路径，未注册则返回 None。

    两个坑：
    - 直接 importlib 会误判：doctor.py 自己就在 scripts/ 里，Python 会把该目录放进
      sys.path[0]，同目录的 deckkit.py 于是"总是能 import"。所以起子进程测。
    - 只报"已注册"不够：机器上可能装着**另一个项目**的同名副本，于是你改了本仓库
      的 deckkit 却不生效，这种问题极难排查。所以要把解析到的路径打出来。
    """
    try:
        r = subprocess.run(
            [sys.executable, "-c",
             "import %s,sys; print(getattr(%s,'__file__','?'))" % (name, name)],
            capture_output=True, text=True, timeout=30,
            cwd=os.path.expanduser("~"))
        return r.stdout.strip() if r.returncode == 0 else None
    except Exception:
        return None


def _cmd(*names):
    for n in names:
        if shutil.which(n):
            return n
    return None


def _has_cjk_font():
    fc = shutil.which("fc-list")
    if not fc:
        return None            # 非 fontconfig 系统（macOS / Windows）无法这样判断
    try:
        out = subprocess.run([fc, ":lang=zh"], capture_output=True, text=True,
                             timeout=20).stdout
        return bool(out.strip())
    except Exception:
        return None


def collect():
    """返回检查项列表：(名称, 是否必需, 状态, 说明)。状态为 True/False/None(未知)。"""
    rows = []

    rows.append(("python-pptx", True, _mod("pptx"), "生成 pptx 的核心依赖"))
    rows.append(("Pillow", True, _mod("PIL"),
                 "图片裁剪、图标与背景生成、渲染图拼版"))

    dk = _registered_path("deckkit")
    if dk:
        # 指向别的仓库时要明确警告：那意味着你改本仓库的代码不会生效
        same = os.path.dirname(os.path.abspath(dk)) == os.path.abspath(HERE)
        note = ("解析到 %s%s" % (dk, "" if same else "  ← 不是本仓库的副本！"))
    else:
        note = "装了之后任何目录都能 import deckkit"
    rows.append(("deckkit 已注册为模块", False, bool(dk), note))

    soffice = _cmd("soffice", "libreoffice")
    rows.append(("LibreOffice", False, bool(soffice),
                 "把 pptx 渲染成图做视觉复核" + (" (%s)" % soffice if soffice else "")))
    pdftoppm = _cmd("pdftoppm")
    rows.append(("poppler (pdftoppm)", False, bool(pdftoppm),
                 "把 PDF 转成逐页 PNG"))
    rows.append(("中文字体", False, _has_cjk_font(),
                 "渲染预览里的中文不出方块"))
    chrome = _cmd("google-chrome", "chromium", "chromium-browser", "chrome")
    rows.append(("Chrome/Chromium", False, bool(chrome),
                 "HTML 汇报的逐页截图" + (" (%s)" % chrome if chrome else "")))
    return rows


def main():
    quiet = "--quiet" in sys.argv
    as_json = "--json" in sys.argv
    rows = collect()

    missing_required = [r for r in rows if r[1] and r[2] is False]
    missing_optional = [r for r in rows if not r[1] and r[2] is False]

    if as_json:
        print(json.dumps({
            "ok": not missing_required,
            "items": [{"name": n, "required": q, "ok": s, "note": d}
                      for n, q, s, d in rows],
        }, ensure_ascii=False, indent=2))
        return 1 if missing_required else 0

    # 指向别的仓库的副本即使"已注册"也要提醒，否则改了代码不生效很难查
    wrong_copy = [r for r in rows if "deckkit" in r[0] and "不是本仓库" in r[3]]

    if quiet and not missing_required and not missing_optional and not wrong_copy:
        return 0

    if not quiet:
        print("汇报材料 skill 环境检查")
        print("  skill 目录：%s" % SKILLS_DIR)
        print()
        for name, required, state, note in rows:
            mark = {True: "就绪", False: "缺失", None: "未知"}[state]
            tag = "必需" if required else "可选"
            print("  [%s] %-22s %s  %s" % (mark, name, tag, note))
        print()

    def hint_for(name):
        key = ("python-pptx" if "pptx" in name else
               "Pillow" if "Pillow" in name else
               "libreoffice" if "LibreOffice" in name else
               "poppler" if "poppler" in name else
               "cjk-font" if "字体" in name else
               "chrome" if "Chrome" in name else None)
        return INSTALL_HINT.get(key)

    if missing_required:
        print("缺少必需依赖，现在连 pptx 都生成不了：")
        for name, _, _, _ in missing_required:
            print("  %s -> %s" % (name, hint_for(name)))
        print()
        print("一次装齐：bash %s" % os.path.join(REPO, ".cursor", "setup.sh"))
        return 1

    if missing_optional:
        print("缺少可选依赖（核心功能可用，但下列能力不可用）：")
        for name, _, _, note in missing_optional:
            if "deckkit" in name:
                continue
            print("  %s（%s）" % (name, note))
            print("      %s" % hint_for(name))
        if any("deckkit" in r[0] for r in missing_optional):
            print("  deckkit 未注册为模块：脚本仍可用，但 import 时要自己加 sys.path。")
            print("      pip install -e %s" % SKILL_ROOT)
        print()
        print("一次装齐：bash %s" % os.path.join(REPO, ".cursor", "setup.sh"))

    if wrong_copy:
        print("注意：import deckkit 解析到的不是本仓库的副本。")
        print("  %s" % wrong_copy[0][3].replace("解析到 ", "当前指向："))
        print("  这意味着你改本仓库的 deckkit.py 不会生效。要切回本仓库：")
        print("      pip install -e %s" % SKILL_ROOT)
    elif not missing_optional and not quiet:
        print("全部就绪：生成、质检、渲染预览、HTML 截图都可用。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
