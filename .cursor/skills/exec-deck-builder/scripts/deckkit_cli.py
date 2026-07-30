"""命令行入口，pip 安装后可直接用 deck-check / deck-render / deck-doctor。

装过之后就不用记 skill 的目录位置了：

    deck-check 汇报.pptx --strict
    deck-render 汇报.pptx preview/ --grid
    deck-doctor

没装也没关系，直接跑脚本等价：
    python3 <skill>/scripts/check_deck.py 汇报.pptx --strict
"""

import os
import runpy
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def _run(script):
    """以脚本方式执行同目录下的工具，保持它们各自的 __main__ 行为不变。"""
    sys.path.insert(0, HERE)
    path = os.path.join(HERE, script)
    if not os.path.exists(path):
        sys.exit("找不到 %s" % path)
    runpy.run_path(path, run_name="__main__")


def check():
    _run("check_deck.py")


def render():
    _run("render_deck.py")


def doctor():
    _run("doctor.py")


def selftest():
    _run("selftest.py")
