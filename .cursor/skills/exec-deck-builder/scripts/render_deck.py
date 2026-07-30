#!/usr/bin/env python3
"""把 pptx 渲染成逐页 PNG，用于肉眼复核版式。

视觉复核是不可跳过的一步：溢出、重叠、对比度不足这些问题在代码里看不出来，
只有渲染出来才会暴露。

用法：
    python3 render_deck.py deck.pptx [输出目录] [--grid]

    --grid  额外拼一张总览图，便于一次看完整套版式节奏

注意：LibreOffice 用的是本机字体。如果本机没装"微软雅黑"，它会用别的中文字体
替代，字宽略有差异，所以预览中"刚好卡边"的文字在真实 PowerPoint 里可能不同。
判断是否溢出以 check_deck.py 的估算为准，预览用来看版式和配色。
"""

import os
import subprocess
import sys
import glob
import shutil
import tempfile


def find_soffice():
    for c in ("soffice", "libreoffice",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe"):
        p = shutil.which(c) or (c if os.path.exists(c) else None)
        if p:
            return p
    return None


def render(pptx, outdir, dpi=110, make_grid=False):
    soffice = find_soffice()
    if not soffice:
        print("找不到 LibreOffice。安装方式：")
        print("  macOS   brew install --cask libreoffice")
        print("  Ubuntu  sudo apt-get install -y libreoffice-impress poppler-utils")
        print("  Windows 从 libreoffice.org 下载安装，并把 program 目录加入 PATH")
        return []
    os.makedirs(outdir, exist_ok=True)
    for old in glob.glob(os.path.join(outdir, "slide-*.png")):
        os.remove(old)

    # 独立 profile 避免与本机 LibreOffice 会话冲突而挂起。profile 和 HOME 都放到
    # 临时目录 —— 指到输出目录会在里面留下 .cache / .config 一堆配置文件。
    work = tempfile.mkdtemp(prefix="deckkit-lo-")
    try:
        cmd = [soffice, "--headless", "--norestore", "--invisible",
               "-env:UserInstallation=file://" + os.path.join(work, "profile"),
               "--convert-to", "pdf", "--outdir", outdir, os.path.abspath(pptx)]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=600,
                           env=dict(os.environ, HOME=work))
    finally:
        shutil.rmtree(work, ignore_errors=True)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    if not os.path.exists(pdf):
        print("转 PDF 失败：", r.stdout, r.stderr)
        return []

    if not shutil.which("pdftoppm"):
        print("已生成 PDF：%s（缺 pdftoppm，无法转 PNG，请安装 poppler-utils）" % pdf)
        return [pdf]
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf,
                    os.path.join(outdir, "slide")], check=True)
    # pdftoppm 按总页数决定补零位数（9 页以内是 slide-1.png，10 页起是 slide-01.png）。
    # 统一成固定位数，免得调用方按页数猜文件名。
    import re
    found = glob.glob(os.path.join(outdir, "slide-*.png"))
    width = 3 if len(found) > 99 else 2
    pages = []
    for p in found:
        m = re.search(r"slide-(\d+)\.png$", os.path.basename(p))
        if not m:
            continue
        new = os.path.join(outdir, "slide-%0*d.png" % (width, int(m.group(1))))
        if os.path.abspath(p) != os.path.abspath(new):
            os.replace(p, new)
        pages.append(new)
    pages.sort()
    print("已渲染 %d 页 -> %s" % (len(pages), outdir))
    for p in pages:
        print("  " + os.path.abspath(p))

    if make_grid and pages:
        try:
            from PIL import Image
            ims = [Image.open(p) for p in pages]
            cols = 3
            rows = (len(ims) + cols - 1) // cols
            tw = 640
            th = int(tw * ims[0].height / ims[0].width)
            grid = Image.new("RGB", (cols * tw, rows * th), "white")
            for i, im in enumerate(ims):
                grid.paste(im.resize((tw, th), Image.LANCZOS),
                           ((i % cols) * tw, (i // cols) * th))
            gp = os.path.join(outdir, "overview.png")
            grid.save(gp)
            print("总览图：" + os.path.abspath(gp))
        except ImportError:
            print("（缺 Pillow，跳过总览图）")
    return pages


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    src = sys.argv[1]
    args = [a for a in sys.argv[2:] if not a.startswith("--")]
    out = args[0] if args else os.path.join(os.path.dirname(os.path.abspath(src))
                                            or ".", "preview")
    render(src, out, make_grid="--grid" in sys.argv)
