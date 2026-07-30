#!/usr/bin/env python3
"""把汇报页逐页截图，用于视觉复核。

HTML 不看渲染结果等于没测 —— 溢出、错位、对比度问题只有截图才看得见。

用法：
    python3 shoot.py example.html out/            # 逐页截图 + 总览图
    python3 shoot.py example.html out/ --pages 2,5,7
    python3 shoot.py example.html out/ --theme light

依赖 Chrome / Chromium。容器环境下 Chrome 常常截完图不退出，所以这里统一加
超时并只以"文件是否生成"判断成败，不看退出码。
"""

import glob
import os
import shutil
import subprocess
import sys
import tempfile
import time


def find_chrome():
    names = ["google-chrome", "chromium", "chromium-browser", "chrome",
             "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
             r"C:\Program Files\Google\Chrome\Application\chrome.exe"]
    for n in names:
        p = shutil.which(n) or (n if os.path.exists(n) else None)
        if p:
            return p
    return None


def shoot(src, outdir, pages=None, theme=None, w=1280, h=720):
    chrome = find_chrome()
    if not chrome:
        print("找不到 Chrome。macOS/Windows 装 Chrome 即可；")
        print("Ubuntu: sudo apt-get install -y chromium-browser")
        return []
    src = os.path.abspath(src)
    os.makedirs(outdir, exist_ok=True)
    for old in glob.glob(os.path.join(outdir, "page-*.png")):
        os.remove(old)

    # 页数：只数 slide 作为独立 class 的元素。不能用 count('class="slide')，
    # 那样 slide-title / slide-head 这些也会被算进来。
    import re
    with open(src, encoding="utf-8") as f:
        n = len(re.findall(r'class="slide[ "]', f.read()))
    if not n:
        print("没找到 .slide 元素，确认这是汇报页 HTML")
        return []
    todo = pages or list(range(1, n + 1))
    print("共 %d 页，准备截图 %d 页" % (n, len(todo)))

    shots = []
    for i in todo:
        url = "file://%s?hideui=1%s#%d" % (
            src, ("&theme=" + theme) if theme else "", i)
        out = os.path.join(outdir, "page-%02d.png" % i)
        profile = tempfile.mkdtemp(prefix="shoot-")
        cmd = [chrome, "--headless=new", "--no-sandbox", "--disable-gpu",
               "--disable-dev-shm-usage", "--hide-scrollbars", "--no-first-run",
               "--user-data-dir=" + profile,
               "--virtual-time-budget=2000",
               "--window-size=%d,%d" % (w, h),
               "--screenshot=" + out, url]
        # Chrome 在容器里截完图常常不退出。等文件出现就收工，别白等超时 ——
        # 逐页各等一次超时的话，十几页要多花十分钟。
        proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL,
                                stderr=subprocess.DEVNULL)
        deadline = time.time() + 30
        while time.time() < deadline:
            if os.path.exists(out) and os.path.getsize(out) > 0:
                time.sleep(0.25)      # 让它把文件写完
                break
            if proc.poll() is not None:
                break
            time.sleep(0.2)
        proc.kill()
        proc.wait(timeout=5)
        shutil.rmtree(profile, ignore_errors=True)
        if os.path.exists(out):
            shots.append(out)
            print("  " + out)
        else:
            print("  第 %d 页截图失败" % i)

    if len(shots) > 1:
        try:
            from PIL import Image
            ims = [Image.open(p) for p in shots]
            cols = 3
            rows = (len(ims) + cols - 1) // cols
            tw = 640
            th = int(tw * ims[0].height / ims[0].width)
            grid = Image.new("RGB", (cols * tw, rows * th), "#101820")
            for k, im in enumerate(ims):
                grid.paste(im.resize((tw, th), Image.LANCZOS),
                           ((k % cols) * tw, (k // cols) * th))
            gp = os.path.join(outdir, "overview.png")
            grid.save(gp)
            print("总览图：" + os.path.abspath(gp))
        except ImportError:
            print("（缺 Pillow，跳过总览图）")
    return shots


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    src = args[0]
    outdir = args[1] if len(args) > 1 else "preview"

    def opt(name):
        for i, a in enumerate(sys.argv):
            if a == "--" + name and i + 1 < len(sys.argv):
                return sys.argv[i + 1]
        return None

    pg = opt("pages")
    shoot(src, outdir,
          pages=[int(v) for v in pg.split(",")] if pg else None,
          theme=opt("theme"))
