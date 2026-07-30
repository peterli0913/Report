#!/usr/bin/env python3
"""把汇报页打包成单个 html 文件，便于通过邮件或 OA 分发。

把外链的 css / js 内联进来，把本地图片转成 base64。这样对方双击就能看，不会
出现"发过去样式全丢了"的情况 —— 邮件附件和 OA 系统通常只留下 html 本体，
同目录的 css、js、图片都不会跟着走。

用法：
    python3 inline.py example.html 汇报.html
    python3 inline.py example.html 汇报.html --no-images   # 图片保持外链
"""

import base64
import mimetypes
import os
import re
import sys


def _read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def inline(src, out, embed_images=True):
    base = os.path.dirname(os.path.abspath(src))
    html = _read(src)
    missing = []

    def css_repl(m):
        href = m.group(1)
        if href.startswith(("http://", "https://", "//", "data:")):
            return m.group(0)
        p = os.path.join(base, href)
        if not os.path.exists(p):
            missing.append(href)
            return m.group(0)
        return "<style>\n%s\n</style>" % _read(p)

    html = re.sub(r'<link[^>]+rel=["\']stylesheet["\'][^>]*href=["\']([^"\']+)["\'][^>]*>',
                  css_repl, html)
    html = re.sub(r'<link[^>]+href=["\']([^"\']+)["\'][^>]*rel=["\']stylesheet["\'][^>]*>',
                  css_repl, html)

    def js_repl(m):
        srcattr = m.group(1)
        if srcattr.startswith(("http://", "https://", "//", "data:")):
            return m.group(0)
        p = os.path.join(base, srcattr)
        if not os.path.exists(p):
            missing.append(srcattr)
            return m.group(0)
        return "<script>\n%s\n</script>" % _read(p)

    html = re.sub(r'<script[^>]+src=["\']([^"\']+)["\'][^>]*>\s*</script>',
                  js_repl, html)

    if embed_images:
        def img_repl(m):
            pre, url, post = m.group(1), m.group(2), m.group(3)
            if url.startswith(("http://", "https://", "//", "data:")):
                return m.group(0)
            p = os.path.join(base, url)
            if not os.path.exists(p):
                missing.append(url)
                return m.group(0)
            mime = mimetypes.guess_type(p)[0] or "application/octet-stream"
            with open(p, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            return "%sdata:%s;base64,%s%s" % (pre, mime, b64, post)

        html = re.sub(r'(<img[^>]+src=["\'])([^"\']+)(["\'])', img_repl, html)
        html = re.sub(r'(url\(["\']?)([^"\')]+)(["\']?\))', img_repl, html)

    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    size = os.path.getsize(out) / 1024.0
    print("已打包：%s（%.0f KB）" % (out, size))
    if missing:
        print("以下资源没找到，仍是外链，对方打开会缺失：")
        for m in sorted(set(missing)):
            print("  " + m)
    if size > 8000:
        print("提示：文件超过 8 MB，部分邮件系统会拒收，考虑压缩图片"
              "或改用 --no-images")
    return out


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    inline(sys.argv[1], sys.argv[2], embed_images="--no-images" not in sys.argv)
