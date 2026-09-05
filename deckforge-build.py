#!/usr/bin/env python3
"""deckforge-shell.html + deckforge-app.js + deckforge-engine.js 를 합쳐 index.html 을 만듭니다.
사용법:  python3 deckforge-build.py   (같은 폴더에서 실행)"""
import pathlib
import re

here = pathlib.Path(__file__).parent
html = (here / "deckforge-shell.html").read_text(encoding="utf-8")
app = (here / "deckforge-app.js").read_text(encoding="utf-8")
lib = (here / "deckforge-engine.js").read_text(encoding="utf-8")

html, n = re.subn(r'<!--PPTX_LOADER_START-->.*?<!--PPTX_LOADER_END-->',
                  lambda m: "<script>\nwindow.__pptxDone=true;window.__pptxInline=true;\n" + lib + "\n</script>",
                  html, count=1, flags=re.S)
assert n == 1, "엔진 로더 구간을 찾지 못했습니다"
html = html.replace(
    '<script src="deckforge-app.js"></script>',
    "<script>\n" + app + "\n</script>")

out = here / "index.html"
out.write_text(html, encoding="utf-8")
print("만들었습니다:", out, f"({out.stat().st_size/1024:.0f} KB)")
