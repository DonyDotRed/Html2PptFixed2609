#!/usr/bin/env python3
"""deckforge-shell.html + deckforge-app.js (+ deckforge-engine.js) 로 두 가지 완성본을 만듭니다.

  index.html          웹판   — 엔진을 인터넷(jsDelivr/unpkg)에서 받습니다. 약 120 KB.
                              GitHub Pages 처럼 인터넷이 되는 곳에 올릴 때 씁니다.
  index-offline.html  오프라인판 — 엔진까지 파일 안에 넣습니다. 약 580 KB.
                              사내망·USB 등 인터넷이 막힌 곳에서 씁니다.

사용법:  python3 deckforge-build.py     (파일들과 같은 폴더에서)
"""
import pathlib, re, sys

here = pathlib.Path(__file__).parent
shell = (here / "deckforge-shell.html").read_text(encoding="utf-8")
app = (here / "deckforge-app.js").read_text(encoding="utf-8")
engine_path = here / "deckforge-engine.js"

LOADER = re.compile(r"<!--PPTX_LOADER_START-->.*?<!--PPTX_LOADER_END-->", re.S)
APPTAG = '<script src="deckforge-app.js"></script>'

WEB_LOADER = """<script>
/* 엔진을 인터넷에서 먼저 받고, 안 되면 같은 폴더의 파일을 찾습니다 */
window.__pptxPaths = [
  'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js',
  'https://unpkg.com/pptxgenjs@3.12.0/dist/pptxgen.bundle.js',
  'deckforge-engine.js', 'pptxgen.bundle.js'
];
(function () {
  var i = 0;
  function done() { window.__pptxDone = true; if (window.checkLib) window.checkLib(); }
  function next() {
    if (typeof PptxGenJS !== 'undefined') return done();
    if (i >= window.__pptxPaths.length) return done();
    var el = document.createElement('script');
    window.__pptxTried = (window.__pptxTried || []).concat(window.__pptxPaths[i]);
    el.src = window.__pptxPaths[i++];
    el.onload = next; el.onerror = next;
    document.head.appendChild(el);
  }
  next();
  setTimeout(done, 10000);
})();
</script>"""


def build(mode):
    html = shell
    if mode == "offline":
        if not engine_path.exists():
            sys.exit("deckforge-engine.js 가 없어 오프라인판을 만들 수 없습니다.")
        lib = engine_path.read_text(encoding="utf-8")
        html, n = LOADER.subn(
            lambda m: "<script>window.__pptxInline=true;</script>", html, count=1)
        # 앱을 먼저 실행하고 엔진을 뒤에 붙입니다. 파일이 잘리면 앱이 그 사실을 알려 줍니다.
        html = html.replace(APPTAG, "<script>\n" + app + "\n</script>\n<script>\n" + lib +
                            "\nwindow.__pptxDone=true;if(window.checkLib)window.checkLib();\n</script>")
        (here / "index-offline.html").write_text(html, encoding="utf-8")
        out = here / "index-offline.html"
        assert n == 1, "엔진 로더 구간을 찾지 못했습니다"
        print(f"만들었습니다: {out.name}  ({out.stat().st_size/1024:.0f} KB)")
        return
    else:
        html, n = LOADER.subn(lambda m: WEB_LOADER, html, count=1)
        out = here / "index.html"
    assert n == 1, "엔진 로더 구간을 찾지 못했습니다"
    html = html.replace(APPTAG, "<script>\n" + app + "\n</script>")
    assert APPTAG not in html, "앱 스크립트 태그를 찾지 못했습니다"
    out.write_text(html, encoding="utf-8")
    print(f"만들었습니다: {out.name}  ({out.stat().st_size/1024:.0f} KB)")


build("web")
build("offline")
