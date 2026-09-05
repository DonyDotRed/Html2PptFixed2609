#!/usr/bin/env python3
"""이 폴더를 http://localhost:8000 으로 띄웁니다. 종료는 Ctrl+C."""
import http.server, socketserver, pathlib, os
os.chdir(pathlib.Path(__file__).parent)
h = http.server.SimpleHTTPRequestHandler
h.extensions_map['.js'] = 'text/javascript'
with socketserver.TCPServer(("", 8000), h) as s:
    print("http://localhost:8000  (Ctrl+C 로 종료)")
    try: s.serve_forever()
    except KeyboardInterrupt: print("\n종료했습니다.")
