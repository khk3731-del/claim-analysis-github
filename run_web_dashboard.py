from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
import webbrowser

ROOT = Path(__file__).resolve().parent
server = ThreadingHTTPServer(("127.0.0.1", 8765), lambda *args, **kwargs: SimpleHTTPRequestHandler(*args, directory=str(ROOT), **kwargs))
print("웹 대시보드: http://127.0.0.1:8765/web_dashboard.html")
webbrowser.open("http://127.0.0.1:8765/web_dashboard.html")
server.serve_forever()
