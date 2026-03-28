"""
BFD Generator
Run:  python3 bfd_app.py
Opens browser at http://localhost:8765
"""
import http.server, json, os, threading, webbrowser
from pathlib import Path
from bfd_engine  import BFDEngine
from bfd_excel   import generate_excel
from bfd_diagram import generate_diagram_html

PORT   = 8765
engine = BFDEngine()
HTML   = open(Path(__file__).parent / "bfd_ui.html").read()

class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self._respond(200, HTML.encode(), "text/html; charset=utf-8")
        elif self.path.startswith("/download/"):
            fname = self.path.replace("/download/", "")
            fpath = Path(__file__).parent / "output" / fname
            if fpath.exists():
                self._respond(200, fpath.read_bytes(), "application/octet-stream",
                              extra={"Content-Disposition": f'attachment; filename="{fname}"'})
            else:
                self._json({"error": "File not found"}, 404)
        else:
            self._json({"error": "Not found"}, 404)

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        body   = json.loads(self.rfile.read(length)) if length else {}

        if self.path == "/api/calculate":
            try:
                self._json(engine.calculate(body))
            except Exception as e:
                self._json({"error": str(e)}, 400)

        elif self.path == "/api/export_excel":
            try:
                os.makedirs(Path(__file__).parent / "output", exist_ok=True)
                fname = generate_excel(body, Path(__file__).parent / "output")
                self._json({"filename": fname, "url": f"/download/{fname}"})
            except Exception as e:
                import traceback
                self._json({"error": str(e), "trace": traceback.format_exc()}, 400)

        elif self.path == "/api/diagram":
            try:
                self._json({"html": generate_diagram_html(body)})
            except Exception as e:
                self._json({"error": str(e)}, 400)
        else:
            self._json({"error": "Unknown endpoint"}, 404)

    def _json(self, data, code=200):
        payload = json.dumps(data).encode()
        self._respond(code, payload, "application/json")

    def _respond(self, code, body, ctype, extra=None):
        self.send_response(code)
        self.send_header("Content-Type",   ctype)
        self.send_header("Content-Length", str(len(body)))
        if extra:
            for k, v in extra.items():
                self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)


def run():
    server = http.server.HTTPServer(("localhost", PORT), Handler)
    print(f"\n{'='*50}")
    print(f"  BFD Generator running at http://localhost:{PORT}")
    print(f"  Press Ctrl+C to stop.")
    print(f"{'='*50}\n")
    threading.Timer(1.2, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")

if __name__ == "__main__":
    run()
