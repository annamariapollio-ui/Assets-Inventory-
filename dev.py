"""Local preview: serves public/ and the API from memory (data is lost on restart).
   python3 dev.py   ->   prints an admin link and a demo market link."""
import os, re, sys, mimetypes
from http.server import ThreadingHTTPServer
from pathlib import Path

os.environ.setdefault("ADMIN_KEY", "dev")
sys.path.insert(0, str(Path(__file__).parent / "api"))
import index as api  # noqa: E402

PUB = Path(__file__).parent / "public"


class Dev(api.handler):
    def do_GET(self):
        p = self.path.split("?")[0]
        if p.startswith("/api/"):
            return super().do_GET()
        if re.fullmatch(r"/m/[a-z0-9-]+/?", p):
            f = "market.html"
        elif p == "/admin":
            f = "admin.html"
        elif p == "/":
            f = "index.html"
        else:
            f = p.lstrip("/")
        fp = (PUB / f).resolve()
        if PUB not in fp.parents or not fp.is_file():
            return self._json({"error": "Not found"}, 404)
        ctype = mimetypes.guess_type(str(fp))[0] or "application/octet-stream"
        self._send(200, fp.read_bytes(), ctype + ("; charset=utf-8" if ctype.startswith("text/") or "javascript" in ctype else ""))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8100"))
    m = api.get_meta("switzerland")
    print(f"Demo market: http://localhost:{port}/m/{m['slug']}?k={m['key']}")
    print(f"Admin:       http://localhost:{port}/admin?a={os.environ['ADMIN_KEY']}")
    ThreadingHTTPServer(("0.0.0.0", port), Dev).serve_forever()
