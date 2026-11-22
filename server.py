#!/usr/bin/env python3
"""
Watchtower in the browser.

    python3 server.py          then open http://localhost:4918

Runs the same watchtower.py checks on the accounts/ workbooks and shows the
morning report as a page. The date slider replays any morning between Aug 4
and Oct 12, so you can watch when each alert would first have fired.
"""
import json
import os
import re
import subprocess
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
PORT = int(os.environ.get("PORT", "4918"))
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def run_checks(as_of):
    cmd = [sys.executable, os.path.join(HERE, "watchtower.py"), "--json"]
    if as_of:
        cmd += ["--as-of", as_of]
    out = subprocess.run(cmd, capture_output=True, text=True, timeout=120, cwd=HERE)
    if out.returncode != 0:
        raise RuntimeError(out.stderr.strip() or "watchtower.py failed")
    return json.loads(out.stdout)


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype):
        data = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        url = urlparse(self.path)
        if url.path in ("/", "/index.html"):
            with open(os.path.join(HERE, "dashboard.html"), encoding="utf-8") as f:
                return self._send(200, f.read(), "text/html; charset=utf-8")
        if url.path == "/api/run":
            as_of = (parse_qs(url.query).get("as_of") or [""])[0]
            if as_of and not DATE_RE.match(as_of):
                return self._send(400, json.dumps({"error": "as_of must be YYYY-MM-DD"}), "application/json")
            try:
                return self._send(200, json.dumps({"as_of": as_of or None, "accounts": run_checks(as_of)}), "application/json")
            except Exception as exc:  # report the failure to the page instead of a blank screen
                return self._send(500, json.dumps({"error": str(exc)}), "application/json")
        self._send(404, "Not found", "text/plain")

    def log_message(self, *args):
        pass


if __name__ == "__main__":
    print(f"Watchtower running at http://localhost:{PORT}")
    ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
