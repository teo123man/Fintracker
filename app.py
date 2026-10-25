# -*- coding: utf-8 -*-
"""Harcamalarım — the local app. Double-click Harcamalarim.bat (or `python app.py`).

Runs a tiny web server on YOUR computer only (127.0.0.1 — nothing on the network can reach it) and opens the
dashboard in your browser. From the page you can:
  • press "Import" → every new file in the Belgeler folder is parsed into data/*.csv and the dashboard rebuilds
    (or drop files on the page: they're copied into that folder first)
  • categorise unknown merchants (as a rule for the future, or for one transaction)
  • set monthly budgets
Nothing is uploaded anywhere: "upload" means the browser hands the file to this program on the same PC.
The only outside request the importer ever makes is to the Central Bank (TCMB) for EUR/USD rates.

    python app.py            # your data (the folder in home.txt, or HARCAMA_HOME)
    python app.py --demo     # made-up data in demo/ (python demo.py creates it) — for developing features
"""
from __future__ import annotations

import json
import os
import re
import secrets
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

sys.stdout.reconfigure(encoding="utf-8")
CODE = Path(__file__).resolve().parent
if "--demo" in sys.argv:
    os.environ["HARCAMA_HOME"] = str(CODE / "demo")
    if not (CODE / "demo" / "data").exists():
        subprocess.run([sys.executable, str(CODE / "demo.py")], check=True)

import harcama as H  # noqa: E402  (reads HARCAMA_HOME at import)

TOKEN = secrets.token_urlsafe(24)       # only the page this server hands out knows it
MAX_UPLOAD = 40 * 1024 * 1024
ALLOWED = {".pdf", ".xlsx"}
LOCK = threading.Lock()                 # one import / rebuild at a time
PORT = 8765


def run_pipeline(*args: str) -> tuple[bool, str]:
    env = dict(os.environ, HARCAMA_HOME=str(H.ROOT), PYTHONIOENCODING="utf-8")
    with LOCK:
        p = subprocess.run([sys.executable, str(CODE / "harcama.py"), *args, "--no-recalc"], env=env,
                           capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900)
    log = (p.stdout + ("\n" + p.stderr if p.returncode else "")).strip()
    return p.returncode == 0, log


def safe_name(raw: str) -> str:
    name = Path(unquote(raw or "")).name.strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)[:120]
    return name


def folder_files():
    """The documents in the folder and what happened to each (from data/files.csv)."""
    reg = {r["sha256"]: r for r in H.read_csv(H.FILES_CSV)}
    out = []
    for f in H.scan_documents():
        r = reg.get(H._sha256(f))
        out.append({"file": f.relative_to(H.ROOT).as_posix() if f.is_relative_to(H.ROOT) else str(f),
                    "size": f.stat().st_size, "status": r["result"] if r else "new", "date": r["date"] if r else ""})
    return out


def save_upload(name: str, body: bytes) -> str:
    """Copy a dropped file into the documents folder (never overwrite; identical copies are not duplicated)."""
    H.DOCS_DIR.mkdir(parents=True, exist_ok=True)
    dest = H.DOCS_DIR / name
    n = 2
    while dest.exists():
        if dest.read_bytes() == body:
            return dest.name
        dest = H.DOCS_DIR / f"{Path(name).stem} ({n}){Path(name).suffix}"
        n += 1
    dest.write_bytes(body)
    return dest.name


def add_rule(match: str, merchant: str, category: str, sub: str):
    """New rules go first: the first matching rule wins, so yours beat the older, broader ones."""
    words = H.fold(match).split()
    pattern = r"\s+".join(re.escape(w) for w in words)
    rules = [r for r in H.read_csv(H.RULES_CSV) if r.get("pattern") != pattern]
    rules.insert(0, dict(pattern=pattern, merchant=merchant, category=category, subcategory=sub,
                         confidence="manual", note="added in the app"))
    H.write_csv(H.RULES_CSV, ["pattern", "merchant", "category", "subcategory", "confidence", "note"], rules)


def override(tx_id: str, merchant: str, category: str, sub: str) -> bool:
    txs = H.read_csv(H.TX_CSV)
    hit = next((t for t in txs if t["id"] == tx_id), None)
    if not hit:
        return False
    hit.update(merchant_override=merchant, category_override=category, subcategory_override=sub)
    H.write_csv(H.TX_CSV, H.TX_FIELDS, txs)
    return True


def set_budgets(budgets: dict):
    rows = H.read_csv(H.BUDGETS_CSV)
    have = {r["category"] for r in rows}
    for r in rows:
        if r["category"] in budgets:
            v = budgets[r["category"]]
            r["monthly_budget"] = "" if v in (None, "", 0) else f"{float(v):.0f}"
    for c, v in budgets.items():
        if c not in have and v:
            rows.append(dict(category=c, monthly_budget=f"{float(v):.0f}"))
    H.write_csv(H.BUDGETS_CSV, ["category", "monthly_budget"], rows)


class Handler(BaseHTTPRequestHandler):
    server_version = "Harcamalarim"

    def log_message(self, fmt, *args):  # keep the console quiet (no file names / amounts in it)
        pass

    # ── guards: only this machine, only our own page ──
    def _host_ok(self) -> bool:
        host = (self.headers.get("Host") or "").lower()
        return host in {f"127.0.0.1:{self.server.server_port}", f"localhost:{self.server.server_port}"}

    def _send(self, code: int, body: bytes, ctype: str):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code: int, obj):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def do_GET(self):
        if not self._host_ok():
            return self._send(403, b"forbidden", "text/plain")
        path = self.path.split("?")[0]
        if path in ("/", "/index.html"):
            if not H.OUT_HTML.exists():
                # A new installation has no statements yet. Serve onboarding without
                # trying to calculate totals from an empty ledger.
                if H.ST_CSV.exists():
                    ok, log = run_pipeline("--build")
                    if not ok:
                        return self._send(500, log.encode("utf-8"), "text/plain; charset=utf-8")
            page = H.OUT_HTML if H.OUT_HTML.exists() else CODE / "first_run.html"
            html = page.read_text(encoding="utf-8")
            app = json.dumps({"token": TOKEN, "demo": "--demo" in sys.argv, "folder": str(H.DOCS_DIR)})
            return self._send(200, html.replace("/*__APP__*/null", app.replace("<", "\\u003c"), 1).encode("utf-8"), "text/html; charset=utf-8")
        if path == "/favicon.ico":
            return self._send(204, b"", "image/x-icon")
        return self._send(404, b"not found", "text/plain")

    def do_POST(self):
        if not self._host_ok() or self.headers.get("X-Token") != TOKEN:
            return self._json(403, {"ok": False, "log": "Not allowed."})
        n = int(self.headers.get("Content-Length") or 0)
        if n > MAX_UPLOAD:
            return self._json(413, {"ok": False, "log": "File is too big (40 MB max)."})
        body = self.rfile.read(n) if n else b""
        path = self.path.split("?")[0]
        try:
            if path == "/api/upload":
                name = safe_name(self.headers.get("X-Filename", ""))
                if not name or Path(name).suffix.lower() not in ALLOWED:
                    return self._json(400, {"ok": False, "log": f"{name or 'file'}: only .pdf and .xlsx files"})
                return self._json(200, {"ok": True, "saved": save_upload(name, body)})
            if path == "/api/files":
                return self._json(200, {"ok": True, "folder": str(H.DOCS_DIR), "files": folder_files()})
            if path == "/api/open-folder":
                H.DOCS_DIR.mkdir(parents=True, exist_ok=True)
                os.startfile(H.DOCS_DIR)  # Windows Explorer
                return self._json(200, {"ok": True})
            if path == "/api/import":
                ok, log = run_pipeline()
                return self._json(200, {"ok": ok, "log": log, "ready": H.OUT_HTML.exists()})
            data = json.loads(body or b"{}")
            if path == "/api/categorize":
                merchant = (data.get("merchant") or "").strip()[:80]
                cat = (data.get("category") or "").strip()[:60]
                sub = (data.get("subcategory") or "").strip()[:60] or "General"
                if not merchant or not cat:
                    return self._json(400, {"ok": False, "log": "Merchant and category are required."})
                if data.get("scope") == "rule":
                    match = (data.get("match") or "").strip()
                    if len(match) < 3:
                        return self._json(400, {"ok": False, "log": "The matching text is too short."})
                    add_rule(match, merchant, cat, sub)
                elif not override(str(data.get("id")), merchant, cat, sub):
                    return self._json(404, {"ok": False, "log": "Transaction not found."})
                ok, log = run_pipeline("--build")
                return self._json(200, {"ok": ok, "log": log})
            if path == "/api/budgets":
                set_budgets({str(k)[:60]: v for k, v in (data.get("budgets") or {}).items()})
                ok, log = run_pipeline("--build")
                return self._json(200, {"ok": ok, "log": log})
            if path == "/api/rebuild":
                ok, log = run_pipeline("--build")
                return self._json(200, {"ok": ok, "log": log})
        except Exception as e:  # report, don't crash the server
            return self._json(500, {"ok": False, "log": f"{type(e).__name__}: {e}"})
        return self._json(404, {"ok": False, "log": "Unknown action."})


def main():
    port = PORT
    for port in range(PORT, PORT + 20):
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            continue
    else:
        sys.exit("No free port between 8765 and 8784.")
    url = f"http://127.0.0.1:{port}/"
    print(f"Harcamalarım is running at {url}  ({'DEMO data' if '--demo' in sys.argv else H.ROOT})")
    print(f"Put your statements in: {H.DOCS_DIR}")
    print("Only this computer can open it. Close this window (or Ctrl+C) to stop.")
    if "--no-browser" not in sys.argv:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
