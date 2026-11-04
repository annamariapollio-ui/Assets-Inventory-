#!/usr/bin/env python3
"""
VI Rollout Strategy - Market Inventory survey app.

One codebase, one link per market. Contributors answer the 33 touchpoints
(workstreams A-F) in the browser; answers are shared live between colleagues
and are written into the existing Excel template on submission.

Run:        python3 server.py
Add market: python3 server.py add "Switzerland" CHF
Needs:      Python 3.9+ and openpyxl (no other dependencies).
"""
import hmac
import io
import json
import mimetypes
import os
import re
import secrets
import smtplib
import sys
import threading
import time
from datetime import datetime, timezone
from email.message import EmailMessage
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent
TEMPLATE = BASE / "template" / "NEW_PMI_Market_Quantitative_Workbook.xlsx"
DATA = BASE / "data"
SUBS = BASE / "submissions"
STATIC = BASE / "static"
CONFIG = BASE / "config.json"

LOCK = threading.RLock()
PRESENCE = {}  # slug -> {name: (ref, timestamp)}
PRESENCE_TTL = 20  # seconds

# Switch to False if workstream owners should be optional for submission.
REQUIRE_OWNERS = True

# ---------------------------------------------------------------- vocabulary
EXISTS = ("yes", "no", "unknown")
YESNO = ("yes", "no")
AUDIENCE = ("internal", "external", "both")
VISIBILITY = ("low", "medium", "high")
EFFORT = ("easy / no-cost change", "some effort", "meaningful investment", "unknown")
EFFORT_WITH_COST = ("some effort", "meaningful investment")
BANDS = ("1", "2 to 5", "6 to 20", "21 to 50", "51 to 100", "101 or more")
QTY_MODES = ("exact", "band", "na")
IDENTITIES = (
    "Official corporate logo",
    "Locally adapted corporate logo",
    "Previous / legacy mark",
    "No mark",
    "Other",
)
SUPPORT_TAGS = (
    "Central guidance",
    "Creative or agency help",
    "Technical capacity",
    "Funding",
    "Procurement",
    "Training",
)
CONSTRAINT_TAGS = (
    "Budget timing",
    "Contract terms",
    "Release window",
    "Refurbishment cycle",
    "Approvals",
    "Procurement lead time",
)


# ------------------------------------------------------------ template shape
def load_structure():
    """Read workstreams and touchpoints (with row numbers) from the template."""
    wb = load_workbook(TEMPLATE)
    ws = wb["Inventory"]
    out, cur = [], None
    for r in range(6, 45):
        ref = ws.cell(r, 1).value
        title = ws.cell(r, 2).value
        if not ref or not title:
            continue
        ref, title = str(ref).strip(), str(title).strip()
        if len(ref) == 1:
            name = title.split(". ", 1)[1] if ". " in title else title
            cur = {"ref": ref, "title": name, "row": r, "touchpoints": []}
            out.append(cur)
        elif cur is not None:
            cur["touchpoints"].append({"ref": ref, "title": title, "row": r})
    return out


STRUCT = load_structure()
ALL_REFS = [t["ref"] for w in STRUCT for t in w["touchpoints"]]
WS_REFS = [w["ref"] for w in STRUCT]


# ------------------------------------------------------------------- helpers
def now_iso():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def slugify(name):
    s = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")
    return s or "market"


def read_json(path, default=None):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def write_json(path, obj):
    tmp = str(path) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def get_config():
    cfg = read_json(CONFIG)
    if not cfg:
        cfg = {"admin_key": secrets.token_urlsafe(12)}
        write_json(CONFIG, cfg)
    return cfg


def market_path(slug):
    if not re.fullmatch(r"[a-z0-9-]+", slug or ""):
        return None
    return DATA / f"{slug}.json"


def load_market(slug):
    p = market_path(slug)
    return read_json(p) if p else None


def save_market(m):
    write_json(market_path(m["slug"]), m)


def create_market(name, currency):
    name = (name or "").strip()[:80]
    currency = (currency or "").strip().upper()[:8]
    if not name:
        raise ValueError("Market name needed")
    with LOCK:
        slug, n = slugify(name), 2
        while market_path(slug).exists():
            slug = f"{slugify(name)}-{n}"
            n += 1
        m = {
            "slug": slug,
            "name": name,
            "currency": currency or "EUR",
            "key": secrets.token_urlsafe(8),
            "created": now_iso(),
            "owners": {},
            "tp": {},
            "submissions": [],
        }
        save_market(m)
    return m


def s(v, n=2000):
    return str(v or "").strip()[:n]


def num(v):
    try:
        x = float(str(v).replace(",", "").replace("'", "").strip())
        return x if x >= 0 else None
    except (ValueError, TypeError):
        return None


def clean_answers(v):
    """Keep only known keys with sane values."""
    v = v or {}
    out = {
        "exists": v.get("exists") if v.get("exists") in EXISTS else "",
        "mark": v.get("mark") if v.get("mark") in YESNO else "",
        "assets": [],
        "qty_mode": v.get("qty_mode") if v.get("qty_mode") in QTY_MODES else "",
        "qty_num": s(v.get("qty_num"), 20),
        "qty_unit": s(v.get("qty_unit"), 60),
        "qty_band": v.get("qty_band") if v.get("qty_band") in BANDS else "",
        "audience": v.get("audience") if v.get("audience") in AUDIENCE else "",
        "visibility": v.get("visibility") if v.get("visibility") in VISIBILITY else "",
        "effort": v.get("effort") if v.get("effort") in EFFORT else "",
        "cost_min": s(v.get("cost_min"), 20),
        "cost_max": s(v.get("cost_max"), 20),
        "cost_unknown": bool(v.get("cost_unknown")),
        "support_tags": [t for t in (v.get("support_tags") or []) if t in SUPPORT_TAGS],
        "support_text": s(v.get("support_text")),
        "constraint_tags": [t for t in (v.get("constraint_tags") or []) if t in CONSTRAINT_TAGS],
        "constraint_text": s(v.get("constraint_text")),
        "owner_exception": s(v.get("owner_exception"), 300),
        "comment": s(v.get("comment"), 4000),
    }
    for a in (v.get("assets") or [])[:40]:
        out["assets"].append(
            {
                "name": s(a.get("name"), 300),
                "identity": a.get("identity") if a.get("identity") in IDENTITIES else "",
                "identity_other": s(a.get("identity_other"), 200),
            }
        )
    return out


def missing(v):
    """List of items still open for a touchpoint (empty list = complete)."""
    ex = v.get("exists")
    if ex not in EXISTS:
        return ["Whether the touchpoint exists"]
    if ex != "yes":
        return []
    m = []
    if v.get("mark") not in YESNO:
        m.append("Mark present today")
    assets = [a for a in v.get("assets", []) if a.get("name")]
    if not assets:
        m.append("What specifically exists")
    elif any(
        not a.get("identity") or (a["identity"] == "Other" and not a.get("identity_other"))
        for a in assets
    ):
        m.append("Visual identity per asset")
    mode = v.get("qty_mode")
    if mode == "exact":
        if num(v.get("qty_num")) is None or not v.get("qty_unit"):
            m.append("Quantity (number and unit)")
    elif mode == "band":
        if not v.get("qty_band"):
            m.append("Quantity band")
    elif mode != "na":
        m.append("Quantity")
    if not v.get("audience"):
        m.append("Audience")
    if not v.get("visibility"):
        m.append("Visibility")
    if not v.get("effort"):
        m.append("Effort of change")
    elif v["effort"] in EFFORT_WITH_COST:
        if not v.get("cost_unknown") and num(v.get("cost_min")) is None and num(v.get("cost_max")) is None:
            m.append("Estimated cost range")
    if not v.get("support_tags") and not v.get("support_text"):
        m.append("Support needed")
    if not v.get("constraint_tags") and not v.get("constraint_text"):
        m.append("Rollout constraints")
    return m


def progress(m):
    done = 0
    for ref in ALL_REFS:
        rec = m["tp"].get(ref)
        if rec and not missing(rec["v"]):
            done += 1
    owners_done = sum(1 for w in WS_REFS if (m["owners"].get(w, {}).get("value") or "").strip())
    return {
        "done": done,
        "total": len(ALL_REFS),
        "owners_done": owners_done,
        "owners_total": len(WS_REFS),
    }


def ready_to_submit(m):
    p = progress(m)
    return p["done"] == p["total"] and (not REQUIRE_OWNERS or p["owners_done"] == p["owners_total"])


def last_activity(m):
    times = [r.get("at", "") for r in m["tp"].values()] + [o.get("at", "") for o in m["owners"].values()]
    return max([t for t in times if t] or [m.get("created", "")])


# ------------------------------------------------------------- Excel export
def fmt_money(x):
    return f"{x:,.0f}" if float(x).is_integer() else f"{x:,.2f}"


def identity_text(a):
    return a.get("identity_other", "") if a.get("identity") == "Other" else a["identity"].lower()


def join_tags(tags, text):
    parts = list(tags) + ([text] if text else [])
    return "; ".join(parts)


def build_xlsx(m):
    """Fill the original template: Inventory tab only, Overview and Recap stay untouched."""
    wb = load_workbook(TEMPLATE)
    ws = wb["Inventory"]
    ws["B2"] = m["name"]
    cur = m.get("currency", "")
    for w in STRUCT:
        o = m["owners"].get(w["ref"], {})
        if o.get("value"):
            ws.cell(w["row"], 12).value = o["value"]
        if o.get("comment"):
            ws.cell(w["row"], 15).value = o["comment"]
        for t in w["touchpoints"]:
            rec = m["tp"].get(t["ref"])
            if not rec:
                continue
            v, r = rec["v"], t["row"]
            if v.get("exists"):
                ws.cell(r, 3).value = v["exists"]
            if v.get("exists") == "yes":
                if v.get("mark"):
                    ws.cell(r, 4).value = v["mark"]
                assets = [a for a in v.get("assets", []) if a.get("name")]
                if assets:
                    ws.cell(r, 5).value = "; ".join(a["name"] for a in assets)
                    ws.cell(r, 6).value = "; ".join(
                        f'{a["name"]}: {identity_text(a)}' for a in assets if a.get("identity")
                    )
                mode = v.get("qty_mode")
                if mode == "exact" and num(v.get("qty_num")) is not None:
                    n = num(v["qty_num"])
                    ws.cell(r, 7).value = f'{fmt_money(n)} {v.get("qty_unit", "")}'.strip()
                elif mode == "band" and v.get("qty_band"):
                    ws.cell(r, 7).value = v["qty_band"]
                elif mode == "na":
                    ws.cell(r, 7).value = "Not applicable"
                for col, key in ((8, "audience"), (9, "visibility"), (10, "effort")):
                    if v.get(key):
                        ws.cell(r, col).value = v[key]
                if v.get("effort") in EFFORT_WITH_COST:
                    lo, hi = num(v.get("cost_min")), num(v.get("cost_max"))
                    if lo is not None and hi is not None:
                        ws.cell(r, 11).value = f"{cur} {fmt_money(lo)}\u2013{fmt_money(hi)}"
                    elif lo is not None:
                        ws.cell(r, 11).value = f"{cur} from {fmt_money(lo)}"
                    elif hi is not None:
                        ws.cell(r, 11).value = f"{cur} up to {fmt_money(hi)}"
                    elif v.get("cost_unknown"):
                        ws.cell(r, 11).value = "Not yet known"
                if v.get("owner_exception"):
                    ws.cell(r, 12).value = v["owner_exception"]
                sup = join_tags(v.get("support_tags", []), v.get("support_text"))
                con = join_tags(v.get("constraint_tags", []), v.get("constraint_text"))
                if sup:
                    ws.cell(r, 13).value = sup
                if con:
                    ws.cell(r, 14).value = con
            if v.get("comment"):
                ws.cell(r, 15).value = v["comment"]
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def send_mail(m, filename, content, by, note):
    host, to = os.environ.get("SMTP_HOST"), os.environ.get("MAIL_TO")
    if not host or not to:
        return False
    try:
        msg = EmailMessage()
        msg["Subject"] = f"VI Rollout - market inventory submitted: {m['name']}"
        msg["From"] = os.environ.get("MAIL_FROM", os.environ.get("SMTP_USER", "noreply@localhost"))
        msg["To"] = to
        msg.set_content(
            f"Market inventory for {m['name']} submitted by {by or 'a contributor'}.\n"
            + (f"Note: {note}\n" if note else "")
            + "The completed workbook is attached."
        )
        msg.add_attachment(
            content,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
        )
        with smtplib.SMTP(host, int(os.environ.get("SMTP_PORT", "587")), timeout=20) as smtp:
            smtp.starttls()
            if os.environ.get("SMTP_USER"):
                smtp.login(os.environ["SMTP_USER"], os.environ.get("SMTP_PASS", ""))
            smtp.send_message(msg)
        return True
    except Exception as e:  # delivery by mail is a convenience; the file is always stored
        print("Mail delivery skipped:", e)
        return False


# ------------------------------------------------------------------- state
def public_state(m, slug):
    tp = {}
    for ref, rec in m["tp"].items():
        tp[ref] = {
            "v": rec["v"],
            "rev": rec["rev"],
            "by": rec.get("by", ""),
            "at": rec.get("at", ""),
            "missing": missing(rec["v"]),
        }
    pres = []
    now = time.time()
    for name, (ref, ts) in list(PRESENCE.get(slug, {}).items()):
        if now - ts <= PRESENCE_TTL:
            pres.append({"by": name, "ref": ref})
    return {
        "market": {"name": m["name"], "slug": m["slug"], "currency": m["currency"]},
        "structure": STRUCT,
        "owners": m["owners"],
        "tp": tp,
        "presence": pres,
        "progress": progress(m),
        "ready": ready_to_submit(m),
        "submissions": [{"at": x["at"], "by": x.get("by", "")} for x in m["submissions"]],
        "options": {
            "identities": IDENTITIES,
            "support": SUPPORT_TAGS,
            "constraints": CONSTRAINT_TAGS,
            "bands": BANDS,
        },
    }


# -------------------------------------------------------------------- HTTP
class Handler(BaseHTTPRequestHandler):
    server_version = "VISurvey/1.0"

    def log_message(self, fmt, *args):
        pass

    # --- plumbing
    def send_bytes(self, code, body, ctype, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, val in (extra or {}).items():
            self.send_header(k, val)
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, obj, code=200):
        self.send_bytes(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def body_json(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > 2_000_000:
            return {}
        try:
            return json.loads(self.rfile.read(n) or b"{}")
        except json.JSONDecodeError:
            return {}

    def static_file(self, name):
        p = (STATIC / name).resolve()
        if STATIC not in p.parents or not p.is_file():
            return self.send_json({"error": "Not found"}, 404)
        ctype = mimetypes.guess_type(str(p))[0] or "application/octet-stream"
        if ctype.startswith("text/") or ctype in ("application/javascript",):
            ctype += "; charset=utf-8"
        self.send_bytes(200, p.read_bytes(), ctype)

    def market_auth(self, slug, q):
        m = load_market(slug)
        if not m:
            self.send_json({"error": "Market not found"}, 404)
            return None
        key = self.headers.get("X-Key") or (q.get("k") or [""])[0]
        if not hmac.compare_digest(str(key), m["key"]):
            self.send_json({"error": "Link not recognised"}, 403)
            return None
        return m

    def admin_auth(self, q):
        key = self.headers.get("X-Admin") or (q.get("a") or [""])[0]
        if not hmac.compare_digest(str(key), get_config()["admin_key"]):
            self.send_json({"error": "Admin key not recognised"}, 403)
            return False
        return True

    def origin(self):
        proto = self.headers.get("X-Forwarded-Proto", "http")
        return f"{proto}://{self.headers.get('Host', 'localhost')}"

    # --- routing
    def do_GET(self):
        u = urlparse(self.path)
        q, path = parse_qs(u.query), u.path
        if path == "/":
            return self.static_file("home.html")
        if path.startswith("/static/"):
            return self.static_file(path[len("/static/"):])
        if re.fullmatch(r"/m/[a-z0-9-]+/?", path):
            return self.static_file("index.html")
        if path == "/admin":
            return self.static_file("admin.html")
        mm = re.fullmatch(r"/api/m/([a-z0-9-]+)/(state|export)", path)
        if mm:
            m = self.market_auth(mm.group(1), q)
            if not m:
                return
            if mm.group(2) == "state":
                return self.send_json(public_state(m, m["slug"]))
            return self.send_bytes(
                200,
                build_xlsx(m),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                {"Content-Disposition": f'attachment; filename="{m["slug"]}_draft.xlsx"'},
            )
        if path == "/api/admin/markets":
            if not self.admin_auth(q):
                return
            return self.send_json({"markets": self.admin_list(), "ready_rule_owners": REQUIRE_OWNERS})
        mm = re.fullmatch(r"/api/admin/file/([a-z0-9-]+)/([A-Za-z0-9_.-]+)", path)
        if mm:
            if not self.admin_auth(q):
                return
            f = (SUBS / mm.group(1) / mm.group(2)).resolve()
            if SUBS not in f.parents or not f.is_file():
                return self.send_json({"error": "Not found"}, 404)
            return self.send_bytes(
                200, f.read_bytes(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                {"Content-Disposition": f'attachment; filename="{f.name}"'},
            )
        mm = re.fullmatch(r"/api/admin/export/([a-z0-9-]+)", path)
        if mm:
            if not self.admin_auth(q):
                return
            m = load_market(mm.group(1))
            if not m:
                return self.send_json({"error": "Market not found"}, 404)
            return self.send_bytes(
                200, build_xlsx(m),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                {"Content-Disposition": f'attachment; filename="{m["slug"]}_draft.xlsx"'},
            )
        self.send_json({"error": "Not found"}, 404)

    def admin_list(self):
        out = []
        for p in sorted(DATA.glob("*.json")):
            m = read_json(p)
            if not m:
                continue
            out.append(
                {
                    "slug": m["slug"],
                    "name": m["name"],
                    "currency": m["currency"],
                    "link": f"{self.origin()}/m/{m['slug']}?k={m['key']}",
                    "progress": progress(m),
                    "ready": ready_to_submit(m),
                    "last_activity": last_activity(m),
                    "submissions": m["submissions"],
                }
            )
        return out

    def do_POST(self):
        u = urlparse(self.path)
        q, path = parse_qs(u.query), u.path
        body = self.body_json()

        if path == "/api/admin/markets":
            if not self.admin_auth(q):
                return
            try:
                m = create_market(body.get("name"), body.get("currency"))
            except ValueError as e:
                return self.send_json({"error": str(e)}, 400)
            return self.send_json({"ok": True, "slug": m["slug"]})

        mm = re.fullmatch(r"/api/m/([a-z0-9-]+)/(tp|owner)/([A-Z][0-9]?)", path)
        if mm:
            m = self.market_auth(mm.group(1), q)
            if not m:
                return
            by = s(body.get("by"), 80)
            with LOCK:
                m = load_market(m["slug"])
                if mm.group(2) == "tp":
                    ref = mm.group(3)
                    if ref not in ALL_REFS:
                        return self.send_json({"error": "Unknown touchpoint"}, 404)
                    rec = m["tp"].get(ref)
                    cur_rev = rec["rev"] if rec else 0
                    base = int(body.get("base_rev") or 0)
                    if rec and base != cur_rev and rec.get("by") != by and not body.get("force"):
                        return self.send_json(
                            {"conflict": True, "current": {"v": rec["v"], "rev": cur_rev, "by": rec.get("by", ""), "at": rec.get("at", "")}},
                            409,
                        )
                    v = clean_answers(body.get("v"))
                    m["tp"][ref] = {"v": v, "rev": cur_rev + 1, "by": by, "at": now_iso()}
                    save_market(m)
                    return self.send_json(
                        {"ok": True, "rev": cur_rev + 1, "missing": missing(v), "at": m["tp"][ref]["at"], "progress": progress(m), "ready": ready_to_submit(m)}
                    )
                ref = mm.group(3)
                if ref not in WS_REFS:
                    return self.send_json({"error": "Unknown workstream"}, 404)
                m["owners"][ref] = {
                    "value": s(body.get("value"), 300),
                    "comment": s(body.get("comment"), 4000),
                    "by": by,
                    "at": now_iso(),
                }
                save_market(m)
                return self.send_json({"ok": True, "progress": progress(m), "ready": ready_to_submit(m)})

        mm = re.fullmatch(r"/api/m/([a-z0-9-]+)/presence", path)
        if mm:
            m = self.market_auth(mm.group(1), q)
            if not m:
                return
            by = s(body.get("by"), 80)
            if by:
                PRESENCE.setdefault(m["slug"], {})[by] = (s(body.get("ref"), 4), time.time())
            return self.send_json({"ok": True})

        mm = re.fullmatch(r"/api/m/([a-z0-9-]+)/submit", path)
        if mm:
            m = self.market_auth(mm.group(1), q)
            if not m:
                return
            with LOCK:
                m = load_market(m["slug"])
                if not ready_to_submit(m):
                    return self.send_json({"error": "Open items remain", "progress": progress(m)}, 409)
                by, note = s(body.get("by"), 80), s(body.get("note"), 1000)
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                fname = f"{m['slug']}_{stamp}.xlsx"
                folder = SUBS / m["slug"]
                folder.mkdir(parents=True, exist_ok=True)
                content = build_xlsx(m)
                (folder / fname).write_bytes(content)
                mailed = send_mail(m, fname, content, by, note)
                m["submissions"].append({"at": now_iso(), "by": by, "note": note, "file": fname, "mailed": mailed})
                save_market(m)
            return self.send_json({"ok": True, "file": fname, "mailed": mailed})

        self.send_json({"error": "Not found"}, 404)


def main():
    DATA.mkdir(exist_ok=True)
    SUBS.mkdir(exist_ok=True)
    cfg = get_config()
    if len(sys.argv) >= 3 and sys.argv[1] == "add":
        m = create_market(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "EUR")
        host = os.environ.get("PUBLIC_URL", "http://localhost:8080")
        print(f"Market created: {m['name']}\nLink: {host}/m/{m['slug']}?k={m['key']}")
        return
    port = int(os.environ.get("PORT", "8080"))
    host = os.environ.get("HOST", "0.0.0.0")
    print(f"Serving on http://localhost:{port}")
    print(f"Admin overview: http://localhost:{port}/admin?a={cfg['admin_key']}")
    ThreadingHTTPServer((host, port), Handler).serve_forever()


if __name__ == "__main__":
    main()
