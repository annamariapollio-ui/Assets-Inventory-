"""
VI Rollout Strategy - Market Inventory survey: Vercel serverless backend.

One function, routed by the `route` query parameter (/api/index?route=...&slug=...).
Storage : Upstash Redis (Vercel Marketplace) via REST; in-memory fallback for local preview.
Output  : the original Excel template, filled and e-mailed on submission.
"""
import base64
import hmac
import io
import json
import os
import re
import secrets
import smtplib
import time
import urllib.request
from datetime import datetime, timezone
from email.message import EmailMessage
from http.server import BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_NAME = "NEW_PMI_Market_Quantitative_Workbook.xlsx"

# ------------------------------------------------------------------ settings
REQUIRE_COMPLETE = os.environ.get("REQUIRE_COMPLETE", "false").lower() == "true"
PRESENCE_TTL = 40  # seconds a contributor stays shown as active

REDIS_URL = os.environ.get("KV_REST_API_URL") or os.environ.get("UPSTASH_REDIS_REST_URL")
REDIS_TOKEN = os.environ.get("KV_REST_API_TOKEN") or os.environ.get("UPSTASH_REDIS_REST_TOKEN")

# ---------------------------------------------------------------- vocabulary
EXISTS = ("yes", "no", "unknown")
YESNO = ("yes", "no")
AUDIENCE = ("internal", "external", "both")
VISIBILITY = ("low", "medium", "high")
EFFORT = ("easy / no-cost change", "some effort", "meaningful investment", "unknown")
EFFORT_WITH_COST = ("some effort", "meaningful investment")
BANDS = ("1", "2 to 5", "6 to 20", "21 to 50", "51 to 100", "101 or more")
QTY_MODES = ("exact", "band", "na")
IDENTITIES = ("Official corporate logo", "Locally adapted corporate logo", "Previous / legacy mark", "No mark", "Other")
SUPPORT_TAGS = ("Central guidance", "Creative or agency help", "Technical capacity", "Funding", "Procurement", "Training")
CONSTRAINT_TAGS = ("Budget timing", "Contract terms", "Release window", "Refurbishment cycle", "Approvals", "Procurement lead time")


# ------------------------------------------------------------------- storage
class Mem:
    """Tiny in-memory stand-in for the Redis commands used (local preview only)."""

    def __init__(self):
        self.d = {}

    def run(self, c):
        op, k = c[0].upper(), (c[1] if len(c) > 1 else None)
        d = self.d
        if op == "GET":
            return d.get(k)
        if op == "SET":
            d[k] = c[2]
            return "OK"
        if op == "INCR":
            d[k] = int(d.get(k, 0)) + 1
            return d[k]
        if op == "HSET":
            h = d.setdefault(k, {})
            for i in range(2, len(c), 2):
                h[c[i]] = c[i + 1]
            return 1
        if op == "HGET":
            return d.get(k, {}).get(c[2])
        if op == "HGETALL":
            out = []
            for f, v in d.get(k, {}).items():
                out += [f, v]
            return out
        if op == "SADD":
            d.setdefault(k, set()).add(c[2])
            return 1
        if op == "SMEMBERS":
            return sorted(d.get(k, set()))
        if op == "RPUSH":
            d.setdefault(k, []).append(c[2])
            return len(d[k])
        if op == "LRANGE":
            return list(d.get(k, []))
        raise ValueError("Unsupported command " + op)


_MEM = Mem()


def pipe(cmds):
    """Run several Redis commands in a single round trip."""
    if not REDIS_URL:
        return [_MEM.run(c) for c in cmds]
    req = urllib.request.Request(
        REDIS_URL.rstrip("/") + "/pipeline",
        data=json.dumps(cmds).encode(),
        headers={"Authorization": "Bearer " + REDIS_TOKEN, "Content-Type": "application/json", "User-Agent": "vi-survey/1.0"},
    )
    with urllib.request.urlopen(req, timeout=8) as r:
        out = json.loads(r.read())
    res = []
    for o in out:
        if o.get("error"):
            raise RuntimeError(o["error"])
        res.append(o.get("result"))
    return res


def cmd(*c):
    return pipe([list(c)])[0]


def todict(flat):
    if isinstance(flat, dict):
        return flat
    flat = flat or []
    return {flat[i]: flat[i + 1] for i in range(0, len(flat), 2)}


def K(kind, slug):
    return f"vi:{kind}:{slug}"


# ------------------------------------------------------------ template shape
def template_path():
    for base in (ROOT, Path.cwd(), Path(__file__).resolve().parent):
        p = base / "template" / TEMPLATE_NAME
        if p.exists():
            return p
    raise FileNotFoundError(TEMPLATE_NAME)


def load_structure():
    ws = load_workbook(template_path())["Inventory"]
    out, cur = [], None
    for r in range(6, 45):
        ref, title = ws.cell(r, 1).value, ws.cell(r, 2).value
        if not ref or not title:
            continue
        ref, title = str(ref).strip(), str(title).strip()
        if len(ref) == 1:
            cur = {"ref": ref, "title": title.split(". ", 1)[1] if ". " in title else title, "row": r, "touchpoints": []}
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
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-") or "market"


def s(v, n=2000):
    return str(v or "").strip()[:n]


def num(v):
    try:
        x = float(str(v).replace(",", "").replace("'", "").strip())
        return x if x >= 0 else None
    except (ValueError, TypeError):
        return None


def clean_answers(v):
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
    elif any(not a.get("identity") or (a["identity"] == "Other" and not a.get("identity_other")) for a in assets):
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


def progress(tp, owners):
    done = sum(1 for ref in ALL_REFS if ref in tp and not missing(tp[ref]["v"]))
    owners_done = sum(1 for w in WS_REFS if (owners.get(w, {}).get("value") or "").strip())
    return {"done": done, "total": len(ALL_REFS), "owners_done": owners_done, "owners_total": len(WS_REFS)}


def is_ready(p):
    return p["done"] == p["total"] and p["owners_done"] == p["owners_total"]


# ------------------------------------------------------------- Excel export
def fmt_money(x):
    return f"{x:,.0f}" if float(x).is_integer() else f"{x:,.2f}"


def identity_text(a):
    return a.get("identity_other", "") if a.get("identity") == "Other" else a["identity"].lower()


def join_tags(tags, text):
    return "; ".join(list(tags) + ([text] if text else []))


def build_xlsx(name, currency, tp, owners):
    """Fill the original template (Inventory tab only); Overview and Recap stay untouched."""
    wb = load_workbook(template_path())
    ws = wb["Inventory"]
    ws["B2"] = name
    for w in STRUCT:
        o = owners.get(w["ref"], {})
        if o.get("value"):
            ws.cell(w["row"], 12).value = o["value"]
        if o.get("comment"):
            ws.cell(w["row"], 15).value = o["comment"]
        for t in w["touchpoints"]:
            rec = tp.get(t["ref"])
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
                    ws.cell(r, 6).value = "; ".join(f'{a["name"]}: {identity_text(a)}' for a in assets if a.get("identity"))
                mode = v.get("qty_mode")
                if mode == "exact" and num(v.get("qty_num")) is not None:
                    ws.cell(r, 7).value = f'{fmt_money(num(v["qty_num"]))} {v.get("qty_unit", "")}'.strip()
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
                        ws.cell(r, 11).value = f"{currency} {fmt_money(lo)}\u2013{fmt_money(hi)}"
                    elif lo is not None:
                        ws.cell(r, 11).value = f"{currency} from {fmt_money(lo)}"
                    elif hi is not None:
                        ws.cell(r, 11).value = f"{currency} up to {fmt_money(hi)}"
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


# --------------------------------------------------------------------- email
def send_mail(name, filename, content, by, note, open_count):
    """Returns a short status string; the submission is stored either way."""
    to = os.environ.get("MAIL_TO", "").strip()
    if not to:
        return "not configured"
    subject = f"VI Rollout - market inventory submitted: {name}"
    text = (
        f"The market inventory for {name} was submitted by {by or 'a contributor'}.\n"
        + (f"Note: {note}\n" if note else "")
        + (f"{open_count} touchpoint(s) were still open at submission.\n" if open_count else "All touchpoints were answered.\n")
        + "The completed workbook is attached."
    )
    try:
        key = os.environ.get("RESEND_API_KEY")
        if key:
            payload = {
                "from": os.environ.get("MAIL_FROM", "VI Survey <onboarding@resend.dev>"),
                "to": [x.strip() for x in to.split(",") if x.strip()],
                "subject": subject,
                "text": text,
                "attachments": [{"filename": filename, "content": base64.b64encode(content).decode()}],
            }
            req = urllib.request.Request(
                "https://api.resend.com/emails",
                data=json.dumps(payload).encode(),
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json", "User-Agent": "vi-survey/1.0"},
            )
            urllib.request.urlopen(req, timeout=20).read()
            return "sent"
        host = os.environ.get("SMTP_HOST")
        if host:
            msg = EmailMessage()
            msg["Subject"], msg["To"] = subject, to
            msg["From"] = os.environ.get("MAIL_FROM", os.environ.get("SMTP_USER", "noreply@localhost"))
            msg.set_content(text)
            msg.add_attachment(
                content, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename
            )
            with smtplib.SMTP(host, int(os.environ.get("SMTP_PORT", "587")), timeout=20) as sm:
                sm.starttls()
                if os.environ.get("SMTP_USER"):
                    sm.login(os.environ["SMTP_USER"], os.environ.get("SMTP_PASS", ""))
                sm.send_message(msg)
            return "sent"
        return "not configured"
    except Exception as e:
        return "failed: " + str(e)[:120]


# ------------------------------------------------------------------ data ops
# Markets that exist from the first request (fixed link, no admin step needed).
# The access key can be overridden with the SWITZERLAND_KEY environment variable.
SEEDS = {
    "switzerland": {
        "slug": "switzerland",
        "name": "Switzerland",
        "currency": "CHF",
        "key": os.environ.get("SWITZERLAND_KEY", "swiss-preview-2026"),
    },
}


def get_meta(slug):
    if not re.fullmatch(r"[a-z0-9-]+", slug or ""):
        return None
    raw = cmd("GET", K("m", slug))
    seed = SEEDS.get(slug)
    if raw:
        meta = json.loads(raw)
        if seed:
            meta["key"] = seed["key"]
        return meta
    if seed:
        meta = dict(seed, created=now_iso())
        pipe([["SET", K("m", slug), json.dumps(meta)], ["SADD", "vi:markets", slug]])
        return meta
    return None


def load_all(slug):
    res = pipe([["HGETALL", K("tp", slug)], ["HGETALL", K("own", slug)], ["LRANGE", K("subs", slug), 0, -1], ["GET", K("ver", slug)]])
    tp = {k: json.loads(v) for k, v in todict(res[0]).items()}
    owners = {k: json.loads(v) for k, v in todict(res[1]).items()}
    subs = [json.loads(x) for x in (res[2] or [])]
    return tp, owners, subs, int(res[3] or 0)


def presence_list(slug, by, ref):
    cmds = [["GET", K("ver", slug)]]
    now = time.time()
    if by:
        cmds.append(["HSET", K("pres", slug), by, f"{ref}|{now}"])
    cmds.append(["HGETALL", K("pres", slug)])
    res = pipe(cmds)
    out = []
    for name, val in todict(res[-1]).items():
        try:
            r, ts = str(val).split("|")
            if now - float(ts) <= PRESENCE_TTL:
                out.append({"by": name, "ref": r})
        except ValueError:
            continue
    return int(res[0] or 0), out


def create_market(name, currency):
    name = (name or "").strip()[:80]
    if not name:
        raise ValueError("Market name needed")
    slug, n = slugify(name), 2
    while cmd("GET", K("m", slug)):
        slug = f"{slugify(name)}-{n}"
        n += 1
    meta = {"slug": slug, "name": name, "currency": (currency or "EUR").strip().upper()[:8], "key": secrets.token_urlsafe(8), "created": now_iso()}
    pipe([["SET", K("m", slug), json.dumps(meta)], ["SADD", "vi:markets", slug]])
    return meta


def full_state(meta, since_ver=None):
    slug = meta["slug"]
    tp, owners, subs, ver = load_all(slug)
    p = progress(tp, owners)
    return {
        "market": {"name": meta["name"], "slug": slug, "currency": meta["currency"]},
        "structure": STRUCT,
        "owners": owners,
        "tp": {ref: {"v": r["v"], "rev": r["rev"], "by": r.get("by", ""), "at": r.get("at", ""), "missing": missing(r["v"])} for ref, r in tp.items()},
        "progress": p,
        "ready": is_ready(p),
        "require_complete": REQUIRE_COMPLETE,
        "submissions": [{"at": x["at"], "by": x.get("by", "")} for x in subs],
        "ver": ver,
        "storage": "redis" if REDIS_URL else "memory",
        "options": {"identities": IDENTITIES, "support": SUPPORT_TAGS, "constraints": CONSTRAINT_TAGS, "bands": BANDS},
    }


XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ---------------------------------------------------------------------- HTTP
class handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > 2_000_000:
            return {}
        try:
            return json.loads(self.rfile.read(n) or b"{}")
        except json.JSONDecodeError:
            return {}

    def _origin(self):
        return f"{self.headers.get('X-Forwarded-Proto', 'http')}://{self.headers.get('Host', 'localhost')}"

    def _market(self, q):
        meta = get_meta(q.get("slug", ""))
        if not meta:
            self._json({"error": "Market not found"}, 404)
            return None
        key = self.headers.get("X-Key") or q.get("k", "")
        if not hmac.compare_digest(str(key), meta["key"]):
            self._json({"error": "Link not recognised"}, 403)
            return None
        return meta

    def _admin(self, q):
        expected = os.environ.get("ADMIN_KEY")
        if not expected:
            self._json({"error": "ADMIN_KEY is not set in the project settings"}, 503)
            return False
        if not hmac.compare_digest(str(self.headers.get("X-Admin") or q.get("a", "")), expected):
            self._json({"error": "Admin key not recognised"}, 403)
            return False
        return True

    def do_GET(self):
        self._route("GET")

    def do_POST(self):
        self._route("POST")

    def _route(self, method):
        q = {k: v[0] for k, v in parse_qs(urlparse(self.path).query).items()}
        route = q.get("route", "")
        try:
            self._dispatch(method, route, q)
        except Exception as e:  # keep the client informed instead of a blank 500
            self._json({"error": "Server error", "detail": str(e)[:200]}, 500)

    def _dispatch(self, method, route, q):
        # ---------------- admin
        if route == "admin_markets":
            if not self._admin(q):
                return
            if method == "POST":
                b = self._body()
                try:
                    meta = create_market(b.get("name"), b.get("currency"))
                except ValueError as e:
                    return self._json({"error": str(e)}, 400)
                return self._json({"ok": True, "slug": meta["slug"]})
            for sd in SEEDS:
                get_meta(sd)
            slugs = cmd("SMEMBERS", "vi:markets") or []
            cmds = []
            for sl in slugs:
                cmds += [["GET", K("m", sl)], ["HGETALL", K("tp", sl)], ["HGETALL", K("own", sl)], ["LRANGE", K("subs", sl), 0, -1]]
            res = pipe(cmds) if cmds else []
            out = []
            for i, sl in enumerate(slugs):
                raw = res[i * 4]
                if not raw:
                    continue
                meta = json.loads(raw)
                tp = {k: json.loads(v) for k, v in todict(res[i * 4 + 1]).items()}
                owners = {k: json.loads(v) for k, v in todict(res[i * 4 + 2]).items()}
                subs = [json.loads(x) for x in (res[i * 4 + 3] or [])]
                p = progress(tp, owners)
                times = [r.get("at", "") for r in tp.values()] + [o.get("at", "") for o in owners.values()]
                out.append(
                    {
                        "slug": sl, "name": meta["name"], "currency": meta["currency"],
                        "link": f"{self._origin()}/m/{sl}?k={meta['key']}",
                        "progress": p, "ready": is_ready(p),
                        "last_activity": max([t for t in times if t] or [meta.get("created", "")]),
                        "submissions": subs,
                    }
                )
            return self._json({"markets": out, "mail_to": os.environ.get("MAIL_TO", ""), "require_complete": REQUIRE_COMPLETE})

        if route in ("admin_export", "admin_file"):
            if not self._admin(q):
                return
            meta = get_meta(q.get("slug", ""))
            if not meta:
                return self._json({"error": "Market not found"}, 404)
            if route == "admin_file":
                raw = cmd("GET", f"vi:snap:{meta['slug']}:{int(q.get('n', 0))}")
                if not raw:
                    return self._json({"error": "Submission not found"}, 404)
                snap = json.loads(raw)
                tp, owners, fname = snap["tp"], snap["owners"], f"{meta['slug']}_submission_{q.get('n')}.xlsx"
            else:
                tp, owners, _, _ = load_all(meta["slug"])
                fname = f"{meta['slug']}_draft.xlsx"
            return self._send(200, build_xlsx(meta["name"], meta["currency"], tp, owners), XLSX, {"Content-Disposition": f'attachment; filename="{fname}"'})

        # ---------------- market
        meta = self._market(q)
        if not meta:
            return
        slug = meta["slug"]

        if route == "state" and method == "GET":
            ver, pres = presence_list(slug, s(q.get("by"), 80), s(q.get("ref"), 4))
            if q.get("since") and int(q["since"]) == ver and ver:
                return self._json({"unchanged": True, "ver": ver, "presence": pres})
            st = full_state(meta)
            st["presence"] = pres
            return self._json(st)

        if route == "export" and method == "GET":
            tp, owners, _, _ = load_all(slug)
            return self._send(200, build_xlsx(meta["name"], meta["currency"], tp, owners), XLSX, {"Content-Disposition": f'attachment; filename="{slug}_draft.xlsx"'})

        if route == "tp" and method == "POST":
            ref, b = q.get("ref", ""), self._body()
            if ref not in ALL_REFS:
                return self._json({"error": "Unknown touchpoint"}, 404)
            by = s(b.get("by"), 80)
            raw = cmd("HGET", K("tp", slug), ref)
            rec = json.loads(raw) if raw else None
            cur_rev = rec["rev"] if rec else 0
            if rec and int(b.get("base_rev") or 0) != cur_rev and rec.get("by") != by and not b.get("force"):
                return self._json({"conflict": True, "current": {"v": rec["v"], "rev": cur_rev, "by": rec.get("by", ""), "at": rec.get("at", "")}}, 409)
            v = clean_answers(b.get("v"))
            new = {"v": v, "rev": cur_rev + 1, "by": by, "at": now_iso()}
            res = pipe([["HSET", K("tp", slug), ref, json.dumps(new)], ["INCR", K("ver", slug)], ["HGETALL", K("tp", slug)], ["HGETALL", K("own", slug)]])
            tp = {k: json.loads(x) for k, x in todict(res[2]).items()}
            owners = {k: json.loads(x) for k, x in todict(res[3]).items()}
            p = progress(tp, owners)
            return self._json({"ok": True, "rev": new["rev"], "missing": missing(v), "at": new["at"], "progress": p, "ready": is_ready(p)})

        if route == "owner" and method == "POST":
            ref, b = q.get("ref", ""), self._body()
            if ref not in WS_REFS:
                return self._json({"error": "Unknown workstream"}, 404)
            rec = {"value": s(b.get("value"), 300), "comment": s(b.get("comment"), 4000), "by": s(b.get("by"), 80), "at": now_iso()}
            res = pipe([["HSET", K("own", slug), ref, json.dumps(rec)], ["INCR", K("ver", slug)], ["HGETALL", K("tp", slug)], ["HGETALL", K("own", slug)]])
            tp = {k: json.loads(x) for k, x in todict(res[2]).items()}
            owners = {k: json.loads(x) for k, x in todict(res[3]).items()}
            p = progress(tp, owners)
            return self._json({"ok": True, "progress": p, "ready": is_ready(p)})

        if route == "submit" and method == "POST":
            b = self._body()
            tp, owners, subs, _ = load_all(slug)
            p = progress(tp, owners)
            if REQUIRE_COMPLETE and not is_ready(p):
                return self._json({"error": "Open items remain", "progress": p}, 409)
            by, note = s(b.get("by"), 80), s(b.get("note"), 1000)
            content = build_xlsx(meta["name"], meta["currency"], tp, owners)
            n = int(cmd("INCR", f"vi:subn:{slug}"))
            fname = f"{slug}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            cmd("SET", f"vi:snap:{slug}:{n}", json.dumps({"tp": tp, "owners": owners}))
            open_count = p["total"] - p["done"]
            status = send_mail(meta["name"], fname, content, by, note, open_count)
            pipe([["RPUSH", K("subs", slug), json.dumps({"n": n, "at": now_iso(), "by": by, "note": note, "file": fname, "mail": status, "open": open_count})], ["INCR", K("ver", slug)]])
            return self._json({"ok": True, "mail": status})

        self._json({"error": "Not found"}, 404)
