# ==============================================================================
# تقرير توثيق الهدر — Waste Documentation Report — Template 1660942
# Groups waste submissions by area manager (location-ID → branch → AM roster)
# Cards: waste list, reason, box photos, after-Foodics photos, missing-photo flags
# Handles both form layouts (new: wastage box / Foodics, old: Item / Quantity)
# Branches that submitted in the lookback window but not on the report date → "لم يرسل"
# Uploads PDF to Google Drive and emails the link (same flow as cancelled orders)
# ==============================================================================
import os, io, re, html, base64, smtplib, asyncio
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import quote

import requests
from PIL import Image
from playwright.async_api import async_playwright
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError

# ==============================================================================
# CONFIG
# ==============================================================================
ZENPUT_API_KEY       = os.environ["ZENPUT_API_KEY"]
GMAIL_PASS           = os.environ["GMAIL_PASS"]
GOOGLE_CLIENT_ID     = os.environ["GOOGLE_CLIENT_ID"]
GOOGLE_CLIENT_SECRET = os.environ["GOOGLE_CLIENT_SECRET"]
GOOGLE_REFRESH_TOKEN = os.environ["GOOGLE_REFRESH_TOKEN"]
DRIVE_FOLDER_ID      = os.environ.get("DRIVE_FOLDER_ID_WASTE", "").strip()
DRIVE_FOLDER_NAME    = "Waste Documentation Reports"

GMAIL_USER = "aof.group.auto@gmail.com"
TO_EMAIL   = ["o.salahaddin@aofgroup.com", "m.alhuaydar@aofgroup.com", "s.alharbi@aofgroup.com"]
CC_EMAIL   = ["a.alsalem@aofgroup.com"]

TEMPLATE_ID    = 1660942
ZENPUT_BASE    = "https://www.zenput.com"
PAGE_SIZE      = 100
MAX_PAGES      = 15
LOOKBACK_DAYS  = 14      # branches active in this window are "expected" to submit
MAX_THUMBS     = 6       # per photo strip; the rest shown as +N
PHOTO_MAX_W    = 520
PHOTO_WORKERS  = 8
PAGE_W_PX      = 1150
PAGE_H_PX      = 1626    # A4 ratio at 1150px wide

KSA         = ZoneInfo("Asia/Riyadh")
REPORT_DATE = (os.environ.get("REPORT_DATE", "").strip()
               or datetime.now(KSA).strftime("%Y-%m-%d"))
LOOKBACK_START = (datetime.strptime(REPORT_DATE, "%Y-%m-%d") - timedelta(days=LOOKBACK_DAYS)).strftime("%Y-%m-%d")

print(f"Report date: {REPORT_DATE}  |  lookback from {LOOKBACK_START}")

# ==============================================================================
# BRANCH REFERENCE — Zenput location ID → label, branch code → area manager
# ==============================================================================
LOC_LABEL = {
    2155652: "NURUH B01", 2155654: "AQRUH B13", 2164013: "KHRUH B02", 2164014: "GHRUH B03",
    2164016: "RWRUH B05", 2164017: "DARUH B06", 2164019: "SWRUH B08", 2164020: "AZRUH B09",
    2164021: "SHRUH B10", 2164022: "NRRUH B11", 2164023: "TWRUH B12", 2164025: "RBRUH B14",
    2164026: "NDRUH B15", 2164027: "BDRUH B16", 2164028: "QRRUH B17", 2164030: "MURUH B19",
    2164031: "SFJED B24", 2164032: "KRRUH B21", 2185452: "OBJED B22", 2190657: "SLAHS B23",
    2197297: "NSRUH B04", 2197298: "TKRUH B18", 2197299: "LBRUH B07", 2199002: "RWAHS B25",
    2199835: "HAJED B26", 2203271: "SARUH B27", 2210205: "MAJED B28", 2211854: "QARUH B30",
    2235670: "ANRUH B31", 2239240: "FYJED B32", 2242934: "HIRJED B33", 2243963: "URRUH B34",
    2250799: "IRRUH B35", 2257790: "SHWMAK B37", 2258220: "PSJED B36", 2260889: "UHDMM B38",
    2263062: "HSRUH B39", 2281339: "MZDMM B40",
    2169459: "Lubda Alaqeq LB01", 2222802: "Lubda Alkhaleej LB02",
    2232755: "Garatis QB01", 2235805: "Garatis QB02", 2254072: "Garatis QB03",
    2256386: "Garatis QB04", 2268360: "Garatis QB05", 2270650: "Garatis QB06",
    2274188: "Garatis QB07", 2276794: "Garatis QB08",
    2171883: "Twesste TW01",
}
EXCLUDE_LOCS = {2175245, 2230615, 2256173}   # factory, central kitchen, warehouse

AM_ROSTER = {
    "Adnan":              ["B21", "B18", "QB05", "QB07", "QB04", "B13", "B17"],
    "ALI ESMAIL":         ["B15", "B01", "B35", "B34", "TW01", "B04", "B39"],
    "M.emad":             ["B26", "B32", "B22", "B37"],
    "Islam":              ["B14", "B11", "B30", "B19", "B31", "QB02", "QB03"],
    "NAIF ALZAHRANI":     ["B16", "B12", "B07", "B06"],
    "Meqdad Ali":         ["QB06", "B02", "B27", "B05"],
    "Naresh Joshe":       ["QB01", "B09", "B10", "B08", "QB08"],
    "abdullah al ghanmi": ["LB01", "LB02"],
    "Salah Aldeen":       ["B23", "B25"],
    "suhail":             ["B38", "B40"],
    "Sujan":              ["B28", "B24", "B33", "B36"],
}
AM_BY_CODE = {code: am for am, codes in AM_ROSTER.items() for code in codes}
UNASSIGNED = "غير محدد"

CODE_RE = re.compile(r"\b(QB\d{2}|LB\d{2}|TW\d{2}|B\d{2})\b", re.I)

def clean(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").replace("\t", " ")).strip()

def code_sort_key(code: str):
    m = re.match(r"([A-Z]+)(\d+)", code or "")
    return (m.group(1), int(m.group(2))) if m else ("ZZ", 999)

# ==============================================================================
# 1. FETCH — paginate with offset, filter dates client-side, dedupe by id
# ==============================================================================
def local_dt(s: dict):
    m   = s.get("smetadata") or {}
    raw = m.get("date_submitted_local") or ""
    if len(raw) >= 16:
        return raw[:10], raw[11:16]
    try:
        d = datetime.fromisoformat((m.get("date_submitted") or "").replace("Z", "+00:00")).astimezone(KSA)
        return d.strftime("%Y-%m-%d"), d.strftime("%H:%M")
    except Exception:
        return "", ""

def fetch_submissions() -> list:
    headers = {"X-API-TOKEN": ZENPUT_API_KEY, "Accept": "application/json"}
    found, offset = {}, 0
    for page in range(MAX_PAGES):
        params = {"form_template_id": TEMPLATE_ID, "limit": PAGE_SIZE,
                  "offset": offset, "start": offset}
        r = requests.get(f"{ZENPUT_BASE}/api/v3/submissions/", headers=headers,
                         params=params, timeout=60)
        r.raise_for_status()
        batch = r.json().get("data", []) or []
        if not batch:
            break
        oldest, kept = None, 0
        for s in batch:
            d, _ = local_dt(s)
            if d and (oldest is None or d < oldest):
                oldest = d
            if d and LOOKBACK_START <= d <= REPORT_DATE and s.get("id") not in found:
                found[s["id"]] = s
                kept += 1
        print(f"  page {page}: {len(batch)} rows, kept {kept}, oldest {oldest}")
        if len(batch) < PAGE_SIZE or (oldest and oldest < LOOKBACK_START):
            break
        offset += PAGE_SIZE
    print(f"  Fetched {len(found)} submissions in window")
    return list(found.values())

# ==============================================================================
# 2. PARSE
# ==============================================================================
def find_ans(answers: list, pred) -> dict:
    for a in answers:
        if pred((a.get("title") or "").strip().lower()):
            return a
    return {}

def text_val(a: dict) -> str:
    v = a.get("value") if a else None
    if v is None:
        return ""
    return v.strip() if isinstance(v, str) else str(v).strip()

def photo_keys(a: dict) -> list:
    v = (a.get("value") if a else None) or []     # read `value` only (image_value duplicates it)
    if isinstance(v, dict):
        v = [v]
    return [p["s3_key"] for p in v if isinstance(p, dict) and p.get("s3_key")]

def combine_old(item_txt: str, qty_txt: str) -> str:
    items = [l.strip() for l in item_txt.splitlines() if l.strip()]
    qtys  = [l.strip() for l in qty_txt.splitlines() if l.strip()]
    if items and len(items) == len(qtys):
        return "\n".join(f"{i} — {q}" for i, q in zip(items, qtys))
    parts = []
    if item_txt: parts.append(f"الصنف: {item_txt}")
    if qty_txt:  parts.append(f"الكمية: {qty_txt}")
    return "\n".join(parts)

def parse(s: dict):
    m   = s.get("smetadata") or {}
    ans = s.get("answers") or []

    loc_a  = find_ans(ans, lambda t: t.startswith("location"))
    loc_id = None
    try:
        loc_id = int(str(loc_a.get("value")).strip())
    except Exception:
        mloc = m.get("location")
        if isinstance(mloc, dict):
            try: loc_id = int(mloc.get("id"))
            except Exception: pass
    if loc_id in EXCLUDE_LOCS:
        return None

    mloc_name = m.get("location", {}).get("name") if isinstance(m.get("location"), dict) else ""
    if loc_id in LOC_LABEL:
        label = LOC_LABEL[loc_id]
        code  = label.split()[-1].upper()
    else:
        label = clean(loc_a.get("display_value") or mloc_name)
        cm    = CODE_RE.search(label)
        code  = cm.group(1).upper() if cm else ""
    am = AM_BY_CODE.get(code, UNASSIGNED)

    box_a     = find_ans(ans, lambda t: "wastage box" in t and "picture" in t and "foodics" not in t)
    foodics_a = find_ans(ans, lambda t: "foodics" in t)
    list_a    = find_ans(ans, lambda t: "write all the wastage" in t)
    reason_a  = find_ans(ans, lambda t: "waste reason" in t)
    item_a    = find_ans(ans, lambda t: "/ item" in t)
    oldph_a   = find_ans(ans, lambda t: "waste photo" in t)
    qty_a     = find_ans(ans, lambda t: "/ quantity" in t)
    prep_a    = find_ans(ans, lambda t: t.startswith("prepared by"))

    layout = "old" if (item_a or oldph_a or qty_a) and not (list_a or box_a) else "new"
    if layout == "new":
        items_txt  = text_val(list_a)
        box_keys   = photo_keys(box_a)
        foodics_ks = photo_keys(foodics_a)
    else:
        items_txt  = combine_old(text_val(item_a), text_val(qty_a))
        box_keys   = photo_keys(oldph_a)
        foodics_ks = []

    date, time_ = local_dt(s)
    return {
        "id":           s.get("id"),
        "date":         date,
        "time":         time_,
        "loc_id":       loc_id,
        "branch":       label or "—",
        "code":         code,
        "am":           am,
        "prepared_by":  text_val(prep_a),
        "reason":       text_val(reason_a),
        "items":        items_txt,
        "layout":       layout,
        "box_keys":     box_keys,
        "foodics_keys": foodics_ks,
        "no_box":       not box_keys,
        "no_foodics":   layout == "new" and not foodics_ks,
        "no_items":     not items_txt,
    }

# ==============================================================================
# 3. PHOTOS — s3_key → signed URL → resized base64 (same run as the API pull)
# ==============================================================================
def get_signed_url(s3_key: str) -> str:
    url = f"{ZENPUT_BASE}/api/v2/users/current/storage/?path={quote(s3_key, safe='')}"
    try:
        r = requests.get(url, headers={"X-API-TOKEN": ZENPUT_API_KEY}, timeout=15)
        if r.status_code == 200:
            return (r.json().get("data") or {}).get("location", "")
        print(f"  ⚠️  sign HTTP {r.status_code}: {s3_key[-40:]}")
    except Exception as e:
        print(f"  ⚠️  sign error: {e}")
    return ""

def fetch_photo_b64(s3_key: str) -> str:
    signed = get_signed_url(s3_key)
    if not signed:
        return ""
    try:
        r = requests.get(signed, timeout=25, headers={"User-Agent": "Mozilla/5.0"})
        if r.status_code == 200 and len(r.content) > 500:
            img = Image.open(io.BytesIO(r.content)).convert("RGB")
            if img.width > PHOTO_MAX_W:
                img = img.resize((PHOTO_MAX_W, int(img.height * PHOTO_MAX_W / img.width)), Image.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=62, optimize=True)
            return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()
        print(f"  ⚠️  photo HTTP {r.status_code}: {s3_key[-40:]}")
    except Exception as e:
        print(f"  ⚠️  photo error: {e}")
    return ""

def fetch_all_photos(records: list) -> dict:
    keys = sorted({k for r in records for k in r["box_keys"] + r["foodics_keys"]})
    print(f"Fetching {len(keys)} photos...")
    with ThreadPoolExecutor(max_workers=PHOTO_WORKERS) as ex:
        results = list(ex.map(fetch_photo_b64, keys))
    print(f"  OK {sum(1 for v in results if v)} / {len(keys)}")
    return dict(zip(keys, results))

# ==============================================================================
# 4. HTML
# ==============================================================================
CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Cairo', 'Segoe UI', sans-serif; background: #f0f2f5; color: #2d3436; padding: 16px; }
.container { width: 100%; max-width: 1070px; margin: 0 auto; }

.header { background: linear-gradient(135deg, #f39c12 0%, #d35400 100%); color: white;
          border-radius: 14px; padding: 28px 40px; margin-bottom: 24px; text-align: center; }
.header h1 { font-size: 28px; font-weight: 700; margin-bottom: 6px; }
.header .sub { font-size: 14px; opacity: 0.9; }

.kpi-strip { display: flex; gap: 12px; margin-bottom: 24px; justify-content: center; }
.kpi-card { background: white; border-radius: 10px; padding: 16px 18px; text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.07); border-top: 4px solid #e67e22; flex: 1; }
.kpi-card.green { border-top-color: #00b894; }
.kpi-card.red   { border-top-color: #d63031; }
.kpi-card.blue  { border-top-color: #0984e3; }
.kpi-card.gray  { border-top-color: #636e72; }
.kpi-label { font-size: 12px; color: #636e72; margin-bottom: 6px; }
.kpi-value { font-size: 28px; font-weight: 700; }

.section-title { display: flex; align-items: center; gap: 14px; margin: 28px 0 16px; }
.section-title h2 { font-size: 19px; font-weight: 700; white-space: nowrap; padding: 6px 18px;
                    border-radius: 20px; color: white; background: #d35400; }
.section-title.red h2 { background: #d63031; }
.section-title::after { content: ''; flex: 1; height: 2px; background: #e0e0e0; }

.missing-box { background: white; border-radius: 12px; padding: 14px 20px;
               box-shadow: 0 2px 8px rgba(0,0,0,0.08); page-break-inside: avoid; }
.missing-row { display: flex; gap: 12px; align-items: center; padding: 8px 0; border-bottom: 1px solid #f1f2f6; }
.missing-row:last-child { border-bottom: none; }
.missing-am { font-weight: 700; min-width: 170px; font-size: 14px; }
.chip { display: inline-block; background: #ffeaea; color: #c0392b; border-radius: 12px;
        padding: 3px 12px; font-size: 12px; font-weight: 700; margin: 2px; }
.all-ok { text-align: center; color: #00b894; font-weight: 700; padding: 10px; }

.am-block { background: white; border-radius: 12px; margin-bottom: 22px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08); overflow: hidden; }
.am-header { background: linear-gradient(90deg, #f39c12, #d35400); color: white; padding: 13px 22px;
             display: flex; justify-content: space-between; align-items: center; }
.am-header h3 { font-size: 16px; font-weight: 700; }
.am-count { background: rgba(255,255,255,0.25); border-radius: 20px; padding: 3px 14px;
            font-size: 13px; font-weight: 600; }

.cards-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; padding: 18px; }
.card { border: 1px solid #eee; border-radius: 10px; overflow: hidden; background: #fafafa;
        break-inside: avoid; page-break-inside: avoid; }
.card-info { padding: 12px 14px; }
.card-branch { font-size: 15px; font-weight: 700; margin-bottom: 2px; }
.card-meta { font-size: 11.5px; color: #8395a7; margin-bottom: 8px; }
.tag-old { display: inline-block; background: #dfe6e9; color: #2d3436; font-size: 10px;
           font-weight: 700; padding: 1px 8px; border-radius: 8px; margin-right: 6px; }
.box-label { font-size: 11.5px; color: #636e72; font-weight: 700; margin: 8px 0 3px; }
.items-box { background: #fffaf0; border-right: 3px solid #f39c12; border-radius: 4px;
             padding: 7px 10px; font-size: 12.5px; white-space: pre-wrap; word-break: break-word; }
.reason-box { background: #fff5f5; border-right: 3px solid #e17055; border-radius: 4px;
              padding: 6px 10px; font-size: 12px; color: #c0392b; font-weight: 600;
              white-space: pre-wrap; word-break: break-word; }
.empty-txt { color: #b2bec3; font-weight: 400; }
.flags { margin-top: 8px; }
.flag { display: inline-block; background: #d63031; color: white; border-radius: 10px;
        padding: 2px 10px; font-size: 11px; font-weight: 700; margin: 2px 0 2px 4px; }

.strip { border-top: 1px solid #eee; padding: 8px 10px 10px; }
.strip-label { font-size: 11px; font-weight: 700; color: #856404; background: #fff3cd;
               border-radius: 6px; padding: 3px 8px; margin-bottom: 6px; display: inline-block; }
.strip.foodics .strip-label { color: #0c5460; background: #d1ecf1; }
.thumbs { display: grid; grid-template-columns: repeat(3, 1fr); gap: 5px; }
.thumbs img, .thumb-fail, .thumb-more { width: 100%; height: 112px; object-fit: cover;
                                        border-radius: 6px; display: block; }
.thumb-fail, .thumb-more { display: flex; align-items: center; justify-content: center;
                           background: #ecf0f1; color: #95a5a6; font-size: 11px; }
.thumb-more { font-size: 18px; font-weight: 700; color: #636e72; }
.strip-empty { font-size: 12px; color: #d63031; font-weight: 700; }

/* ── pagination: never split a card, never leave a header alone at a page bottom ── */
.header, .kpi-strip, .missing-row { break-inside: avoid; page-break-inside: avoid; }
.section-title, .am-header { break-after: avoid; page-break-after: avoid; }
.am-block { overflow: visible; }
.am-header { border-radius: 12px 12px 0 0; }
@media print { body { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }
"""

def strip_html(keys: list, photos: dict, label: str, cls: str, missing_text: str) -> str:
    if not keys:
        return (f'<div class="strip {cls}"><span class="strip-label">{label}</span>'
                f'<div class="strip-empty">{missing_text}</div></div>')
    shown = keys[:MAX_THUMBS]
    extra = len(keys) - len(shown)
    cells = []
    for k in shown:
        b = photos.get(k, "")
        cells.append(f'<img src="{b}">' if b else '<div class="thumb-fail">تعذر التحميل</div>')
    if extra > 0:
        cells[-1] = f'<div class="thumb-more">+{extra + 1}</div>'
    return (f'<div class="strip {cls}"><span class="strip-label">{label} ({len(keys)})</span>'
            f'<div class="thumbs">{"".join(cells)}</div></div>')

def card_html(r: dict, photos: dict) -> str:
    old_tag = '<span class="tag-old">نموذج قديم</span>' if r["layout"] == "old" else ""
    items = (html.escape(r["items"]) if r["items"]
             else '<span class="empty-txt">لم يتم تسجيل الهدر</span>')
    reason = (html.escape(r["reason"]) if r["reason"]
              else '<span class="empty-txt">لا يوجد</span>')
    flags = []
    if r["no_box"]:     flags.append("بدون صورة الهدر")
    if r["no_foodics"]: flags.append("بدون صورة فوديكس")
    if r["no_items"]:   flags.append("بدون تسجيل الكميات")
    flags_html = ('<div class="flags">' + "".join(f'<span class="flag">⚠️ {f}</span>' for f in flags)
                  + "</div>") if flags else ""
    prep = f" &nbsp;|&nbsp; 👤 {html.escape(r['prepared_by'])}" if r["prepared_by"] else ""

    strips = strip_html(r["box_keys"], photos, "📦 صور الهدر بالبوكس", "box", "لا توجد صور")
    if r["layout"] == "new":
        strips += strip_html(r["foodics_keys"], photos, "🧾 بعد التسجيل على فوديكس", "foodics",
                             "لا توجد صورة بعد التسجيل على فوديكس")
    return f"""
    <div class="card">
        <div class="card-info">
            <div class="card-branch">📍 {html.escape(r["branch"])} {old_tag}</div>
            <div class="card-meta">🕐 {html.escape(r["time"])}{prep}</div>
            <div class="box-label">الهدر والكميات</div>
            <div class="items-box">{items}</div>
            <div class="box-label">سبب الهدر</div>
            <div class="reason-box">{reason}</div>
            {flags_html}
        </div>
        {strips}
    </div>"""

def am_order(name: str):
    return (name == UNASSIGNED, name.lower())

def build_html(day: list, missing: dict, photos: dict, kpis: dict) -> str:
    kpi_html = f"""
    <div class="kpi-strip">
        <div class="kpi-card"><div class="kpi-label">عدد النماذج</div><div class="kpi-value">{kpis['subs']}</div></div>
        <div class="kpi-card green"><div class="kpi-label">فروع أرسلت</div><div class="kpi-value">{kpis['sent']}</div></div>
        <div class="kpi-card red"><div class="kpi-label">فروع لم ترسل</div><div class="kpi-value">{kpis['missing']}</div></div>
        <div class="kpi-card blue"><div class="kpi-label">بدون صورة فوديكس</div><div class="kpi-value">{kpis['no_foodics']}</div></div>
        <div class="kpi-card gray"><div class="kpi-label">إجمالي الصور</div><div class="kpi-value">{kpis['photos']}</div></div>
    </div>"""

    if missing:
        rows = "".join(
            f'<div class="missing-row"><div class="missing-am">{html.escape(am)}</div><div>'
            + "".join(f'<span class="chip">{html.escape(b)}</span>' for b in branches)
            + "</div></div>"
            for am, branches in sorted(missing.items(), key=lambda x: am_order(x[0])))
        missing_html = f'<div class="missing-box">{rows}</div>'
    else:
        missing_html = '<div class="missing-box"><div class="all-ok">✅ جميع الفروع أرسلت نموذج الهدر</div></div>'

    groups = {}
    for r in day:
        groups.setdefault(r["am"], []).append(r)
    blocks = ""
    for am in sorted(groups, key=am_order):
        recs = sorted(groups[am], key=lambda r: (code_sort_key(r["code"]), r["time"]))
        n_br = len({r["code"] or r["branch"] for r in recs})
        cards = "".join(card_html(r, photos) for r in recs)
        blocks += f"""
        <div class="am-block">
            <div class="am-header"><h3>{html.escape(am)}</h3>
                <span class="am-count">{len(recs)} نموذج · {n_br} فرع</span></div>
            <div class="cards-grid">{cards}</div>
        </div>"""
    if not day:
        blocks = "<p style='text-align:center;color:#b2bec3;padding:40px'>لا توجد نماذج هدر لهذا اليوم</p>"

    return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<title>تقرير توثيق الهدر — {REPORT_DATE}</title>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>تقرير توثيق الهدر</h1>
        <div class="sub">{REPORT_DATE} &nbsp;|&nbsp; {kpis['subs']} نموذج &nbsp;|&nbsp; {kpis['sent']} فرع</div>
    </div>
    {kpi_html}
    <div class="section-title red"><h2>فروع لم ترسل نموذج الهدر</h2></div>
    {missing_html}
    <div class="section-title"><h2>تفاصيل الهدر حسب مدير المنطقة</h2></div>
    {blocks}
</div>
</body>
</html>"""

# ==============================================================================
# 5. PDF VIA PLAYWRIGHT
# ==============================================================================
WAIT_IMAGES_JS = """async () => {
    const imgs = Array.from(document.images);
    await Promise.all(imgs.map(img => {
        if (img.complete && img.naturalWidth > 0) return img.decode().catch(() => {});
        return new Promise(res => { img.onload = img.onerror = res; })
                   .then(() => img.decode().catch(() => {}));
    }));
    const broken = imgs.filter(i => !i.naturalWidth).length;
    return [imgs.length, broken];
}"""

async def build_pdf(html_content: str):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, args=["--no-sandbox"])
        page    = await browser.new_page(viewport={"width": PAGE_W_PX, "height": 1200})
        await page.set_content(html_content, wait_until="networkidle")

        # stretch the viewport over the whole report so every photo gets painted
        height = int(await page.evaluate("() => document.body.scrollHeight"))
        await page.set_viewport_size({"width": PAGE_W_PX, "height": height + 100})
        total, broken = await page.evaluate(WAIT_IMAGES_JS)
        print(f"  images in page: {total}, broken: {broken}")
        await page.wait_for_timeout(1500)

        # normal A4-shaped pages → Drive can preview it
        pdf = await page.pdf(width=f"{PAGE_W_PX}px", height=f"{PAGE_H_PX}px",
                             print_background=True,
                             margin={"top": "24px", "bottom": "24px",
                                     "left": "20px", "right": "20px"})
        await browser.close()
    return pdf, height

# ==============================================================================
# 6. GOOGLE DRIVE
# ==============================================================================
def get_drive_service():
    creds = Credentials(
        token=None,
        refresh_token=GOOGLE_REFRESH_TOKEN,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=GOOGLE_CLIENT_ID,
        client_secret=GOOGLE_CLIENT_SECRET,
        scopes=["https://www.googleapis.com/auth/drive.file"],
    )
    creds.refresh(Request())
    return build("drive", "v3", credentials=creds, cache_discovery=False)

def get_or_create_folder(svc) -> str:
    q = (f"name = '{DRIVE_FOLDER_NAME}' and mimeType = 'application/vnd.google-apps.folder' "
         f"and trashed = false")
    found = svc.files().list(q=q, fields="files(id, name)", spaces="drive").execute().get("files", [])
    if found:
        return found[0]["id"]
    folder = svc.files().create(body={"name": DRIVE_FOLDER_NAME,
                                      "mimeType": "application/vnd.google-apps.folder"},
                                fields="id").execute()
    print(f"  📁 Created Drive folder '{DRIVE_FOLDER_NAME}' → {folder['id']}")
    return folder["id"]

def upload_to_drive(pdf_bytes: bytes, filename: str) -> str:
    svc = get_drive_service()

    def create(parent_id: str):
        media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype="application/pdf", resumable=False)
        return svc.files().create(body={"name": filename, "parents": [parent_id]},
                                  media_body=media, fields="id, webViewLink",
                                  supportsAllDrives=True).execute()
    try:
        up = create(DRIVE_FOLDER_ID or get_or_create_folder(svc))
    except HttpError as e:
        if DRIVE_FOLDER_ID and e.resp.status == 404:
            print("  ⚠️  DRIVE_FOLDER_ID not visible to this app (drive.file scope) — using auto folder")
            up = create(get_or_create_folder(svc))
        else:
            raise
    svc.permissions().create(fileId=up["id"], body={"type": "anyone", "role": "reader"},
                             supportsAllDrives=True).execute()
    print(f"  ☁️  Uploaded → {up['webViewLink']}")
    return up["webViewLink"]

# ==============================================================================
# 7. WASTAGE RANKING — parse free-text quantities, rank branches & items
# ==============================================================================
TOP_N_BRANCHES = 5
TOP_N_ITEMS    = 5

AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
NUM_RE    = re.compile(r"\d+(?:[.,]\d+)?")
_LETTERS  = "A-Za-z\u0600-\u06FF"

def _alt(words):
    return "|".join(sorted((re.escape(w) for w in words), key=len, reverse=True))

def _tok(words):
    return re.compile(rf"(?<![{_LETTERS}])(?:{_alt(words)})(?![{_LETTERS}])", re.I)

KG_WORDS = ["kg", "kgs", "kilo", "kilos", "kilogram", "kilograms", "kig", "kgm",
            "كيلو", "كيلوجرام", "كيلوغرام", "كجم", "كغ", "كغم", "كلغ", "كيلوات"]
G_WORDS  = ["g", "gm", "gms", "gr", "grm", "grms", "gram", "grams", "gramm",
            "غ", "غم", "غرام", "جرام", "جم", "جرامات", "غرامات"]
PC_WORDS = ["pcs", "pc", "pce", "pces", "psc", "pic", "pics", "piece", "pieces", "peses",
            "pese", "pices", "peace", "peaces", "حبة", "حبه", "حبات", "قطعة", "قطع"]
FILLER   = ["هدر", "waste", "wastage"]

KG_RE, G_RE, PC_RE, FILLER_RE = _tok(KG_WORDS), _tok(G_WORDS), _tok(PC_WORDS), _tok(FILLER)

def _unit_at(text: str):
    t = text.lstrip(" -=:/'\".,")
    if KG_RE.match(t): return "kg"
    if G_RE.match(t):  return "g"
    if PC_RE.match(t): return "pcs"
    return None

def parse_waste_lines(text: str) -> list:
    """One line = one item. First number = quantity.
    Unit: word right after the number, else a unit word anywhere on the line, else pieces.
    Grams are converted to kg so weights add up in one unit."""
    out = []
    for raw in (text or "").translate(AR_DIGITS).splitlines():
        line = raw.strip(" -=:/'\"\t.,—")
        if not line or line.startswith(("الصنف:", "الكمية:")):
            continue
        m = NUM_RE.search(line)
        qty, unit = None, None
        if m:
            qty  = float(m.group().replace(",", "."))
            unit = _unit_at(line[m.end():])
        if unit is None:
            if KG_RE.search(line):   unit = "kg"
            elif G_RE.search(line):  unit = "g"
            else:                    unit = "pcs"
        if unit == "g":
            unit = "kg"
            if qty is not None:
                qty = qty / 1000.0
        name = NUM_RE.sub(" ", line)
        for rx in (KG_RE, G_RE, PC_RE, FILLER_RE):
            name = rx.sub(" ", name)
        name = re.sub(r"[=\-/:'\"(),.—]+", " ", name)
        name = re.sub(r"\s+", " ", name).strip()
        if name:
            out.append({"name": name, "qty": qty, "unit": unit, "raw": raw.strip()})
    return out

def item_key(name: str) -> str:
    words = [w[:-1] if len(w) > 3 and w.endswith("s") else w for w in name.lower().split()]
    return " ".join(words)

def fmt_num(x: float) -> str:
    return f"{x:,.2f}".rstrip("0").rstrip(".")

def compute_rankings(day: list):
    branches, items = {}, {}
    for r in day:
        key = r["code"] or r["branch"]
        b = branches.setdefault(key, {"branch": r["branch"], "am": r["am"], "pcs": 0.0,
                                      "kg": 0.0, "lines": 0, "subs": 0, "items": {}})
        b["subs"] += 1
        for w in parse_waste_lines(r["items"]):
            b["lines"] += 1
            b.setdefault("parsed", []).append(w)
            k = item_key(w["name"])
            bi = b["items"].setdefault(k, {"name": w["name"], "pcs": 0.0, "kg": 0.0})
            it = items.setdefault(k, {"name": w["name"], "pcs": 0.0, "kg": 0.0, "branches": set()})
            it["branches"].add(key)
            if w["qty"] is not None:
                b[w["unit"]]  += w["qty"]
                bi[w["unit"]] += w["qty"]
                it[w["unit"]] += w["qty"]
    top_b = sorted(branches.values(), key=lambda b: (b["pcs"], b["kg"], b["lines"]), reverse=True)
    top_i = sorted(items.values(), key=lambda i: (i["pcs"], i["kg"], len(i["branches"])), reverse=True)
    totals = {"pcs": sum(b["pcs"] for b in branches.values()),
              "kg":  sum(b["kg"]  for b in branches.values())}
    return top_b[:TOP_N_BRANCHES], top_i[:TOP_N_ITEMS], totals

# ==============================================================================
# 8. EMAIL — wastage summary + Drive link
# ==============================================================================
TD = "padding:8px 10px;border-bottom:1px solid #eee;text-align:right;vertical-align:top"
TH = "padding:8px 10px;background:#d35400;color:white;text-align:right;font-weight:700"

def qty_cell(pcs: float, kg: float) -> str:
    parts = []
    if pcs: parts.append(f"<strong>{fmt_num(pcs)}</strong> حبة")
    if kg:  parts.append(f"<strong>{fmt_num(kg)}</strong> كجم")
    return " + ".join(parts) or "—"

def send_email(link: str, kpis: dict, top_b: list, top_i: list, totals: dict):
    subject = f"تقرير توثيق الهدر — {REPORT_DATE}"

    if top_b:
        rows = ""
        for n, b in enumerate(top_b, 1):
            by_pcs = sorted((i for i in b["items"].values() if i["pcs"]), key=lambda i: i["pcs"], reverse=True)[:3]
            by_kg  = sorted((i for i in b["items"].values() if i["kg"]),  key=lambda i: i["kg"],  reverse=True)[:2]
            lines  = [f"{html.escape(i['name'])} ({fmt_num(i['pcs'])} حبة)" for i in by_pcs]
            lines += [f"{html.escape(i['name'])} ({fmt_num(i['kg'])} كجم)"  for i in by_kg]
            items_txt = "<br>".join(lines) or "—"
            bg = "background:#fff5eb;" if n == 1 else ""
            rows += (f'<tr style="{bg}"><td style="{TD};font-weight:700">{n}</td>'
                     f'<td style="{TD};font-weight:700" dir="ltr">{html.escape(b["branch"])}</td>'
                     f'<td style="{TD}">{html.escape(b["am"])}</td>'
                     f'<td style="{TD}">{qty_cell(b["pcs"], b["kg"])}</td>'
                     f'<td style="{TD};font-size:12px" dir="ltr">{items_txt}</td></tr>')
        branches_tbl = (f'<table style="border-collapse:collapse;width:100%;max-width:760px;font-size:13px">'
                        f'<tr><th style="{TH}">#</th><th style="{TH}">الفرع</th><th style="{TH}">مدير المنطقة</th>'
                        f'<th style="{TH}">إجمالي الهدر</th><th style="{TH}">أعلى الأصناف</th></tr>{rows}</table>')
    else:
        branches_tbl = "<p>لا توجد نماذج هدر لهذا اليوم.</p>"

    if top_i:
        rows = "".join(
            f'<tr><td style="{TD}" dir="ltr">{html.escape(i["name"])}</td>'
            f'<td style="{TD}">{qty_cell(i["pcs"], i["kg"])}</td>'
            f'<td style="{TD}">{len(i["branches"])}</td></tr>' for i in top_i)
        items_tbl = (f'<table style="border-collapse:collapse;width:100%;max-width:760px;font-size:13px">'
                     f'<tr><th style="{TH}">الصنف</th><th style="{TH}">الكمية</th>'
                     f'<th style="{TH}">عدد الفروع</th></tr>{rows}</table>')
    else:
        items_tbl = ""

    body = f"""
<div dir="rtl" style="font-family:Cairo,Arial,sans-serif;font-size:14px;color:#2d3436;line-height:1.8">
  <p>السلام عليكم،</p>
  <p>ملخص الهدر ليوم <strong>{REPORT_DATE}</strong>:</p>
  <p style="margin:6px 0">
    إجمالي الهدر المسجل: {qty_cell(totals['pcs'], totals['kg'])}
    &nbsp;|&nbsp; {kpis['subs']} نموذج من {kpis['sent']} فرع
    &nbsp;|&nbsp; <span style="color:#c0392b">{kpis['missing']} فرع لم يرسل</span>
  </p>
  <br>
  <p style="color:#d35400;font-weight:700;font-size:15px">أعلى {len(top_b)} فروع في الهدر</p>
  {branches_tbl}
  <br>
  <p style="color:#d35400;font-weight:700;font-size:15px">أكثر الأصناف هدراً</p>
  {items_tbl}
  <br>
  <p><a href="{link}" style="display:inline-block;background:#d35400;color:white;padding:12px 28px;
        border-radius:8px;text-decoration:none;font-weight:700;font-size:15px;">عرض التقرير التفصيلي مع الصور</a></p>
  <p style="color:#636e72;font-size:11.5px">* الكميات مستخرجة من النص المكتوب في النموذج وقد تكون تقريبية — التفاصيل والصور في التقرير.</p>
  <br>
  <p>Business Intelligence<br>AOF Group</p>
</div>"""
    msg = MIMEMultipart("mixed")
    msg["From"]    = f"Business Intelligence <{GMAIL_USER}>"
    msg["To"]      = ", ".join(TO_EMAIL)
    msg["Cc"]      = ", ".join(CC_EMAIL)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "html", "utf-8"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_PASS)
        server.sendmail(GMAIL_USER, TO_EMAIL + CC_EMAIL, msg.as_string())
    print(f"✅ Email sent → {', '.join(TO_EMAIL)}, CC: {', '.join(CC_EMAIL)}")

# ==============================================================================
# MAIN
# ==============================================================================
def main():
    raw = fetch_submissions()
    records = [p for p in (parse(s) for s in raw) if p]
    if not records:
        print("⚠️  No submissions in the whole lookback window — exiting.")
        return

    day = [r for r in records if r["date"] == REPORT_DATE]

    # expected = branches active in lookback; missing = expected − submitted on report date
    expected = {}
    for r in records:
        expected[r["code"] or r["branch"]] = (r["branch"], r["am"])
    sent_keys = {r["code"] or r["branch"] for r in day}
    missing = {}
    for key in sorted(set(expected) - sent_keys, key=code_sort_key):
        label, am = expected[key]
        missing.setdefault(am, []).append(label)

    kpis = {
        "subs":       len(day),
        "sent":       len(sent_keys),
        "missing":    sum(len(v) for v in missing.values()),
        "no_foodics": sum(1 for r in day if r["no_foodics"]),
        "photos":     sum(len(r["box_keys"]) + len(r["foodics_keys"]) for r in day),
    }
    print(f"  Report day: {kpis}")
    print(f"  Layouts: new={sum(r['layout']=='new' for r in day)} old={sum(r['layout']=='old' for r in day)}")
    unassigned = sorted({r["branch"] for r in records if r["am"] == UNASSIGNED})
    if unassigned:
        print(f"  ⚠️  Branches with no area manager: {unassigned}")

    photos = fetch_all_photos(day)

    print("Building HTML...")
    html_content = build_html(day, missing, photos, kpis)

    print("Generating PDF...")
    pdf_bytes, height = asyncio.run(build_pdf(html_content))
    print(f"  PDF {len(pdf_bytes):,} bytes, page height {height}px")

    print("Uploading to Google Drive...")
    link = upload_to_drive(pdf_bytes, f"waste_documentation_{REPORT_DATE}.pdf")

    top_b, top_i, totals = compute_rankings(day)
    print(f"  Totals: {totals}")
    for b in top_b:
        print(f"  [{b['branch']}] pcs={fmt_num(b['pcs'])} kg={fmt_num(b['kg'])}")
        for w in b.get("parsed", []):
            q = "-" if w["qty"] is None else fmt_num(w["qty"])
            print(f"      {w['raw']!r:45} -> {w['name']} | {q} {w['unit']}")

    print("Sending email...")
    send_email(link, kpis, top_b, top_i, totals)

if __name__ == "__main__":
    main()
