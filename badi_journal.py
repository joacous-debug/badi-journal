# -*- coding: utf-8 -*-
"""
Badi-Journal – Betriebsjournal für das Schwimmbad.
Läuft auf dem Büro-PC, Zugriff per Browser vom PC (http://localhost:8080) und vom Tablet (http://<PC-IP>:8080).
Nur Python 3.9+ nötig, keine weiteren Pakete.
"""
import csv, datetime as dt, email, html, io, json, os, re, secrets, socket, sqlite3, sys, threading
from http import cookies
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlencode, urlparse

BASE = os.path.dirname(os.path.abspath(__file__))
CFG_PATH = os.path.join(BASE, "config.json")
DEFAULT_CFG = {
    "port": 8080,
    "badi": "Badi Elgg",
    "becken": "MZB",
    "admin_passwort": "",
    "visa": ["TF", "ER"],
    "orte": ["Büro", "Maschinenraum", "Beckenrand"],
    "grenzwerte": {
        "ph":     {"referenz": [7.0, 7.4], "toleranz": [6.8, 7.6]},
        "cl_frei": {"referenz": [0.2, 0.4], "toleranz": [0.2, 0.8]},
        "cl_geb": {"referenz": [0.0, 0.2], "toleranz": [0.0, 0.2]},
    },
    "runden": [["Morgen", "00:00", "12:00"], ["Mittag", "12:00", "16:00"], ["Abend", "16:00", "24:00"]],
    "backup_ordner": "backup",
    "backup_tage": 365,
    "csv_dezimalzeichen": ".",
}


def load_cfg():
    cfg = json.loads(json.dumps(DEFAULT_CFG))
    if os.path.exists(CFG_PATH):
        with open(CFG_PATH, encoding="utf-8") as f:
            cfg.update(json.load(f))
    else:
        with open(CFG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    return cfg


CFG = load_cfg()
DB_PATH = os.path.join(BASE, "badi_journal.db")
BACKUP_DIR = CFG["backup_ordner"] if os.path.isabs(CFG["backup_ordner"]) else os.path.join(BASE, CFG["backup_ordner"])

# ---------------------------------------------------------------- Felder
# key, Bezeichnung, CSV/Excel-Spalte, Typ, min, max, Nachkommastellen, Grenzwert
F = [
    ("hand_cl_frei", "Freies Chlor (mg/l)", "Hand: Freies Chlor", "num", 0, 5, 2, "cl_frei"),
    ("hand_cl_geb", "Geb. Chlor (mg/l)", "Hand: Geb. Chlor", "num", 0, 5, 2, "cl_geb"),
    ("hand_ph", "pH-Wert", "Hand: pH", "num", 5, 9, 2, "ph"),
    ("wassertemp", "Wassertemperatur °C", "Hand: Wassertemp °C", "num", 0, 45, 1, None),
    ("lufttemp", "Lufttemperatur °C", "Lufttemp °C", "num", -20, 50, 1, None),
    ("auto_cl_frei", "Freies Chlor (mg/l)", "Auto: Freies Chlor", "num", 0, 5, 2, "cl_frei"),
    ("auto_ph", "pH-Wert", "Auto: pH", "num", 5, 9, 2, "ph"),
    ("auto_redox", "Redox (mV)", "Auto: Redox mV", "num", 0, 1200, 0, None),
    ("frischwasser", "Frischwasser (m³/d)", "Frischwasser m3", "num", 0, 10000, 1, None),
    ("quelle", "Quelle (Std)", "Quelle Std", "num", 0, 24, 1, None),
    ("besucher", "Besucherzahl (Pers./d)", "Besucher", "num", 0, 100000, 0, None),
    ("rueckspuelung", "Filterrückspülung (Zeit)", "Filterrückspülung (Zeit)", "time", None, None, None, None),
    ("anschwemmung", "Anschwemmung (Zeit)", "Anschwemmung (Zeit)", "time", None, None, None, None),
    ("ph_stapeltank", "pH Stapeltank", "pH Stapeltank", "num", 5, 9, 2, None),
    ("bemerkungen", "Bemerkungen", "Bemerkungen", "text", None, None, None, None),
    ("massnahmen", "Massnahmen", "Massnahmen", "text", None, None, None, None),
    ("stoerungen", "Betriebsstörungen", "Betriebsstörungen", "text", None, None, None, None),
    ("kalibrierungen", "Kalibrierungen", "Kalibrierungen", "text", None, None, None, None),
]
FD = {f[0]: dict(zip(("key", "label", "csv", "typ", "min", "max", "dec", "limit"), f)) for f in F}
KEYS = [f[0] for f in F]
HAND = ["hand_cl_frei", "hand_cl_geb", "hand_ph", "wassertemp", "lufttemp"]
AUTO = ["auto_cl_frei", "auto_ph", "auto_redox"]
CSV_HEAD = ["Datum", "Zeit", "Ort", "Visum"] + [FD[k]["csv"] for k in KEYS] + ["ID", "Erfasst am"]

ERFASSUNG = {
    "swan": ("Swan-Messung", "Handmessung mit dem Swan Chematest", HAND, None),
    "maschinenraum": ("Maschinenraum", "Sensorwerte und Anlagen im Maschinenraum",
                      AUTO + ["ph_stapeltank", "frischwasser", "quelle", "rueckspuelung", "anschwemmung"], "Maschinenraum"),
    "tag": ("Tageswerte & Notizen", "Besucher, Bemerkungen, Massnahmen, Störungen, Kalibrierungen",
            ["besucher", "bemerkungen", "massnahmen", "stoerungen", "kalibrierungen"], "Büro"),
}
WOCHENTAGE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]

# ---------------------------------------------------------------- Datenbank
SCHEMA = f"""
CREATE TABLE IF NOT EXISTS messungen (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  datum TEXT NOT NULL, zeit TEXT NOT NULL, ort TEXT, visum TEXT,
  {", ".join(f"{k} {'REAL' if FD[k]['typ'] == 'num' else 'TEXT'}" for k in KEYS)},
  erfasst_am TEXT, geaendert_am TEXT);
CREATE INDEX IF NOT EXISTS ix_datum ON messungen(datum, zeit);
CREATE TABLE IF NOT EXISTS protokoll (
  id INTEGER PRIMARY KEY AUTOINCREMENT, zeit TEXT, aktion TEXT, messung_id INTEGER, client TEXT, daten TEXT);
"""
DB_LOCK = threading.Lock()


def db():
    con = sqlite3.connect(DB_PATH, timeout=10)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with db() as con:
        con.executescript(SCHEMA)


def now():
    return dt.datetime.now().replace(microsecond=0)


def log(con, aktion, mid, client, daten):
    con.execute("INSERT INTO protokoll(zeit,aktion,messung_id,client,daten) VALUES (?,?,?,?,?)",
                (now().isoformat(" "), aktion, mid, client, json.dumps(daten, ensure_ascii=False)))


def save_row(data, client, mid=None):
    cols = ["datum", "zeit", "ort", "visum"] + KEYS
    with DB_LOCK, db() as con:
        if mid:
            old = con.execute("SELECT * FROM messungen WHERE id=?", (mid,)).fetchone()
            con.execute(f"UPDATE messungen SET {', '.join(c + '=?' for c in cols)}, geaendert_am=? WHERE id=?",
                        [data.get(c) for c in cols] + [now().isoformat(" "), mid])
            log(con, "geändert", mid, client, {"vorher": dict(old) if old else None, "nachher": data})
        else:
            if data.get("id"):
                cur = con.execute(f"INSERT INTO messungen(id,{','.join(cols)},erfasst_am) VALUES (?,{','.join('?' * len(cols))},?)",
                                  [data["id"]] + [data.get(c) for c in cols] + [data.get("erfasst_am") or now().isoformat(" ")])
            else:
                cur = con.execute(f"INSERT INTO messungen({','.join(cols)},erfasst_am) VALUES ({','.join('?' * len(cols))},?)",
                                  [data.get(c) for c in cols] + [now().isoformat(" ")])
            mid = cur.lastrowid
            log(con, "erfasst", mid, client, data)
    backup()
    return mid


def delete_row(mid, client):
    with DB_LOCK, db() as con:
        old = con.execute("SELECT * FROM messungen WHERE id=?", (mid,)).fetchone()
        if old:
            con.execute("DELETE FROM messungen WHERE id=?", (mid,))
            log(con, "gelöscht", mid, client, dict(old))
    backup()


def rows_between(von, bis):
    with db() as con:
        return [dict(r) for r in con.execute(
            "SELECT * FROM messungen WHERE datum BETWEEN ? AND ? ORDER BY datum, zeit, id", (von.isoformat(), bis.isoformat()))]


# ---------------------------------------------------------------- CSV & Backup
def fmt_num(v, dec, sep="."):
    if v is None or v == "":
        return ""
    s = f"{v:.{dec}f}" if dec is not None else str(v)
    return s.replace(".", sep)


def to_csv(rows):
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";", lineterminator="\r\n")
    w.writerow(CSV_HEAD)
    sep = CFG["csv_dezimalzeichen"]
    for r in rows:
        d = dt.date.fromisoformat(r["datum"]).strftime("%d.%m.%Y")
        vals = [(f"{r[k]:g}".replace(".", sep) if FD[k]["typ"] == "num" and r[k] is not None else (r[k] or "")) for k in KEYS]
        w.writerow([d, r["zeit"], r["ort"] or "", r["visum"] or ""] + vals + [r["id"], r["erfasst_am"] or ""])
    return "﻿" + buf.getvalue()


BACKUP_STATUS = {"ok": True, "zeit": None, "fehler": ""}


def backup():
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        with db() as con:
            rows = [dict(r) for r in con.execute("SELECT * FROM messungen ORDER BY datum, zeit, id")]
        data = to_csv(rows)
        today = dt.date.today().isoformat()
        for name in ("messungen_aktuell.csv", f"messungen_{today}.csv"):
            tmp = os.path.join(BACKUP_DIR, name + ".tmp")
            with open(tmp, "w", encoding="utf-8", newline="") as f:
                f.write(data)
            os.replace(tmp, os.path.join(BACKUP_DIR, name))
        src = sqlite3.connect(DB_PATH)
        dst = sqlite3.connect(os.path.join(BACKUP_DIR, f"badi_journal_{today}.db"))
        with dst:
            src.backup(dst)
        src.close(); dst.close()
        limit = dt.date.today() - dt.timedelta(days=int(CFG["backup_tage"]))
        for fn in os.listdir(BACKUP_DIR):
            m = re.match(r"(?:messungen|badi_journal)_(\d{4}-\d{2}-\d{2})\.(csv|db)$", fn)
            if m and dt.date.fromisoformat(m.group(1)) < limit:
                os.remove(os.path.join(BACKUP_DIR, fn))
        BACKUP_STATUS.update(ok=True, zeit=now(), fehler="")
    except Exception as e:  # Backup darf das Speichern nie verhindern
        BACKUP_STATUS.update(ok=False, fehler=str(e))


def parse_date(s):
    s = (s or "").strip()
    for f in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def import_csv(text, client):
    text = text.lstrip("﻿")
    first = text.splitlines()[0] if text else ""
    delim = ";" if first.count(";") >= first.count(",") else ","
    rd = csv.DictReader(io.StringIO(text), delimiter=delim)
    by_csv = {FD[k]["csv"]: k for k in KEYS}
    neu = upd = 0
    fehler = []
    with db() as con:
        existing = {r[0] for r in con.execute("SELECT id FROM messungen")}
    for i, r in enumerate(rd, start=2):
        if not any((v or "").strip() for v in r.values() if isinstance(v, str)):
            continue
        form = {"datum": r.get("Datum", ""), "zeit": r.get("Zeit", ""), "ort": r.get("Ort", ""), "visum": r.get("Visum", "")}
        for col, k in by_csv.items():
            form[k] = r.get(col, "")
        data, errs = validate(form, KEYS, need_dt=True, strict=False)
        if errs:
            fehler.append(f"Zeile {i}: " + "; ".join(errs))
            continue
        rid = (r.get("ID") or "").strip()
        rid = int(rid) if rid.isdigit() else None
        if rid and rid in existing:
            save_row(data, client, mid=rid); upd += 1
        else:
            data["id"] = rid
            data["erfasst_am"] = r.get("Erfasst am") or None
            save_row(data, client); neu += 1
    return neu, upd, fehler


# ---------------------------------------------------------------- Validierung & Auswertung
def to_float(s):
    s = (s or "").strip().replace(",", ".").replace("'", "")
    return float(s) if s else None


def validate(form, keys, need_dt=False, strict=True):
    data, errs = {}, []
    if need_dt:
        d = parse_date(form.get("datum"))
        if not d:
            errs.append("Datum fehlt oder ungültig")
        data["datum"] = d.isoformat() if d else None
        z = (form.get("zeit") or "").strip()[:5]
        if not re.fullmatch(r"\d{1,2}:\d{2}", z):
            errs.append("Zeit fehlt oder ungültig (hh:mm)")
        data["zeit"] = z.zfill(5)
    data["ort"] = (form.get("ort") or "").strip() or None
    data["visum"] = (form.get("visum") or "").strip().upper() or None
    for k in keys:
        f, raw = FD[k], form.get(k, "")
        if f["typ"] == "num":
            try:
                v = to_float(raw)
            except ValueError:
                errs.append(f"{f['label']}: «{raw}» ist keine Zahl"); continue
            if v is not None and not (f["min"] <= v <= f["max"]):
                errs.append(f"{f['label']}: {raw} ist unplausibel (erlaubt {f['min']}–{f['max']})"); continue
            data[k] = v
        elif f["typ"] == "time":
            z = (raw or "").strip()[:5]
            if z and not re.fullmatch(r"\d{1,2}:\d{2}", z):
                errs.append(f"{f['label']}: Zeit im Format hh:mm"); continue
            data[k] = z.zfill(5) if z else None
        else:
            data[k] = (raw or "").strip() or None
    if strict and not data["visum"]:
        errs.append("Visum fehlt")
    if strict and all(data.get(k) is None for k in keys):
        errs.append("Kein Wert eingegeben")
    return data, errs


def status(key, v):
    lim = FD[key]["limit"]
    if v is None or v == "" or not lim:
        return ""
    g = CFG["grenzwerte"][lim]
    if not (g["toleranz"][0] <= v <= g["toleranz"][1]):
        return "bad"
    if not (g["referenz"][0] <= v <= g["referenz"][1]):
        return "warn"
    return "good"


def runde(zeit):
    for i, (_, a, b) in enumerate(CFG["runden"]):
        if a <= zeit < b or (b == "24:00" and zeit >= a):
            return i
    return len(CFG["runden"]) - 1


def monday_of(d):
    return d - dt.timedelta(days=d.weekday())


def week_model(monday, von=None, bis=None):
    days = [monday + dt.timedelta(days=i) for i in range(7)]
    rows = rows_between(days[0], days[-1])
    nR = len(CFG["runden"])
    inside = [(von is None or d >= von) and (bis is None or d <= bis) for d in days]
    by = {(d.isoformat(), s): [] for d in days for s in range(nR)}
    byday = {d.isoformat(): [] for d in days}
    for r in rows:
        by[(r["datum"], runde(r["zeit"]))].append(r)
        byday[r["datum"]].append(r)

    def avg(rs, k):
        v = [r[k] for r in rs if r[k] is not None]
        return sum(v) / len(v) if v else None

    def slot_rows(keys):
        out = []
        for k in keys:
            cells = []
            for di, d in enumerate(days):
                for s in range(nR):
                    v = avg(by[(d.isoformat(), s)], k) if inside[di] else None
                    cells.append((fmt_num(v, FD[k]["dec"]), status(k, v)))
            out.append((FD[k]["label"], cells, 1))
        vis = []
        for di, d in enumerate(days):
            for s in range(nR):
                rs = [r for r in by[(d.isoformat(), s)] if any(r[k] is not None for k in HAND)] if inside[di] else []
                vis.append((" ".join(dict.fromkeys(r["visum"] or "" for r in rs)).strip(), ""))
        out.append(("Visum", vis, 1))
        return out

    def day_rows(spec):
        out = []
        for k, how in spec:
            cells = []
            for di, d in enumerate(days):
                rs = [r for r in byday[d.isoformat()] if r[k] not in (None, "")] if inside[di] else []
                if not rs:
                    v = ""
                elif how == "sum":
                    v = fmt_num(sum(r[k] for r in rs), FD[k]["dec"])
                elif how == "avg":
                    v = fmt_num(avg(rs, k), FD[k]["dec"])
                else:
                    v = " · ".join(f"{r[k]}" for r in rs) if how == "list" else rs[-1][k]
                cells.append((v, ""))
            out.append((FD[k]["label"], cells, nR))
        return out

    auto = slot_rows(AUTO)
    auto.pop()  # kein Visum bei Automat
    sections = [
        ("Automat. Messung (Sensoren)", auto),
        ("Handmessung (Swan)", slot_rows(HAND)),
        ("Tageswerte", day_rows([("frischwasser", "sum"), ("quelle", "sum"), ("besucher", "sum"),
                                  ("rueckspuelung", "list"), ("anschwemmung", "list"), ("ph_stapeltank", "avg")])),
        ("Notizen", day_rows([("bemerkungen", "list"), ("massnahmen", "list"), ("stoerungen", "list"), ("kalibrierungen", "list")])),
    ]
    return days, inside, sections


# ---------------------------------------------------------------- HTML
E = lambda s: html.escape("" if s is None else str(s))

CSS = """
:root{--b:#1f4e78;--l:#ddebf7;--g:#e2efda;--w:#ffe699;--r:#f8b4b4;--t:#222;--m:#666;--bd:#b7b7b7}
*{box-sizing:border-box}body{margin:0;font-family:system-ui,Segoe UI,Arial,sans-serif;color:var(--t);background:#f4f6f9}
header{background:var(--b);color:#fff;padding:10px 16px;display:flex;gap:16px;align-items:center;flex-wrap:wrap}
header a{color:#fff;text-decoration:none;font-weight:600;padding:6px 10px;border-radius:6px}header a:hover,header a.on{background:#ffffff26}
header .t{font-size:1.15rem;font-weight:700;margin-right:auto}
main{padding:16px;max-width:1400px;margin:0 auto}
h1{font-size:1.4rem;margin:.2em 0 .6em}h2{font-size:1.1rem}
.btn{display:inline-block;border:0;background:var(--b);color:#fff;padding:10px 16px;border-radius:8px;font-size:1rem;text-decoration:none;cursor:pointer;font-weight:600}
.btn.sec{background:#fff;color:var(--b);border:1px solid var(--b)}.btn.del{background:#b42318}.btn.sm{padding:4px 9px;font-size:.85rem}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:16px;margin-bottom:24px}
.tile{display:block;background:#fff;border:2px solid var(--b);border-radius:14px;padding:26px 18px;text-decoration:none;color:var(--b)}
.tile b{display:block;font-size:1.5rem}.tile span{color:var(--m)}
form.f{background:#fff;border-radius:12px;padding:18px;max-width:640px}
.fld{margin-bottom:14px}.fld label{display:block;font-weight:600;margin-bottom:4px}
.fld input,.fld select,.fld textarea{width:100%;font-size:1.25rem;padding:10px;border:1px solid var(--bd);border-radius:8px;background:#fff}
.fld textarea{min-height:70px;font-size:1rem}.fld small{color:var(--m)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.msg{padding:12px 14px;border-radius:8px;margin-bottom:14px}.msg.ok{background:var(--g)}.msg.err{background:var(--r)}.msg.inf{background:var(--l)}
table{border-collapse:collapse;background:#fff}td,th{border:1px solid var(--bd);padding:4px 6px;font-size:.85rem;text-align:center}
th{background:var(--l)}td.lab{text-align:left;font-weight:600;background:var(--l);white-space:nowrap}
tr.sec td{background:var(--b);color:#fff;font-weight:700;text-align:left}
.good{}.warn{background:var(--w)!important}.bad{background:var(--r)!important;font-weight:700}
td.out{background:#eee}.wrap{overflow-x:auto}
.list td{text-align:right}.list td.l{text-align:left}
.bar{display:flex;gap:10px;flex-wrap:wrap;align-items:end;margin-bottom:14px}.bar .fld{margin:0}.bar input{font-size:1rem;padding:7px}
.lim{font-size:.8rem;color:var(--m)}
.wk{margin-bottom:22px}.wk h2{margin:0 0 4px}.wk .hd{display:flex;justify-content:space-between;align-items:end;gap:10px;flex-wrap:wrap}
.notes td{text-align:left;vertical-align:top;height:34px}
@media print{header,.noprint{display:none!important}body{background:#fff}main{padding:0;max-width:none}
 .wk{page-break-after:always;margin:0}.wk:last-child{page-break-after:auto}
 td,th{font-size:7.5pt;padding:2px 3px}table{width:100%}
 *{-webkit-print-color-adjust:exact;print-color-adjust:exact}@page{size:A4 landscape;margin:8mm}}
"""

JS_LIMITS = """
<script>
const LIM=%s;
document.querySelectorAll('input[data-lim]').forEach(el=>{const f=()=>{const g=LIM[el.dataset.lim];
 const v=parseFloat(el.value.replace(',','.'));el.classList.remove('warn','bad');if(isNaN(v))return;
 if(v<g.toleranz[0]||v>g.toleranz[1])el.classList.add('bad');else if(v<g.referenz[0]||v>g.referenz[1])el.classList.add('warn');};
 el.addEventListener('input',f);f();});
</script>"""


def page(title, body, admin=False, nav="", refresh=None):
    links = [("start", "/", "Erfassen"), ("woche", "/woche", "Woche")]
    if admin:
        links += [("admin", "/admin", "Messungen"), ("bericht", "/admin/bericht", "PDF-Bericht"),
                  ("daten", "/admin/daten", "Backup & CSV"), ("prot", "/admin/protokoll", "Protokoll")]
    else:
        links += [("login", "/admin", "PC-Bereich")]
    nl = "".join(f'<a href="{u}" class="{"on" if k == nav else ""}">{t}</a>' for k, u, t in links)
    bk = "" if BACKUP_STATUS["ok"] or not admin else f'<div class="msg err">Backup fehlgeschlagen: {E(BACKUP_STATUS["fehler"])}</div>'
    rf = f'<meta http-equiv="refresh" content="{refresh}">' if refresh else ""
    return f"""<!doctype html><html lang="de"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
{rf}<title>{E(title)} – {E(CFG['badi'])}</title><style>{CSS}</style></head><body>
<header><span class="t">{E(CFG['badi'])} · Becken {E(CFG['becken'])}</span>{nl}</header><main>{bk}{body}</main></body></html>"""


def limits_text():
    g = CFG["grenzwerte"]
    def rng(a): return f"≤ {a[1]}" if a[0] == 0 else f"{a[0]}–{a[1]}"
    return (f'<span class="lim">Referenz / Toleranz: pH {rng(g["ph"]["referenz"])} / {rng(g["ph"]["toleranz"])} · '
            f'Freies Chlor {rng(g["cl_frei"]["referenz"])} / {rng(g["cl_frei"]["toleranz"])} mg/l · '
            f'Geb. Chlor {rng(g["cl_geb"]["referenz"])} / {rng(g["cl_geb"]["toleranz"])} mg/l</span>')


def field_html(k, val=""):
    f = FD[k]
    v = "" if val is None else val
    if f["typ"] == "num" and isinstance(v, float):
        v = fmt_num(v, f["dec"])
    if f["typ"] == "text":
        inp = f'<textarea name="{k}">{E(v)}</textarea>'
    elif f["typ"] == "time":
        inp = f'<input type="time" name="{k}" value="{E(v)}">'
    else:
        lim = f' data-lim="{f["limit"]}"' if f["limit"] else ""
        inp = f'<input type="text" inputmode="decimal" autocomplete="off" name="{k}" value="{E(v)}"{lim}>'
    return f'<div class="fld"><label>{E(f["label"])}</label>{inp}</div>'


def visum_html(val=""):
    opts = "".join(f'<option value="{E(v)}">' for v in CFG["visa"])
    return (f'<div class="fld"><label>Visum</label><input name="visum" list="visa" value="{E(val)}" autocomplete="off" '
            f'style="text-transform:uppercase"><datalist id="visa">{opts}</datalist></div>')


def ort_html(val=""):
    opts = "".join(f'<option {"selected" if o == val else ""}>{E(o)}</option>' for o in CFG["orte"])
    return f'<div class="fld"><label>Ort</label><select name="ort">{opts}</select></div>'


def week_table(days, inside, sections, compact=False):
    nR = len(CFG["runden"])
    h1 = "".join(f'<th colspan="{nR}" class="{"" if inside[i] else "out"}">{WOCHENTAGE[i]}<br>{d.strftime("%d.%m.%Y")}</th>'
                 for i, d in enumerate(days))
    h2 = "".join(f'<th class="{"" if inside[i] else "out"}">{E(r[0])}</th>' for i in range(7) for r in CFG["runden"])
    body = ""
    for title, rows in sections:
        body += f'<tr class="sec"><td colspan="{1 + 7 * nR}">{E(title)}</td></tr>'
        notes = title == "Notizen"
        for label, cells, span in rows:
            tds = ""
            for i, (v, st) in enumerate(cells):
                di = i * span // nR
                cls = st or ("out" if not inside[di] else "")
                tds += f'<td colspan="{span}" class="{cls}">{E(v)}</td>'
            body += f'<tr class="{"notes" if notes else ""}"><td class="lab">{E(label)}</td>{tds}</tr>'
    return f'<div class="wrap"><table style="width:100%"><tr><th></th>{h1}</tr><tr><th></th>{h2}</tr>{body}</table></div>'


def week_block(monday, von=None, bis=None):
    days, inside, sections = week_model(monday, von, bis)
    iy, iw, _ = (monday + dt.timedelta(days=3)).isocalendar()
    return (f'<section class="wk"><div class="hd"><h2>Betriebsprotokoll · Becken {E(CFG["becken"])} · {iy} · KW {iw}</h2>'
            f'{limits_text()}</div>{week_table(days, inside, sections)}</section>')


# ---------------------------------------------------------------- HTTP
SESSIONS = set()


class H(BaseHTTPRequestHandler):
    server_version = "BadiJournal/1.0"

    def log_message(self, fmt, *a):
        pass

    # --- helpers
    def client(self):
        return self.client_address[0]

    def is_admin(self):
        if self.client() in ("127.0.0.1", "::1"):
            return True
        c = cookies.SimpleCookie(self.headers.get("Cookie", ""))
        return "sid" in c and c["sid"].value in SESSIONS

    def send(self, body, code=200, ctype="text/html; charset=utf-8", headers=None):
        b = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(b)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(b)

    def redirect(self, url, headers=None):
        self.send_response(303)
        self.send_header("Location", url)
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()

    def form(self):
        n = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(n)
        ct = self.headers.get("Content-Type", "")
        if ct.startswith("multipart/form-data"):
            msg = email.message_from_bytes(b"Content-Type: " + ct.encode() + b"\r\n\r\n" + raw)
            out = {}
            for part in msg.get_payload():
                name = part.get_param("name", header="content-disposition")
                out[name] = part.get_payload(decode=True)
            return out
        return {k: v[0] for k, v in parse_qs(raw.decode("utf-8"), keep_blank_values=True).items()}

    def q(self):
        return {k: v[0] for k, v in parse_qs(urlparse(self.path).query).items()}

    # --- routing
    def do_GET(self):
        self.route("GET")

    def do_POST(self):
        self.route("POST")

    def route(self, method):
        p = urlparse(self.path).path.rstrip("/") or "/"
        try:
            if p == "/":
                return self.start()
            if p.startswith("/erfassen/"):
                return self.erfassen(p.split("/")[2], method)
            if p == "/woche":
                return self.woche()
            if p == "/login":
                return self.login(method)
            if p.startswith("/admin"):
                if not self.is_admin():
                    return self.redirect("/login")
                parts = p.split("/")[2:]
                sub = parts[0] if parts else ""
                if sub == "":
                    return self.admin_list()
                if sub == "neu":
                    return self.admin_edit(None, method)
                if sub == "edit":
                    return self.admin_edit(int(parts[1]), method)
                if sub == "delete" and method == "POST":
                    delete_row(int(parts[1]), self.client())
                    return self.redirect("/admin?" + urlencode(self.q()) + "&msg=Messung+gel%C3%B6scht")
                if sub == "bericht":
                    return self.bericht()
                if sub == "export.csv":
                    return self.export()
                if sub == "daten":
                    return self.daten(method)
                if sub == "protokoll":
                    return self.protokoll()
            self.send(page("Nicht gefunden", "<h1>Seite nicht gefunden</h1>"), 404)
        except Exception as e:
            self.send(page("Fehler", f'<div class="msg err">Fehler: {E(e)}</div>'), 500)

    # --- Tablet / alle
    def start(self):
        q = self.q()
        msg = f'<div class="msg ok">{E(q["ok"])}</div>' if "ok" in q else ""
        tiles = "".join(f'<a class="tile" href="/erfassen/{k}"><b>{E(t)}</b><span>{E(s)}</span></a>'
                        for k, (t, s, _, _) in ERFASSUNG.items())
        today = dt.date.today()
        rows = rows_between(today, today)
        lst = ""
        for r in reversed(rows):
            vals = [f'{E(FD[k]["label"])}: <span class="{status(k, r[k])}">{E(fmt_num(r[k], FD[k]["dec"]) if FD[k]["typ"] == "num" else r[k])}</span>'
                    for k in KEYS if r[k] not in (None, "")]
            lst += f'<tr><td>{E(r["zeit"])}</td><td>{E(r["ort"])}</td><td>{E(r["visum"])}</td><td class="l" style="text-align:left">{" · ".join(vals)}</td></tr>'
        lst = lst or '<tr><td colspan="4">Heute noch keine Messungen.</td></tr>'
        body = (f'{msg}<h1>Was möchtest du erfassen?</h1><div class="tiles">{tiles}</div>'
                f'<h2>Heute, {today.strftime("%d.%m.%Y")}</h2><div class="wrap"><table class="list" style="width:100%">'
                f'<tr><th>Zeit</th><th>Ort</th><th>Visum</th><th>Werte</th></tr>{lst}</table></div>'
                f'<p>{limits_text()}</p>')
        self.send(page("Erfassen", body, self.is_admin(), "start", refresh=60))

    def erfassen(self, art, method):
        if art not in ERFASSUNG:
            return self.redirect("/")
        title, sub, keys, ort = ERFASSUNG[art]
        vals, err = {}, ""
        if method == "POST":
            vals = self.form()
            n = now()
            vals.update(datum=n.date().isoformat(), zeit=n.strftime("%H:%M"))
            if ort:
                vals["ort"] = ort
            data, errs = validate(vals, keys, need_dt=True)
            if not errs:
                save_row(data, self.client())
                warn = [FD[k]["label"] for k in keys if status(k, data.get(k)) in ("warn", "bad")]
                txt = f"Gespeichert: {title} um {data['zeit']} ({data['visum']})."
                if warn:
                    txt += " Achtung, ausserhalb Referenz/Toleranz: " + ", ".join(warn)
                return self.redirect("/?" + urlencode({"ok": txt}))
            err = '<div class="msg err">' + "<br>".join(E(e) for e in errs) + "</div>"
        fields = "".join(field_html(k, vals.get(k, "")) for k in keys)
        top = visum_html(vals.get("visum", "")) + ("" if ort else ort_html(vals.get("ort", "Büro")))
        body = (f'<h1>{E(title)}</h1>{err}<form class="f" method="post">'
                f'<p class="msg inf">Datum und Zeit werden automatisch gesetzt. Nur gemessene Werte ausfüllen.</p>'
                f'{top}{fields}<button class="btn" style="width:100%;font-size:1.3rem;padding:14px">Speichern</button>'
                f'<p><a href="/">Abbrechen</a></p></form>{JS_LIMITS % json.dumps(CFG["grenzwerte"])}')
        self.send(page(title, body, self.is_admin(), "start"))

    def woche(self):
        q = self.q()
        d = parse_date(q.get("datum")) or dt.date.today()
        mo = monday_of(d)
        prev, nxt = mo - dt.timedelta(days=7), mo + dt.timedelta(days=7)
        nav = (f'<div class="bar noprint"><a class="btn sec" href="/woche?datum={prev}">◀ Vorwoche</a>'
               f'<a class="btn sec" href="/woche">Diese Woche</a><a class="btn sec" href="/woche?datum={nxt}">Nächste ▶</a>'
               f'<form method="get" style="display:flex;gap:8px;align-items:end"><div class="fld"><label>Datum</label>'
               f'<input type="date" name="datum" value="{mo}"></div><button class="btn sec">Anzeigen</button></form></div>')
        self.send(page("Woche", nav + week_block(mo), self.is_admin(), "woche", refresh=120))

    def login(self, method):
        if not CFG["admin_passwort"]:
            return self.send(page("PC-Bereich", '<h1>PC-Bereich</h1><div class="msg inf">Der PC-Bereich (Bearbeiten, Löschen, '
                                  'Berichte) ist nur direkt am Büro-PC unter <b>http://localhost:%s</b> verfügbar.</div>' % CFG["port"]))
        err = ""
        if method == "POST":
            if secrets.compare_digest(self.form().get("pw", ""), CFG["admin_passwort"]):
                sid = secrets.token_urlsafe(24)
                SESSIONS.add(sid)
                return self.redirect("/admin", {"Set-Cookie": f"sid={sid}; HttpOnly; Path=/; SameSite=Strict"})
            err = '<div class="msg err">Falsches Passwort</div>'
        self.send(page("Anmelden", f'<h1>PC-Bereich</h1>{err}<form class="f" method="post"><div class="fld"><label>Passwort</label>'
                                   f'<input type="password" name="pw" autofocus></div><button class="btn">Anmelden</button></form>'))

    # --- PC-Bereich
    def range_from_q(self, default_days=6):
        q = self.q()
        bis = parse_date(q.get("bis")) or dt.date.today()
        von = parse_date(q.get("von")) or (bis - dt.timedelta(days=default_days))
        return von, bis

    def admin_list(self):
        von, bis = self.range_from_q()
        q = self.q()
        rows = rows_between(von, bis)
        msg = f'<div class="msg ok">{E(q["msg"])}</div>' if "msg" in q else ""
        keys_used = [k for k in KEYS if any(r[k] not in (None, "") for r in rows)] or HAND
        head = "".join(f"<th>{E(FD[k]['csv'])}</th>" for k in keys_used)
        qs = urlencode({"von": von, "bis": bis})
        trs = ""
        for r in rows:
            tds = ""
            for k in keys_used:
                v = r[k]
                txt = fmt_num(v, FD[k]["dec"]) if FD[k]["typ"] == "num" else (v or "")
                cls = status(k, v) + (" l" if FD[k]["typ"] == "text" else "")
                tds += f'<td class="{cls}">{E(txt)}</td>'
            trs += (f'<tr><td class="l">{dt.date.fromisoformat(r["datum"]).strftime("%d.%m.%Y")}</td><td>{E(r["zeit"])}</td>'
                    f'<td class="l">{E(r["ort"])}</td><td>{E(r["visum"])}</td>{tds}'
                    f'<td style="white-space:nowrap"><a class="btn sm sec" href="/admin/edit/{r["id"]}?{qs}" title="Bearbeiten">✎</a> '
                    f'<form method="post" action="/admin/delete/{r["id"]}?{qs}" style="display:inline" '
                    f'onsubmit="return confirm(\'Messung wirklich löschen?\')"><button class="btn sm del" title="Löschen">✕</button></form></td></tr>')
        trs = trs or f'<tr><td colspan="{5 + len(keys_used)}">Keine Messungen im Zeitraum.</td></tr>'
        body = (f'<h1>Messungen</h1>{msg}<form class="bar" method="get">'
                f'<div class="fld"><label>Von</label><input type="date" name="von" value="{von}"></div>'
                f'<div class="fld"><label>Bis</label><input type="date" name="bis" value="{bis}"></div>'
                f'<button class="btn sec">Anzeigen</button><a class="btn" href="/admin/neu?{qs}">+ Neue Messung</a>'
                f'<a class="btn sec" href="/admin/bericht?{qs}">PDF-Bericht</a><a class="btn sec" href="/admin/export.csv?{qs}">CSV</a></form>'
                f'<p>{len(rows)} Messungen · {limits_text()}</p><div class="wrap"><table class="list"><tr><th>Datum</th><th>Zeit</th>'
                f'<th>Ort</th><th>Visum</th>{head}<th></th></tr>{trs}</table></div>')
        self.send(page("Messungen", body, True, "admin"))

    def admin_edit(self, mid, method):
        back = "/admin?" + urlencode(self.q())
        vals, err = {}, ""
        if mid and method == "GET":
            with db() as con:
                r = con.execute("SELECT * FROM messungen WHERE id=?", (mid,)).fetchone()
            if not r:
                return self.redirect(back)
            vals = dict(r)
        if method == "POST":
            vals = self.form()
            data, errs = validate(vals, KEYS, need_dt=True)
            if not errs:
                save_row(data, self.client(), mid)
                return self.redirect(back + "&msg=Gespeichert")
            err = '<div class="msg err">' + "<br>".join(E(e) for e in errs) + "</div>"
        n = now()
        top = (f'<div class="grid2"><div class="fld"><label>Datum</label><input type="date" name="datum" value="{E(vals.get("datum") or n.date())}"></div>'
               f'<div class="fld"><label>Zeit</label><input type="time" name="zeit" value="{E(vals.get("zeit") or n.strftime("%H:%M"))}"></div></div>'
               f'<div class="grid2">{ort_html(vals.get("ort") or "Büro")}{visum_html(vals.get("visum") or "")}</div>')
        groups = [("Handmessung (Swan)", HAND), ("Automat. Messung (Sensoren)", AUTO),
                  ("Tageswerte", ["frischwasser", "quelle", "besucher", "rueckspuelung", "anschwemmung", "ph_stapeltank"]),
                  ("Notizen", ["bemerkungen", "massnahmen", "stoerungen", "kalibrierungen"])]
        fs = "".join(f'<h2>{E(t)}</h2><div class="grid2">{"".join(field_html(k, vals.get(k)) for k in ks)}</div>' for t, ks in groups)
        info = f'<p class="lim">Erfasst: {E(vals.get("erfasst_am"))} · Geändert: {E(vals.get("geaendert_am"))}</p>' if mid else ""
        body = (f'<h1>{"Messung bearbeiten" if mid else "Neue Messung"}</h1>{err}<form class="f" method="post" style="max-width:900px">'
                f'{info}{top}{fs}<button class="btn">Speichern</button> <a class="btn sec" href="{back}">Abbrechen</a></form>'
                f'{JS_LIMITS % json.dumps(CFG["grenzwerte"])}')
        self.send(page("Bearbeiten", body, True, "admin"))

    def bericht(self):
        q = self.q()
        t = dt.date.today().replace(day=1)
        last_end = t - dt.timedelta(days=1)
        von = parse_date(q.get("von")) or last_end.replace(day=1)
        bis = parse_date(q.get("bis")) or last_end
        if bis < von:
            von, bis = bis, von
        mo, weeks = monday_of(von), []
        while mo <= bis and len(weeks) < 60:
            weeks.append(week_block(mo, von, bis))
            mo += dt.timedelta(days=7)
        bar = (f'<form class="bar noprint" method="get"><div class="fld"><label>Von</label><input type="date" name="von" value="{von}"></div>'
               f'<div class="fld"><label>Bis</label><input type="date" name="bis" value="{bis}"></div>'
               f'<button class="btn sec">Anzeigen</button><button type="button" class="btn" onclick="window.print()">Drucken / als PDF speichern</button></form>'
               f'<p class="noprint msg inf">Im Druckdialog «Als PDF speichern» bzw. «Microsoft Print to PDF» wählen. '
               f'Jede Woche kommt auf eine A4-Seite quer. Tage ausserhalb des Zeitraums sind grau.</p>')
        self.send(page("Bericht", bar + "".join(weeks), True, "bericht"))

    def export(self):
        von, bis = self.range_from_q(36500)
        name = f"messungen_{von}_bis_{bis}.csv"
        self.send(to_csv(rows_between(von, bis)), ctype="text/csv; charset=utf-8",
                  headers={"Content-Disposition": f'attachment; filename="{name}"'})

    def daten(self, method):
        msg = ""
        if method == "POST":
            f = self.form()
            raw = f.get("datei") or b""
            if isinstance(raw, str):
                raw = raw.encode()
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError:
                text = raw.decode("cp1252")
            if not text.strip():
                msg = '<div class="msg err">Keine Datei gewählt.</div>'
            else:
                neu, upd, fehler = import_csv(text, self.client())
                msg = f'<div class="msg ok">Import: {neu} neu, {upd} aktualisiert.</div>'
                if fehler:
                    msg += '<div class="msg err">Übersprungen:<br>' + "<br>".join(E(x) for x in fehler[:50]) + "</div>"
        st = BACKUP_STATUS
        files = sorted(os.listdir(BACKUP_DIR), reverse=True)[:15] if os.path.isdir(BACKUP_DIR) else []
        body = (f'<h1>Backup & CSV</h1>{msg}'
                f'<h2>Automatisches Backup</h2><p>Bei jeder Änderung wird gesichert nach:<br><code>{E(BACKUP_DIR)}</code><br>'
                f'Letztes Backup: {E(st["zeit"]) if st["zeit"] else "–"} {"✔" if st["ok"] else "✖ " + E(st["fehler"])}</p>'
                f'<ul>{"".join(f"<li>{E(x)}</li>" for x in files)}</ul>'
                f'<p><code>messungen_aktuell.csv</code> = immer der neuste Stand. Pro Tag zusätzlich eine datierte CSV und eine Kopie der Datenbank '
                f'(werden {CFG["backup_tage"]} Tage aufbewahrt).</p>'
                f'<h2>Alles als CSV herunterladen</h2><p><a class="btn" href="/admin/export.csv">CSV herunterladen</a></p>'
                f'<h2>CSV importieren / wiederherstellen</h2><form class="f" method="post" enctype="multipart/form-data">'
                f'<p>Gleiches Format wie der Export (auch aus der Excel-Vorlage gespeichert). Zeilen mit bekannter ID werden überschrieben, '
                f'alle anderen neu angelegt.</p><div class="fld"><input type="file" name="datei" accept=".csv,text/csv"></div>'
                f'<button class="btn">Importieren</button></form>')
        self.send(page("Backup & CSV", body, True, "daten"))

    def protokoll(self):
        with db() as con:
            rows = con.execute("SELECT * FROM protokoll ORDER BY id DESC LIMIT 300").fetchall()
        trs = "".join(f'<tr><td>{E(r["zeit"])}</td><td>{E(r["aktion"])}</td><td>{E(r["messung_id"])}</td><td>{E(r["client"])}</td>'
                      f'<td class="l" style="text-align:left;font-size:.75rem;max-width:700px">{E(r["daten"])}</td></tr>' for r in rows)
        body = ('<h1>Änderungsprotokoll</h1><p>Jede Erfassung, Änderung und Löschung wird mit Zeit und Gerät (IP) festgehalten.</p>'
                f'<div class="wrap"><table class="list"><tr><th>Zeit</th><th>Aktion</th><th>ID</th><th>Gerät</th><th>Daten</th></tr>{trs}</table></div>')
        self.send(page("Protokoll", body, True, "prot"))


def lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except OSError:
        return "IP-des-PCs"


def main():
    init_db()
    backup()
    port = int(CFG["port"])
    srv = ThreadingHTTPServer(("0.0.0.0", port), H)
    print(f"Badi-Journal läuft.\n  Am PC:      http://localhost:{port}\n  Am Tablet:  http://{lan_ip()}:{port}\n"
          f"  Daten:      {DB_PATH}\n  Backups:    {BACKUP_DIR}\nFenster offen lassen. Beenden mit Ctrl+C.")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
