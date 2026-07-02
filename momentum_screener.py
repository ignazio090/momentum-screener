#!/usr/bin/env python3
"""
Momentum Screener — zoekt aandelen met een CONSISTENTE opwaartse trend:
gestegen over dag EN week EN maand EN halfjaar EN jaar tegelijk.

Haalt echte koersen + nieuws op (Finnhub), werkt de Excel bij, stuurt een
Telegram-bericht, en genereert een dashboard-website.

GEEN BELEGGINGSADVIES. Brede momentum-trends keren ook abrupt om.
"""

import os, sys, csv, time, json
from datetime import datetime, timedelta
import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side

# ============================================================
# INSTELLINGEN
# ============================================================
FINNHUB_KEY      = os.environ.get("FINNHUB_KEY", "")
TELEGRAM_TOKEN   = os.environ.get("TELEGRAM_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

EXCEL_PATH   = "beleggen_tracker.xlsx"
TICKERS_CSV  = "tickers.csv"
HISTORY_JSON = "results_history.json"
DASHBOARD    = "docs/index.html"

PERIODS = {"1D": 1, "1W": 7, "1M": 30, "6M": 182, "1J": 365}
REQUIRE_ALL_RISING = True
MIN_PER_PERIOD     = 0.00

MAX_HITS_IN_MESSAGE = 10
NEWS_PER_STOCK      = 3
SLEEP_BETWEEN_CALLS = 1.05

BASE = "https://finnhub.io/api/v1"


# ============================================================
# DATA
# ============================================================
def get_json(url, params):
    r = requests.get(url, params=dict(params, token=FINNHUB_KEY), timeout=20)
    r.raise_for_status()
    return r.json()

def load_universe():
    if not os.path.exists(TICKERS_CSV):
        sys.exit(f"Geen {TICKERS_CSV} gevonden.")
    out = []
    with open(TICKERS_CSV, newline="") as f:
        for row in csv.DictReader(f):
            t = (row.get("ticker") or "").strip()
            if t: out.append(t)
    return out

def pct_changes(ticker):
    now = int(time.time())
    frm = now - 400 * 86400
    candles = get_json(f"{BASE}/stock/candle",
                       {"symbol": ticker, "resolution": "D", "from": frm, "to": now})
    if candles.get("s") != "ok" or not candles.get("c"):
        return None
    closes, stamps = candles["c"], candles["t"]
    last = closes[-1]
    out = {"price": last}
    for label, days in PERIODS.items():
        target = now - days * 86400
        ref = None
        for ts, c in zip(stamps, closes):
            if ts <= target: ref = c
            else: break
        out[label] = ((last - ref) / ref) if (ref and ref > 0) else None
    return out

def get_news(ticker):
    to = datetime.utcnow().date()
    frm = to - timedelta(days=7)
    try:
        data = get_json(f"{BASE}/company-news", {"symbol": ticker, "from": str(frm), "to": str(to)})
    except: return []
    return [{"headline": a.get("headline","")[:100], "url": a.get("url","")}
            for a in data[:NEWS_PER_STOCK]]


# ============================================================
# FILTER + SCORE
# ============================================================
def is_hit(ch):
    vals = [v for k,v in ch.items() if k != "price"]
    if any(v is None for v in vals): return False
    return all(v >= MIN_PER_PERIOD for v in vals) if REQUIRE_ALL_RISING else any(v >= MIN_PER_PERIOD for v in vals)

def momentum_score(ch):
    vals = [v for k,v in ch.items() if k != "price" and v is not None]
    return sum(vals)/len(vals) if vals else -99


# ============================================================
# EXCEL
# ============================================================
def update_excel(hits):
    if not os.path.exists(EXCEL_PATH): return
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Screener"] if "Screener" in wb.sheetnames else wb.create_sheet("Screener")
    row = 5
    while ws.cell(row=row, column=1).value not in (None, ""): row += 1
    today = datetime.now().strftime("%Y-%m-%d")
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for h in hits:
        c = h["changes"]
        vals = [today, h["ticker"], h["ticker"], c["price"],
                c.get("1D"), c.get("1W"), c.get("1M"), c.get("6M"), c.get("1J"),
                "JA", h["news"][0]["url"] if h["news"] else ""]
        for col, v in enumerate(vals, 1):
            cell = ws.cell(row=row, column=col, value=v)
            cell.border = border
            if col == 4: cell.number_format = "€#,##0.00"
            if 5 <= col <= 9 and v is not None:
                cell.number_format = "0.0%;(0.0%)"
                cell.fill = PatternFill("solid", start_color="DDEFE2" if v >= 0 else "F7DDDD")
        row += 1
    wb.save(EXCEL_PATH)
    print(f"Excel: {len(hits)} kandidaten toegevoegd.")


# ============================================================
# HISTORY
# ============================================================
def load_history():
    if os.path.exists(HISTORY_JSON):
        with open(HISTORY_JSON) as f: return json.load(f)
    return {}

def save_history(history, hits):
    today = datetime.now().strftime("%Y-%m-%d")
    history[today] = [{
        "ticker": h["ticker"], "price": h["changes"]["price"],
        "1D": h["changes"].get("1D"), "1W": h["changes"].get("1W"),
        "1M": h["changes"].get("1M"), "6M": h["changes"].get("6M"),
        "1J": h["changes"].get("1J"),
        "news": h["news"][:2], "score": round(momentum_score(h["changes"])*100, 1)
    } for h in hits]
    # bewaar max 90 dagen
    cutoff = (datetime.now() - timedelta(days=90)).strftime("%Y-%m-%d")
    history = {k:v for k,v in history.items() if k >= cutoff}
    with open(HISTORY_JSON, "w") as f: json.dump(history, f, indent=2)
    return history


# ============================================================
# DASHBOARD GENEREREN
# ============================================================
def generate_dashboard(history):
    os.makedirs("docs", exist_ok=True)
    dates = sorted(history.keys(), reverse=True)
    today = dates[0] if dates else "—"
    today_hits = history.get(today, [])

    def fmt(v):
        if v is None: return "—"
        s = f"{v*100:+.1f}%".replace(".",",")
        return s

    def cls(v):
        if v is None: return "muted"
        return "up" if v >= 0 else "down"

    # build rows for today
    rows_html = ""
    for s in today_hits:
        news_links = ""
        for n in s.get("news", []):
            if n.get("url"):
                title = n["headline"][:60]
                news_links += f'<a class="news" href="{n["url"]}" target="_blank">{title}… ↗</a><br>'
        if not news_links: news_links = '<span class="muted">geen</span>'
        rows_html += f"""<tr>
<td><span class="tk">{s['ticker']}</span></td>
<td>${s['price']:.2f}</td>
<td class="{cls(s.get('1D'))}">{fmt(s.get('1D'))}</td>
<td class="{cls(s.get('1W'))}">{fmt(s.get('1W'))}</td>
<td class="{cls(s.get('1M'))}">{fmt(s.get('1M'))}</td>
<td class="{cls(s.get('6M'))}">{fmt(s.get('6M'))}</td>
<td class="{cls(s.get('1J'))}">{fmt(s.get('1J'))}</td>
<td class="score">{s.get('score','—')}</td>
<td>{news_links}</td>
</tr>"""
    if not rows_html:
        rows_html = '<tr><td colspan="9" class="empty">Geen kandidaten vandaag — geen enkel aandeel stijgt over álle periodes tegelijk.</td></tr>'

    # history section
    hist_html = ""
    for d in dates[:30]:
        n = len(history[d])
        tickers = ", ".join(h["ticker"] for h in history[d][:8])
        if len(history[d]) > 8: tickers += f" +{len(history[d])-8}"
        hist_html += f'<div class="hist-row"><span class="hist-date">{d}</span><span class="hist-count">{n}</span><span class="hist-tickers">{tickers}</span></div>'
    if not hist_html:
        hist_html = '<div class="hist-row"><span class="muted">Nog geen historie — morgen verschijnt hier de eerste dag.</span></div>'

    html = f"""<!DOCTYPE html>
<html lang="nl"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Momentum Screener Dashboard</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;700&display=swap');
:root{{--bg:#0B0E0C;--panel:#141A16;--panel2:#1B231D;--line:#26302A;--text:#E8EFE9;--muted:#8A958C;--up:#4ADE80;--down:#F87171;--accent:#4ADE80;}}
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{background:var(--bg);color:var(--text);font-family:'Inter',sans-serif;padding:36px 20px 80px;}}
.wrap{{max-width:1120px;margin:0 auto;}}
.eyebrow{{font-family:'JetBrains Mono',monospace;font-size:11px;letter-spacing:.18em;text-transform:uppercase;color:var(--accent);display:flex;align-items:center;gap:8px;}}
.pulse{{width:7px;height:7px;border-radius:50%;background:var(--accent);animation:pulse 2s infinite;}}
@keyframes pulse{{0%{{box-shadow:0 0 0 0 rgba(74,222,128,.5)}}70%{{box-shadow:0 0 0 8px rgba(74,222,128,0)}}100%{{box-shadow:0 0 0 0 rgba(74,222,128,0)}}}}
h1{{font-family:'Space Grotesk',sans-serif;font-weight:700;font-size:clamp(28px,4vw,42px);letter-spacing:-.02em;margin:10px 0;}}
.sub{{color:var(--muted);font-size:14px;max-width:600px;line-height:1.55;margin-bottom:6px;}}
.badge{{display:inline-flex;align-items:center;gap:6px;padding:5px 11px;border:1px solid var(--line);border-radius:8px;font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--muted);margin-bottom:28px;}}
.panel{{background:var(--panel);border:1px solid var(--line);border-radius:14px;overflow:hidden;margin-bottom:24px;}}
.panel-head{{padding:16px 20px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;}}
.panel-head h2{{font-family:'Space Grotesk';font-size:16px;font-weight:600;}}
.panel-head .meta{{font-family:'JetBrains Mono';font-size:10.5px;color:var(--muted);}}
table{{width:100%;border-collapse:collapse;}}
th{{font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);font-weight:500;text-align:right;padding:12px 12px;border-bottom:1px solid var(--line);}}
th:first-child,td:first-child{{text-align:left;padding-left:20px;}}
th:last-child,td:last-child{{padding-right:20px;}}
td{{padding:12px 12px;text-align:right;font-family:'JetBrains Mono',monospace;font-size:12.5px;border-bottom:1px solid var(--line);vertical-align:top;}}
tr:last-child td{{border-bottom:none;}}
tbody tr:hover{{background:var(--panel2);}}
.tk{{font-family:'Space Grotesk';font-weight:700;font-size:14px;}}
.up{{color:var(--up);}}
.down{{color:var(--down);}}
.muted{{color:var(--muted);}}
.score{{color:var(--accent);font-weight:700;}}
.news{{color:var(--accent);text-decoration:none;font-family:'Inter';font-size:11px;display:inline-block;margin:2px 0;}}
.news:hover{{text-decoration:underline;}}
.empty{{text-align:center;color:var(--muted);font-family:'Inter';font-size:13px;padding:30px;}}
.hist-row{{display:flex;gap:14px;align-items:baseline;padding:9px 20px;border-bottom:1px solid var(--line);font-size:13px;}}
.hist-row:last-child{{border-bottom:none;}}
.hist-date{{font-family:'JetBrains Mono';font-size:12px;color:var(--muted);min-width:90px;}}
.hist-count{{font-family:'Space Grotesk';font-weight:700;color:var(--accent);min-width:30px;}}
.hist-tickers{{color:var(--text);font-size:12.5px;}}
footer{{margin-top:36px;padding-top:16px;border-top:1px solid var(--line);font-size:10.5px;color:var(--muted);font-family:'JetBrains Mono',monospace;line-height:1.6;}}
@media(max-width:700px){{th,td{{padding:8px 6px;font-size:11px;}} .tk{{font-size:12px;}}}}
</style></head><body>
<div class="wrap">
<div class="eyebrow"><span class="pulse"></span> MOMENTUM SCREENER</div>
<h1>Consistente stijgers</h1>
<p class="sub">Aandelen die over dag, week, maand, halfjaar én jaar tegelijk stijgen. Automatisch bijgewerkt elke werkdag om 08:00.</p>
<div class="badge">Laatst bijgewerkt: {today} · {len(today_hits)} kandidaten · 319 aandelen gescand</div>

<div class="panel">
<div class="panel-head"><h2>Kandidaten — {today}</h2><span class="meta">FILTER: POSITIEF OVER ALLE PERIODES</span></div>
<table><thead><tr>
<th>Ticker</th><th>Koers</th><th>1D</th><th>1W</th><th>1M</th><th>6M</th><th>1J</th><th>Score</th><th>Nieuws</th>
</tr></thead><tbody>{rows_html}</tbody></table>
</div>

<div class="panel">
<div class="panel-head"><h2>Historie (laatste 30 dagen)</h2><span class="meta">AANTAL KANDIDATEN PER DAG</span></div>
{hist_html}
</div>

<footer>MOMENTUM SCREENER — AUTOMATISCH GEGENEREERD · DIT IS GEEN BELEGGINGSADVIES<br>
De screener filtert op brede uptrend; dat is geen garantie dat de trend doorzet.</footer>
</div></body></html>"""

    with open(DASHBOARD, "w") as f:
        f.write(html)
    print(f"Dashboard gegenereerd: {DASHBOARD}")


# ============================================================
# TELEGRAM
# ============================================================
def fmt_pct(v):
    return f"{v*100:+.1f}%".replace(".",",") if v is not None else "—"

def build_message(hits):
    if not hits:
        return "📉 Geen aandelen met een consistente uptrend over alle periodes vandaag."
    lines = [f"📈 *Consistente stijgers — {datetime.now().strftime('%d-%m-%Y')}*",
             f"{len(hits)} aandelen stijgen over ALLE periodes:", ""]
    for h in hits[:MAX_HITS_IN_MESSAGE]:
        c = h["changes"]
        trend = "  ".join(f"{k} {fmt_pct(c.get(k))}" for k in PERIODS)
        lines.append(f"• *{h['ticker']}*")
        lines.append(f"   {trend}")
        for n in h["news"][:1]:
            if n["url"]: lines.append(f"   ↳ [{n['headline'][:50]}]({n['url']})")
    if len(hits) > MAX_HITS_IN_MESSAGE:
        lines.append(f"\n+{len(hits)-MAX_HITS_IN_MESSAGE} meer — zie dashboard.")
    lines.append("\n_Geen advies — zelf onderzoeken._")
    return "\n".join(lines)

def send_telegram(text):
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        print("Telegram niet geconfigureerd.\n"); print(text); return
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    r = requests.post(url, data={"chat_id": TELEGRAM_CHAT_ID, "text": text,
        "parse_mode": "Markdown", "disable_web_page_preview": True}, timeout=20)
    print("Telegram verstuurd." if r.ok else f"Telegram-fout: {r.text}")


# ============================================================
# MAIN
# ============================================================
def run_real():
    if not FINNHUB_KEY:
        sys.exit("Stel FINNHUB_KEY in als omgevingsvariabele.")
    universe = load_universe()
    print(f"Screenen van {len(universe)} aandelen...")
    hits = []
    errors = 0
    for i, tk in enumerate(universe, 1):
        try:
            ch = pct_changes(tk)
        except Exception as e:
            errors += 1
            time.sleep(SLEEP_BETWEEN_CALLS); continue
        if ch and is_hit(ch):
            news = get_news(tk)
            hits.append({"ticker": tk, "changes": ch, "news": news})
            print(f"  ✓ {tk}")
        time.sleep(SLEEP_BETWEEN_CALLS)
        if i % 50 == 0:
            print(f"  ...{i}/{len(universe)} ({len(hits)} hits, {errors} fouten)")
    hits.sort(key=lambda h: momentum_score(h["changes"]), reverse=True)
    print(f"\nKlaar: {len(hits)} kandidaten uit {len(universe)} aandelen ({errors} fouten).\n")

    update_excel(hits)
    history = load_history()
    history = save_history(history, hits)
    generate_dashboard(history)
    send_telegram(build_message(hits))


def run_demo():
    print("DEMO-modus.\n")
    demo = [
        {"ticker":"NVDA","changes":{"price":118.5,"1D":0.021,"1W":0.034,"1M":0.071,"6M":0.18,"1J":0.42},
         "news":[{"headline":"NVIDIA kondigt nieuwe chip aan","url":"https://example.com/nvda"}]},
        {"ticker":"ADYEN","changes":{"price":1450,"1D":0.012,"1W":0.028,"1M":0.04,"6M":0.10,"1J":0.22},
         "news":[{"headline":"Adyen breidt uit","url":"https://example.com/adyen"}]},
        {"ticker":"SHELL","changes":{"price":32.05,"1D":-0.004,"1W":0.011,"1M":0.03,"6M":0.06,"1J":0.14},"news":[]},
    ]
    hits = [h for h in demo if is_hit(h["changes"])]
    hits.sort(key=lambda h: momentum_score(h["changes"]), reverse=True)
    print(f"{len(hits)} van {len(demo)} voldoen aan EN-filter.\n")
    update_excel(hits)
    history = load_history()
    history = save_history(history, hits)
    generate_dashboard(history)
    print("\n--- Telegram-bericht ---\n")
    print(build_message(hits))

if __name__ == "__main__":
    run_demo() if "--demo" in sys.argv else run_real()
