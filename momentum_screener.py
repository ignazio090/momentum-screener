#!/usr/bin/env python3
"""
Momentum Screener — zoekt aandelen met een CONSISTENTE opwaartse trend:
gestegen over dag EN week EN maand EN (halfjaar) EN jaar tegelijk. De filosofie:
zo'n breed gedragen uptrend zet zich op de korte termijn vaker door.

Haalt echte koersen + nieuws op (Finnhub), werkt de Excel bij en stuurt een
Telegram-bericht.

GEEN BELEGGINGSADVIES. De uitkomst is een lijst kandidaten om zelf te
onderzoeken, geen koopopdracht. Brede momentum-trends keren ook abrupt om.

GEBRUIK:
  1. Vul de instellingen in (bij voorkeur via omgevingsvariabelen).
  2. Zet je tickers in tickers.csv (kolom 'ticker'), of laat Finnhub ze ophalen.
  3. Draai:  python momentum_screener.py
     Test zonder internet:  python momentum_screener.py --demo

VEILIGHEID: zet API-keys NOOIT hard in dit bestand. Gebruik omgevingsvariabelen.
"""

import os
import sys
import csv
import time
from datetime import datetime, timedelta

import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side

# ============================================================
# INSTELLINGEN
# ============================================================
FINNHUB_KEY      = os.environ.get("FINNHUB_KEY", "ZET_JE_FINNHUB_KEY_HIER")
TELEGRAM_TOKEN   = os.environ.get("TELEGRAM_TOKEN", "ZET_JE_BOT_TOKEN_HIER")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "ZET_JE_CHAT_ID_HIER")

EXCEL_PATH  = "beleggen_tracker.xlsx"
TICKERS_CSV = "tickers.csv"

# --- HET FILTER (jouw strategie) ---
# Een aandeel is kandidaat als het over ALLE onderstaande periodes stijgt.
PERIODS = {"1D": 1, "1W": 7, "1M": 30, "6M": 182, "1J": 365}
REQUIRE_ALL_RISING = True   # True = moet over ELKE periode stijgen (jouw kernidee)
MIN_PER_PERIOD     = 0.00   # minimum per periode. 0.00 = "gewoon positief".
                            # Zet bv op 0.02 (=2%) of 0.05 (=5%) om strenger te zijn.
                            # Let op: 5% op de DAG is zeldzaam; begin laag en stel bij.

MAX_HITS_IN_MESSAGE = 10
NEWS_PER_STOCK      = 2
SLEEP_BETWEEN_CALLS = 1.1   # Finnhub gratis ≈ 60 calls/min

# Kosten (alleen voor de netto-indicatie in het bericht; pas aan op je broker).
COST_FIXED  = 1.00
COST_VAR    = 0.0000
COST_FX     = 0.0000
TARGET_GAIN = 0.05   # alleen om te tonen wat 5% koerswinst na kosten oplevert

UNIVERSE_FROM_FINNHUB = False
FINNHUB_EXCHANGE = "US"

BASE = "https://finnhub.io/api/v1"


# ============================================================
# DATA
# ============================================================
def get_json(url, params):
    r = requests.get(url, params=dict(params, token=FINNHUB_KEY), timeout=20)
    r.raise_for_status()
    return r.json()


def load_universe():
    if UNIVERSE_FROM_FINNHUB:
        data = get_json(f"{BASE}/stock/symbol", {"exchange": FINNHUB_EXCHANGE})
        return [d["symbol"] for d in data if d.get("type") == "Common Stock"]
    if not os.path.exists(TICKERS_CSV):
        sys.exit(f"Geen {TICKERS_CSV} gevonden. Maak een CSV met een kolom 'ticker'.")
    out = []
    with open(TICKERS_CSV, newline="") as f:
        for row in csv.DictReader(f):
            t = (row.get("ticker") or "").strip()
            if t:
                out.append(t)
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
            if ts <= target:
                ref = c
            else:
                break
        out[label] = ((last - ref) / ref) if (ref and ref > 0) else None
    return out


def get_news(ticker):
    to = datetime.utcnow().date()
    frm = to - timedelta(days=7)
    try:
        data = get_json(f"{BASE}/company-news", {"symbol": ticker, "from": str(frm), "to": str(to)})
    except Exception:
        return []
    return [{"headline": a.get("headline", "")[:90], "url": a.get("url", "")}
            for a in data[:NEWS_PER_STOCK]]


# ============================================================
# FILTER + SCORE  (consistente uptrend)
# ============================================================
def period_vals(ch):
    return [v for k, v in ch.items() if k != "price"]


def is_hit(ch):
    vals = period_vals(ch)
    if any(v is None for v in vals):      # onvoldoende historie -> overslaan
        return False
    if REQUIRE_ALL_RISING:
        return all(v >= MIN_PER_PERIOD for v in vals)
    # alternatief (uit): minstens één periode boven de drempel
    return any(v >= MIN_PER_PERIOD for v in vals)


def momentum_score(ch):
    """Gemiddelde stijging over alle periodes = hoe sterk de brede uptrend is."""
    vals = [v for v in period_vals(ch) if v is not None]
    return sum(vals) / len(vals) if vals else -99


# ============================================================
# EXCEL
# ============================================================
def update_excel(hits):
    if not os.path.exists(EXCEL_PATH):
        print(f"Let op: {EXCEL_PATH} niet gevonden; Excel-update overgeslagen.")
        return
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Screener"] if "Screener" in wb.sheetnames else wb.create_sheet("Screener")
    row = 5
    while ws.cell(row=row, column=1).value not in (None, ""):
        row += 1
    today = datetime.now().strftime("%Y-%m-%d")
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for h in hits:
        c = h["changes"]
        vals = [today, h["ticker"], h.get("name", h["ticker"]), c["price"],
                c.get("1D"), c.get("1W"), c.get("1M"), c.get("6M"), c.get("1J"),
                "JA", h["news"][0]["url"] if h["news"] else ""]
        for col, v in enumerate(vals, 1):
            cell = ws.cell(row=row, column=col, value=v)
            cell.border = border
            if col == 4:
                cell.number_format = "€#,##0.00"
            if 5 <= col <= 9 and v is not None:
                cell.number_format = "0.0%;(0.0%)"
                cell.fill = PatternFill("solid", start_color="DDEFE2" if v >= 0 else "F7DDDD")
        row += 1
    wb.save(EXCEL_PATH)
    print(f"Excel bijgewerkt: {len(hits)} kandidaten toegevoegd aan 'Screener'.")


# ============================================================
# TELEGRAM
# ============================================================
def fmt(v):
    return f"{v*100:+.1f}%".replace(".", ",") if v is not None else "—"


def net_after_costs(order=1000.0):
    sell = order * (1 + TARGET_GAIN)
    cost = 2 * COST_FIXED + (COST_VAR + COST_FX) * (order + sell)
    return (order * TARGET_GAIN - cost) / order


def build_message(hits):
    if not hits:
        return "📉 Geen aandelen met een consistente uptrend over alle periodes vandaag."
    lines = [f"📈 *Consistente stijgers — {datetime.now().strftime('%d-%m-%Y')}*",
             f"{len(hits)} aandelen stijgen over ALLE periodes (toon max {MAX_HITS_IN_MESSAGE}):", ""]
    for h in hits[:MAX_HITS_IN_MESSAGE]:
        c = h["changes"]
        trend = "  ".join(f"{k} {fmt(c.get(k))}" for k in PERIODS)
        lines.append(f"• *{h['ticker']}*")
        lines.append(f"   {trend}")
        for n in h["news"]:
            if n["url"]:
                lines.append(f"   ↳ [{n['headline']}]({n['url']})")
    lines.append("")
    lines.append("_Geen advies — zelf onderzoeken voor je iets koopt._")
    return "\n".join(lines)


def send_telegram(text):
    if "ZET_JE" in TELEGRAM_TOKEN or "ZET_JE" in str(TELEGRAM_CHAT_ID):
        print("Telegram niet geconfigureerd — bericht zou zijn:\n")
        print(text)
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    r = requests.post(url, data={"chat_id": TELEGRAM_CHAT_ID, "text": text,
                                 "parse_mode": "Markdown", "disable_web_page_preview": True}, timeout=20)
    print("Telegram-bericht verstuurd." if r.ok else f"Telegram-fout: {r.text}")


# ============================================================
# MAIN
# ============================================================
def run_real():
    if "ZET_JE" in FINNHUB_KEY:
        sys.exit("Stel eerst FINNHUB_KEY in (omgevingsvariabele).")
    universe = load_universe()
    print(f"{len(universe)} tickers te screenen...")
    hits = []
    for i, tk in enumerate(universe, 1):
        try:
            ch = pct_changes(tk)
        except Exception as e:
            print(f"  {tk}: fout ({e})"); time.sleep(SLEEP_BETWEEN_CALLS); continue
        if ch and is_hit(ch):
            hits.append({"ticker": tk, "name": tk, "changes": ch, "news": get_news(tk)})
            print(f"  ✓ {tk} — consistente uptrend")
        time.sleep(SLEEP_BETWEEN_CALLS)
        if i % 25 == 0:
            print(f"  ...{i}/{len(universe)}")
    hits.sort(key=lambda h: momentum_score(h["changes"]), reverse=True)
    update_excel(hits)
    send_telegram(build_message(hits))


def run_demo():
    print("DEMO-modus: fictieve data, geen API-calls.\n")
    demo = [
        {"ticker": "NVDA", "name": "NVIDIA",
         "news": [{"headline": "NVIDIA kondigt nieuwe chip aan", "url": "https://example.com/nvda"}],
         "changes": {"price": 118.5, "1D": 0.021, "1W": 0.034, "1M": 0.071, "6M": 0.18, "1J": 0.42}},
        {"ticker": "ADYEN", "name": "Adyen",
         "news": [{"headline": "Adyen breidt uit in Azië", "url": "https://example.com/adyen"}],
         "changes": {"price": 1450.0, "1D": 0.012, "1W": 0.028, "1M": 0.04, "6M": 0.10, "1J": 0.22}},
        # SHELL: daalt op 1 dag -> valt buiten het EN-filter (laat de logica zien)
        {"ticker": "SHELL", "name": "Shell plc", "news": [],
         "changes": {"price": 32.05, "1D": -0.004, "1W": 0.011, "1M": 0.03, "6M": 0.06, "1J": 0.14}},
    ]
    hits = [h for h in demo if is_hit(h["changes"])]
    hits.sort(key=lambda h: momentum_score(h["changes"]), reverse=True)
    print(f"{len(hits)} van {len(demo)} aandelen voldoen aan het EN-filter "
          f"(SHELL valt af: daalt op 1D).\n")
    update_excel(hits)
    print("--- Telegram-bericht (preview) ---\n")
    print(build_message(hits))


if __name__ == "__main__":
    run_demo() if "--demo" in sys.argv else run_real()
