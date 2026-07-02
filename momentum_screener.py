name: Momentum Screener

on:
  schedule:
    - cron: '0 6 * * 1-5'
  workflow_dispatch:

permissions:
  contents: write

jobs:
  screener:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'

      - name: Installeer packages
        run: pip install requests openpyxl yfinance

      - name: Draai screener
        env:
          FINNHUB_KEY: ${{ secrets.FINNHUB_KEY }}
          TELEGRAM_TOKEN: ${{ secrets.TELEGRAM_TOKEN }}
          TELEGRAM_CHAT_ID: ${{ secrets.TELEGRAM_CHAT_ID }}
        run: python momentum_screener.py

      - name: Sla resultaten op in repo
        run: |
          git config user.name "Momentum Bot"
          git config user.email "bot@users.noreply.github.com"
          git add docs/index.html results_history.json beleggen_tracker.xlsx
          git diff --cached --quiet || git commit -m "Update $(date +%Y-%m-%d)"
          git push

      - name: Upload Excel als artifact
        uses: actions/upload-artifact@v4
        with:
          name: beleggen-tracker
          path: beleggen_tracker.xlsx
