name: Taiwan Stock Report

on:
  schedule:
    - cron: "0 6 * * 1-5"
  workflow_dispatch:

jobs:
  send-report:
    runs-on: ubuntu-latest

    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Setup Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"
          cache: "pip"

      - name: Install dependencies
        run: pip install yfinance openpyxl pandas pytz matplotlib numpy Pillow resend

      - name: Run report
        env:
          RESEND_API_KEY: ${{ secrets.RESEND_API_KEY }}
          EMAIL_FROM: ${{ secrets.EMAIL_FROM }}
          EMAIL_TO: ${{ secrets.EMAIL_TO }}
        run: python stock_report.py
