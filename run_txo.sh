#!/bin/sh
# macOS / Linux 排程用：產生 docs/index.html
cd "$(dirname "$0")" && python3 txo_report.py
