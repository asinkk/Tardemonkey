"""下載期交所台指期（TX）每日行情（日盤「一般」與夜盤「盤後」），存成 history/tx_daily.csv，給 D2 日K 策略算指標用。

來源：https://www.taifex.com.tw/cht/3/futDataDown（每次最多查 1 個月，Big5 CSV）。
每列是一個交易日、一個契約月份、一個時段的開高低收與成交量（不含價差單）。
「盤後」那列的日期是它所屬的交易日：例如 2026-11-06 盤後 = 11/5 15:00 到 11/6 05:00 的夜盤。換月調整在 live.py 做。
第一次會從 START 抓到今天；之後只重抓最後一筆日期那個月到今天。只用 Python 標準函式庫。

用法：python strategy/fetch_tx_daily.py [輸出 csv]
"""
import csv, datetime as dt, io, os, sys, time, urllib.parse, urllib.request

URL = "https://www.taifex.com.tw/cht/3/futDataDown"
START = dt.date(2024, 1, 1)
UA = {"User-Agent": "Mozilla/5.0 (txo-dashboard)"}
FIELDS = ["date", "session", "contract", "open", "high", "low", "close", "volume"]


def fetch_month(y, m):
    first = dt.date(y, m, 1)
    last = (first + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)
    body = urllib.parse.urlencode({"down_type": "1", "commodity_id": "TX",
                                   "queryStartDate": f"{first:%Y/%m/%d}", "queryEndDate": f"{last:%Y/%m/%d}"}).encode()
    req = urllib.request.Request(URL, body, UA)
    with urllib.request.urlopen(req, timeout=120) as r:
        text = r.read().decode("big5", errors="replace")
    if text.lstrip().startswith("<"):
        return []          # 這個月還沒有資料（或查詢失敗）時期交所回 HTML
    out = []
    for row in csv.reader(io.StringIO(text)):
        if len(row) < 18 or row[1].strip() != "TX" or row[17].strip() not in ("一般", "盤後"):
            continue
        contract = row[2].strip()
        if "/" in contract or not contract.isdigit():
            continue
        try:
            o, h, l, c = (float(row[k]) for k in (3, 4, 5, 6))
            v = int(row[9])
        except ValueError:
            continue
        out.append({"date": row[0].strip().replace("/", "-"), "session": "日盤" if row[17].strip() == "一般" else "夜盤", "contract": contract,
                    "open": f"{o:g}", "high": f"{h:g}", "low": f"{l:g}", "close": f"{c:g}", "volume": v})
    return out


def main():
    dst = sys.argv[1] if len(sys.argv) > 1 else os.path.join("history", "tx_daily.csv")
    rows = {}
    if os.path.exists(dst):
        for r in csv.DictReader(open(dst, encoding="utf-8")):
            rows[(r["date"], r["session"], r["contract"])] = r
    start = START
    if rows:
        last = max(k[0] for k in rows)
        start = (dt.date.fromisoformat(last) - dt.timedelta(days=10)).replace(day=1)   # 往前多抓，月底的夜盤歸到下個月
    today = dt.date.today()
    y, m, got = start.year, start.month, 0
    while (y, m) <= (today.year, today.month):
        try:
            month = fetch_month(y, m)
        except Exception as e:
            print("略過", y, m, e)
            month = []
        for r in month:
            rows[(r["date"], r["session"], r["contract"])] = r
        got += len(month)
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
        time.sleep(1)
    if not rows:
        sys.exit("沒有抓到任何資料")
    tmp = dst + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, FIELDS)
        w.writeheader()
        for k in sorted(rows):
            w.writerow({x: rows[k][x] for x in FIELDS})
    os.replace(tmp, dst)
    print("輸出", dst, len(rows), "列，本次更新", got, "列，最後一天", max(k[0] for k in rows))


if __name__ == "__main__":
    main()
