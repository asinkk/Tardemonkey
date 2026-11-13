"""抓加權指數 60 分K（Yahoo，日盤 09:00–13:30，最多約 3 年），合併存到 history/twii_60min.csv。
沒有夜盤，用來回測「日盤版」策略；檔案會跟舊資料合併，所以可以越存越長。
用法：python strategy/fetch_twii.py [輸出 csv]
"""
import csv, datetime as dt, json, os, sys, urllib.request

URL = "https://query1.finance.yahoo.com/v8/finance/chart/%5ETWII?interval=60m&range=730d"
out = sys.argv[1] if len(sys.argv) > 1 else os.path.join("history", "twii_60min.csv")

req = urllib.request.Request(URL, headers={"User-Agent": "Mozilla/5.0"})
d = json.load(urllib.request.urlopen(req, timeout=60))["chart"]["result"][0]
q = d["indicators"]["quote"][0]
rows = {}
if os.path.exists(out):
    rows = {r["start"]: r for r in csv.DictReader(open(out, encoding="utf-8"))}
today = dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).strftime("%Y-%m-%d")
for i, t in enumerate(d["timestamp"]):
    o, h, l, c = (q[k][i] for k in ("open", "high", "low", "close"))
    if None in (o, h, l, c):
        continue
    start = (dt.datetime.fromtimestamp(t, dt.timezone.utc) + dt.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M")
    if start.endswith("13:30") or start.startswith(today):  # 收盤競價點與當天未收完的 K 不存
        continue
    rows[start] = dict(start=start, open=f"{o:.2f}", high=f"{h:.2f}", low=f"{l:.2f}", close=f"{c:.2f}",
                       volume=q["volume"][i] or 0)
with open(out, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=["start", "open", "high", "low", "close", "volume"])
    w.writeheader()
    for k in sorted(rows):
        w.writerow(rows[k])
print("輸出", out, len(rows), "根", min(rows), "~", max(rows))
