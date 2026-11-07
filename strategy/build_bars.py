"""把 history/tx_1min/ 的 1 分K 合成台指期近月連續 60 分K（日盤+夜盤）。

- 每個檔案用成交量最大的契約當近月；換月時把舊資料加上新舊月價差（回溯調整）。
- 日盤 K 棒從 08:45 起算（5 根），夜盤從 15:00 起算（14 根）。
用法：python strategy/build_bars.py [輸入資料夾] [輸出 csv] [分鐘數，預設 60]
"""
import csv, collections, datetime as dt, glob, os, sys

src = sys.argv[1] if len(sys.argv) > 1 else os.path.join("history", "tx_1min")
dst = sys.argv[2] if len(sys.argv) > 2 else "tx_60min.csv"
minutes = int(sys.argv[3]) if len(sys.argv) > 3 else 60

files = sorted(glob.glob(os.path.join(src, "Daily_*.csv")))
data = []
for f in files:
    rows = list(csv.DictReader(open(f, encoding="utf-8")))
    vol = collections.Counter()
    last = {}
    for r in rows:
        vol[r["contract"]] += int(r["volume"])
        last[r["contract"]] = float(r["close"])  # 依契約、時間排序，所以最後一筆就是收盤
    front = max(vol, key=vol.get)
    data.append((front, last, [r for r in rows if r["contract"] == front]))

offset = [0.0] * len(data)
for i in range(len(data) - 1, 0, -1):
    offset[i - 1] = offset[i]
    new, old = data[i][0], data[i - 1][0]
    if new != old:
        a, b = data[i - 1][1].get(new), data[i - 1][1].get(old)
        if a is not None and b is not None:
            offset[i - 1] += a - b
            print(f"換月 {old}->{new}，價差 {a - b:g}")

bars = {}
for (front, _, rows), off in zip(data, offset):
    for r in rows:
        d = dt.datetime.strptime(r["date"] + r["time"], "%Y%m%d%H%M")
        t = d.time()
        if dt.time(8, 45) <= t <= dt.time(13, 45):
            start, last_k = dt.datetime.combine(d.date(), dt.time(8, 45)), 300 // minutes - 1
        else:
            base = d.date() if t >= dt.time(15, 0) else d.date() - dt.timedelta(days=1)
            start, last_k = dt.datetime.combine(base, dt.time(15, 0)), 840 // minutes - 1
        k = min(int((d - start).total_seconds() // (minutes * 60)), last_k)
        key = start + dt.timedelta(minutes=k * minutes)
        o, h, l, c = (float(r[x]) + off for x in ("open", "high", "low", "close"))
        v = int(r["volume"])
        b = bars.get(key)
        if b is None:
            bars[key] = [o, h, l, c, v]
        else:
            b[1] = max(b[1], h); b[2] = min(b[2], l); b[3] = c; b[4] += v

with open(dst, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["start", "open", "high", "low", "close", "volume"])
    for k in sorted(bars):
        w.writerow([k.strftime("%Y-%m-%d %H:%M")] + [f"{x:g}" for x in bars[k][:4]] + [bars[k][4]])
print("輸出", dst, len(bars), "根", minutes, "分K")
