"""台指期 60 分K 趨勢＋布林擠壓突破策略回測：原版（草稿）與修訂版。

先用 build_bars.py 產生 60 分K，再執行：
    python strategy/backtest.py tx_60min.csv
每筆以 2 單位計（可分批出場），每單位扣 COST 點交易成本。結果以點數與 R（初始風險倍數）表示。
"""
import csv, datetime as dt, statistics, sys

COST = 3  # 每單位來回成本（點）


def load(path):
    rows = list(csv.DictReader(open(path, encoding="utf-8")))
    T = [dt.datetime.strptime(r["start"], "%Y-%m-%d %H:%M") for r in rows]
    O, H, L, C = ([float(r[k]) for r in rows] for k in ("open", "high", "low", "close"))
    return T, O, H, L, C


def ema(x, p):
    k, out = 2 / (p + 1), [x[0]]
    for v in x[1:]:
        out.append(out[-1] + k * (v - out[-1]))
    return out


def us_dst(d):
    """美國夏令：3 月第二個週日到 11 月第一個週日（用台灣日期近似）。"""
    def nth_sunday(year, month, n):
        first = dt.date(year, month, 1)
        return first + dt.timedelta(days=(6 - first.weekday()) % 7 + 7 * (n - 1))
    return nth_sunday(d.year, 3, 2) <= d.date() < nth_sunday(d.year, 11, 1)


def allowed_revised(t):
    """修訂版時段濾網（看訊號 K 的開始時間）。沒有個別數據日曆，數據那根一律不開。"""
    hm = t.strftime("%H:%M")
    if hm in ("08:45", "09:45", "10:45", "11:45"):
        return True
    night = ("21:00", "22:00", "23:00") if us_dst(t) else ("22:00", "23:00", "00:00")
    return hm in night


def allowed_draft(t):
    return t.strftime("%H:%M") in ("08:45", "09:45", "10:45", "21:00", "22:00", "23:00")


def run(T, O, H, L, C, revised, timefilter=True, fallback60=False):
    n = len(C)
    E10, E20, E60 = ema(C, 10), ema(C, 20), ema(C, 60)
    macd = [a - b for a, b in zip(ema(C, 12), ema(C, 26))]
    hist = [a - b for a, b in zip(macd, ema(macd, 9))]
    MID, SD = [None] * n, [None] * n
    for i in range(19, n):
        w = C[i - 19:i + 1]
        MID[i], SD[i] = sum(w) / 20, statistics.pstdev(w)
    BBW = [4 * SD[i] / MID[i] if MID[i] else None for i in range(n)]
    ATR, a = [], None
    for i in range(n):
        tr = H[i] - L[i] if i == 0 else max(H[i] - L[i], abs(H[i] - C[i - 1]), abs(L[i] - C[i - 1]))
        a = tr if a is None else a + (tr - a) / 14
        ATR.append(a)

    def squeeze(i, look=100, pct=0.2, within=6):
        for j in range(i - within, i):
            w = [x for x in BBW[max(19, j - look):j + 1] if x]
            if len(w) < 40:
                return False
            if sorted(w).index(BBW[j]) <= pct * (len(w) - 1):
                return True
        return False

    trades, skipped, i = [], 0, 100
    while i < n - 1:
        ok = (allowed_revised(T[i]) if revised else allowed_draft(T[i])) if timefilter else True
        if revised and T[i].strftime("%H:%M") in ("12:45", "04:00"):
            ok = False
        side = 0
        for s in (1, -1):
            body, rng = (C[i] - O[i]) * s, H[i] - L[i]
            if revised:
                strong = body > 0 and body >= 0.5 * rng and body >= 0.5 * ATR[i - 1]
                mom = hist[i] * s > 0
            else:
                strong = body > 0 and body >= 0.6 * rng
                mom = hist[i] * s > 0 and (hist[i] - hist[i - 1]) * s > 0
            if ok and (E20[i] - E60[i]) * s > 0 and (C[i] - (MID[i] + 2 * SD[i] * s)) * s > 0 \
                    and strong and mom and squeeze(i):
                side = s
        if not side:
            i += 1
            continue
        e = O[i + 1]
        ext = L[i] if side > 0 else H[i]
        if revised:
            raw = abs(e - (ext - 0.1 * ATR[i] * side))
            if raw > 2.0 * ATR[i]:
                skipped += 1; i += 1; continue
            stop = e - min(raw, 1.5 * ATR[i]) * side
        else:
            stop = ext - 10 * side
            if abs(e - stop) > 90:
                skipped += 1
                if not fallback60:
                    i += 1; continue
                stop = e - 60 * side
        R = abs(e - stop)
        units, pnl, be, partial, j = 2, 0.0, False, False, i + 1
        while j < n:
            worst = L[j] if side > 0 else H[j]
            if (worst - stop) * side <= 0:  # 觸及停損；跳空時以開盤價成交
                px = stop if (O[j] - stop) * side > 0 else O[j]
                pnl += units * (px - e) * side; units = 0; break
            if not be:
                hit = (C[j] - e) * side >= R if revised else ((H[j] if side > 0 else L[j]) - e) * side >= R
                if hit:
                    be, stop = True, e + (COST * side if revised else 0)
            if j + 1 >= n:
                pnl += units * (C[j] - e) * side; units = 0; break
            active = be if revised else True
            if active and not partial and (C[j] - E10[j]) * side < 0:
                pnl += (O[j + 1] - e) * side; units -= 1; partial = True
            if active and (C[j] - E20[j]) * side < 0:
                pnl += units * (O[j + 1] - e) * side; units = 0; j += 1; break
            j += 1
        net = pnl - 2 * COST
        trades.append((T[i], side, round(R), round(net), round(net / (2 * R), 2)))
        i = j
    return trades, skipped


def report(name, trades, skipped):
    tot = sum(t[4] for t in trades)
    wins = sum(1 for t in trades if t[4] > 0.1)
    losses = sum(1 for t in trades if t[4] < -0.1)
    print(f"== {name}：交易 {len(trades)} 筆（勝 {wins}、負 {losses}），停損過遠 {skipped} 次，合計 {tot:+.2f}R")
    for t in trades:
        print(f"   {t[0]:%Y-%m-%d %H:%M} {'多' if t[1] > 0 else '空'} R={t[2]} 損益 {t[3]:+} 點 {t[4]:+.2f}R")


if __name__ == "__main__":
    bars = load(sys.argv[1] if len(sys.argv) > 1 else "tx_60min.csv")
    print("K 棒", len(bars[0]), "根：", bars[0][0], "~", bars[0][-1])
    report("原版（停損超過 90 點不做）", *run(*bars, revised=False))
    report("原版（超過 90 點改固定 60 點）", *run(*bars, revised=False, fallback60=True))
    report("修訂版", *run(*bars, revised=True))
    report("修訂版（無時段濾網）", *run(*bars, revised=True, timefilter=False))
