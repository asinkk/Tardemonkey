"""台指 60 分K 策略 v3：在 v2 的布林突破之外，加上日線趨勢濾網、趨勢中的拉回進場、多空不對稱與移動停利。
用法：python strategy/v3.py history/twii_60min.csv [台指期 60 分K csv]
"""
import datetime as dt, itertools, statistics as st, sys
from collections import defaultdict
from backtest import load, indicators, ema, twii_revised, tx_revised, COST

IS_START = dt.datetime(2023, 11, 3)   # 指標暖機後約從 2024/1 開始有訊號
SPLIT = dt.datetime(2026, 3, 7)       # 之前挑參數，之後驗證


def daily_regime(T, C, days, last_bar):
    """日線 EMA(days)，以日盤最後一根 K 的收盤當日收盤；第 i 根 K 只用已收完的日線（不偷看）。
    回傳 (日收盤, 日 EMA) 對應到每根 K 的前一個已完成交易日。"""
    closes, ema_v, out_c, out_e = [], None, [None] * len(C), [None] * len(C)
    k = 2 / (days + 1)
    for i, t in enumerate(T):
        out_c[i], out_e[i] = (closes[-1] if closes else None), ema_v
        if t.strftime("%H:%M") == last_bar:
            closes.append(C[i])
            ema_v = C[i] if ema_v is None else ema_v + k * (C[i] - ema_v)
    return out_c, out_e


def run(bars, ind, allowed, last_bar, p, start=None, end=None):
    T, O, H, L, C = bars
    n = len(C)
    E10, E20, E60, hist = ind["E10"], ind["E20"], ind["E60"], ind["hist"]
    MID, SD, ATR = ind["MID"], ind["SD"], ind["ATR"]
    if p["regime"]:
        key = ("reg", p["regime"], last_bar)
        if key not in ind:
            ind[key] = daily_regime(T, C, p["regime"], last_bar)
        DC, DE = ind[key]
    trades, i = [], 100
    while i < n - 1:
        sig = None
        ok = (start is None or T[i] >= start) and (end is None or T[i] < end) and (allowed is None or allowed(T[i]))
        if ok:
            for s in (1, -1):
                if s < 0 and p["shorts"] == "none":
                    continue
                trend = (E20[i] - E60[i]) * s > 0
                reg_ok = True
                if p["regime"] and DE[i] is not None:
                    reg_ok = (DC[i] - DE[i]) * s > 0
                    if s < 0 and p["shorts"] == "regime" and not reg_ok:
                        continue
                    if s > 0 and p["longs"] == "regime" and not reg_ok:
                        continue
                elif s < 0 and p["shorts"] == "regime":
                    continue
                if not trend:
                    continue
                body, rng = (C[i] - O[i]) * s, H[i] - L[i]
                # 1) 布林突破（v2）
                if body > 0 and body >= 0.5 * rng and body >= p["body_atr"] * ATR[i - 1] \
                        and (C[i] - (MID[i] + 2 * SD[i] * s)) * s > 0 and hist[i] * s > 0:
                    sig = (s, "突破"); break
                # 2) 拉回：趨勢中碰到 EMA20 後收回 EMA20 之上，且收在前一根高點之上
                if p["pullback"] and body > 0 and body >= 0.5 * rng \
                        and ((L[i] if s > 0 else H[i]) - E20[i]) * s <= 0 and (C[i] - E20[i]) * s > 0 \
                        and (C[i] - (H[i - 1] if s > 0 else L[i - 1])) * s > 0 \
                        and (E60[i] - E60[i - 5]) * s > 0:
                    sig = (s, "拉回"); break
        if not sig:
            i += 1
            continue
        side, kind = sig
        e = O[i + 1]
        ext = L[i] if side > 0 else H[i]
        raw = abs(e - (ext - 0.1 * ATR[i] * side))
        if raw > (p["stop_atr"] + 0.5) * ATR[i]:  # 追太高不做
            i += 1; continue
        stop = e - max(min(raw, p["stop_atr"] * ATR[i]), 0.5 * ATR[i]) * side
        R = abs(e - stop)
        pnl, be, j, best = 0.0, False, i + 1, e
        stop0, xt, xp, still_open = stop, None, None, False
        while j < n:
            worst = L[j] if side > 0 else H[j]
            if (worst - stop) * side <= 0:
                px = stop if (O[j] - stop) * side > 0 else O[j]
                pnl = (px - e) * side; xt, xp = T[j], px; break
            if not be and (C[j] - e) * side >= p["be_r"] * R:
                be, stop = True, e + COST * side
            if j + 1 >= n:
                pnl = (C[j] - e) * side; xt, xp, still_open = T[j], C[j], True; break
            if be:
                if p["exit"] == "ema20":
                    if (C[j] - E20[j]) * side < 0:
                        pnl = (O[j + 1] - e) * side; xt, xp = T[j + 1], O[j + 1]; j += 1; break
                else:  # chandelier k
                    best = max(best * side, C[j] * side) * side
                    trail = best - p["exit"] * ATR[j] * side
                    if (trail - stop) * side > 0:
                        stop = trail
                    lock = p.get("lock")   # 鎖利：最高收盤獲利達 lock[0] R 後，停損至少鎖住其中 lock[1]
                    if lock and (best - e) * side >= lock[0] * R:
                        floor = e + lock[1] * (best - e)
                        if (floor - stop) * side > 0:
                            stop = floor
            j += 1
        net = pnl - COST
        trades.append(dict(t=T[i], side=side, kind=kind, R=R, r=net / R, bars=j - i, entry=e, entry_t=T[i + 1],
                           stop0=stop0, stop=stop, exit_t=xt, exit=xp, open=still_open, be=be))
        i = j
    return trades


def stats(tr, months=None):
    if not tr:
        return dict(n=0, tot=0, avg=0, dd=0, pf=0, t=0, win=0)
    r = [x["r"] for x in tr]
    eq = pk = dd = 0
    for v in r:
        eq += v; pk = max(pk, eq); dd = min(dd, eq - pk)
    g = sum(v for v in r if v > 0); l = -sum(v for v in r if v < 0)
    t = st.mean(r) / st.stdev(r) * len(r) ** 0.5 if len(r) > 2 and st.stdev(r) else 0
    return dict(n=len(r), tot=sum(r), avg=st.mean(r), dd=dd, pf=g / l if l else 99, t=t,
                win=sum(v > 0.1 for v in r) / len(r))


def fmt(s):
    return (f"{s['n']:>3} 筆 勝率 {s['win']:.0%} 平均 {s['avg']:+.2f}R 合計 {s['tot']:+6.1f}R "
            f"回落 {s['dd']:5.1f}R PF {s['pf']:.2f} t {s['t']:.2f}")


GRID = dict(
    regime=[None, 20, 60, 120],
    longs=["all", "regime"],
    shorts=["all", "regime", "none"],
    pullback=[False, True],
    body_atr=[0.8],
    stop_atr=[1.5],
    be_r=[1.0, 1.5],
    exit=["ema20", 3.0, 4.0, 5.0, 6.0],
    lock=[None, (3.0, 0.33)],
)
V2 = dict(regime=None, longs="all", shorts="all", pullback=False, body_atr=0.8, stop_atr=1.5, be_r=1.0, exit="ema20", lock=None)
# v3（2026-11-07）：v2 + 拉回進場 + 保本後改用「最高收盤 − 5 ATR」移動停利
# 2026-11-13 加鎖利：最高收盤獲利達 3R 後，停損至少拉到「進場價 + 最高收盤獲利的 33%」
# （台指期 9/21 起 +9.6R → +11.8R；加權指數 3 年 +75.1R → +75.5R，勝率 26% → 32%）
V3 = dict(V2, pullback=True, exit=5.0, lock=(3.0, 0.33))
# v3 保守版：空單只在加權指數日收盤低於 120 日 EMA 時做（回落較小、交易較少）
V3_SAFE = dict(V3, regime=120, shorts="regime")

if __name__ == "__main__":
    b = load(sys.argv[1]); ind = indicators(*b)
    tx = None
    if len(sys.argv) > 2:
        tb = load(sys.argv[2]); tx = (tb, indicators(*tb))
    res = []
    for combo in itertools.product(*GRID.values()):
        p = dict(zip(GRID, combo))
        if p["regime"] is None and (p["longs"] == "regime"):
            continue
        tr = run(b, ind, twii_revised, "13:00", p, start=IS_START)
        ins = stats([x for x in tr if x["t"] < SPLIT]); oos = stats([x for x in tr if x["t"] >= SPLIT])
        res.append((p, ins, oos, stats(tr)))
    print("組合數", len(res))
    def show(p, ins, oos, al):
        print(p); print("   前段", fmt(ins)); print("   後段", fmt(oos)); print("   全部", fmt(al))
        if tx:
            print("   台指期1月", fmt(stats(run(tx[0], tx[1], tx_revised, "12:45", p))))
    for name, cfg in (("v2", V2), ("v3", V3), ("v3 保守版", V3_SAFE)):
        print(f"== {name}")
        for r in res:
            if r[0] == cfg: show(*r)
    print("== 前段合計 R 最高 12 組（前段至少 25 筆）")
    for r in sorted([r for r in res if r[1]["n"] >= 25], key=lambda r: -r[1]["tot"])[:12]:
        show(*r)
    print("== 各參數值平均（前段 / 後段 合計 R）")
    for k in GRID:
        for v in GRID[k]:
            sub = [r for r in res if r[0][k] == v]
            if sub:
                print(f"   {k}={v}: {sum(r[1]['tot'] for r in sub)/len(sub):+.1f} / {sum(r[2]['tot'] for r in sub)/len(sub):+.1f}  ({len(sub)})")
