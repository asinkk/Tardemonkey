"""修訂版參數搜尋：前段（in-sample）挑參數，後段（out-of-sample）驗證，避免只是把過去的運氣調進去。
用法：python strategy/optimize.py history/twii_60min.csv [台指期 60 分K csv]
"""
import datetime as dt, itertools, sys
from backtest import load, indicators, run, summary, twii_revised, tx_revised

START = dt.datetime(2024, 11, 7)   # 近 2 年
SPLIT = dt.datetime(2026, 3, 7)    # 前 16 個月調參數，後 8 個月驗證

GRID = dict(
    squeeze=[None, 0.2, 0.35],          # None = 不要求擠壓
    stop_atr=[1.5, 2.0, 2.5],
    be_r=[None, 1.0, 1.5],
    exit=["split", "ema20", "chand2.5", "chand3.5", "target2", "target3"],
    body_atr=[0.3, 0.5, 0.8],
)


def kwargs(p):
    kw = dict(squeeze_on=p["squeeze"] is not None, squeeze_pct=p["squeeze"] or 0.2,
              stop_atr=p["stop_atr"], max_atr=p["stop_atr"] + 0.5, be_r=p["be_r"], body_atr=p["body_atr"])
    ex = p["exit"]
    if ex.startswith("chand"):
        kw.update(exit_mode="chand", chand_k=float(ex[5:]))
    elif ex.startswith("target"):
        kw.update(exit_mode="target", target_r=float(ex[6:]))
    else:
        kw.update(exit_mode=ex)
    return kw


def fmt(s):
    if not s["n"]:
        return "0 筆"
    return f"{s['n']:>3} 筆 勝率 {s['win']:.0%} 平均 {s['avg']:+.2f}R 合計 {s['tot']:+6.1f}R 回落 {s['dd']:5.1f}R PF {s['pf']:.2f}"


if __name__ == "__main__":
    bars = load(sys.argv[1]); ind = indicators(*bars)
    tx = None
    if len(sys.argv) > 2:
        txb = load(sys.argv[2]); tx = (txb, indicators(*txb))
    keys = list(GRID)
    results = []
    for combo in itertools.product(*GRID.values()):
        p = dict(zip(keys, combo)); kw = kwargs(p)
        ins = summary(run(bars, ind, True, twii_revised, start=START, end=SPLIT, **kw)[0])
        oos = summary(run(bars, ind, True, twii_revised, start=SPLIT, **kw)[0])
        results.append((p, ins, oos))
    ok = [r for r in results if r[1]["n"] >= 12]
    ok.sort(key=lambda r: r[1]["tot"], reverse=True)
    base = dict(squeeze=0.2, stop_atr=1.5, be_r=1.0, exit="split", body_atr=0.5)
    print("組合數", len(results))
    for label, rows in (("目前修訂版", [r for r in results if r[0] == base]), ("前段最佳 15 組", ok[:15])):
        print("==", label)
        for p, ins, oos in rows:
            extra = ""
            if tx:
                extra = " | 台指期1個月 " + fmt(summary(run(tx[0], tx[1], True, tx_revised, **kwargs(p))[0]))
            print(p, "\n   前段", fmt(ins), "\n   後段", fmt(oos), extra)
    # 每個參數單獨看：其餘參數平均下，各值在前段/後段的平均合計 R
    print("== 各參數值的平均表現（前段 / 後段，合計 R）")
    for k in keys:
        for v in GRID[k]:
            sub = [r for r in results if r[0][k] == v]
            a = sum(r[1].get("tot", 0) for r in sub) / len(sub); b = sum(r[2].get("tot", 0) for r in sub) / len(sub)
            print(f"   {k}={v}: {a:+.1f} / {b:+.1f}")
