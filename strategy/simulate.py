"""用固定口數試算帳戶金額：預設 20 萬元、1 口微台（每點 10 元）、每次下單 25 元（進出各一次）。
用法：python strategy/simulate.py history/twii_60min.csv [台指期 60 分K csv] [本金] [每點價值] [每邊費用]
"""
import datetime as dt, sys
from collections import defaultdict
from backtest import load, indicators, twii_revised, tx_revised, COST
from v3 import run, V2, V3, V3_SAFE, IS_START

CAP = float(sys.argv[3]) if len(sys.argv) > 3 else 200000
MULT = float(sys.argv[4]) if len(sys.argv) > 4 else 10
FEE = float(sys.argv[5]) if len(sys.argv) > 5 else 25
TWO = dt.datetime(2024, 11, 7)


def sim(tr, label):
    eq = pk = CAP; dd = ddp = worst = 0; streak = maxstreak = 0; yr = defaultdict(float)
    for x in tr:
        twd = (x["r"] * x["R"] + COST) * MULT - 2 * FEE   # 引擎扣的 COST 點加回，改扣實際費用
        eq += twd; pk = max(pk, eq); dd = min(dd, eq - pk); ddp = min(ddp, (eq - pk) / pk); worst = min(worst, twd)
        streak = streak + 1 if twd < 0 else 0; maxstreak = max(maxstreak, streak)
        yr[x["t"].year] += twd
    print(f"{label}: {len(tr)} 筆，期末 {eq:,.0f} 元（{eq / CAP - 1:+.1%}），最大回落 {dd:,.0f} 元（{ddp:.1%}），"
          f"單筆最大虧損 {worst:,.0f} 元，最多連虧 {maxstreak} 筆，費用 {len(tr) * 2 * FEE:,.0f} 元")
    print("   各年損益", {k: round(v) for k, v in sorted(yr.items())})


def buy_hold(bars, start, label):
    T, C = bars[0], bars[4]
    i0 = next(i for i, t in enumerate(T) if t >= start)
    pk = C[i0]; dd = 0
    for c in C[i0:]:
        pk = max(pk, c); dd = min(dd, (c - pk) * MULT)
    pnl = (C[-1] - C[i0]) * MULT - 2 * FEE
    print(f"{label}: 期末 {CAP + pnl:,.0f} 元（{pnl / CAP:+.1%}），最大回落 {dd:,.0f} 元（約 {dd / CAP:.0%} 本金）")


if __name__ == "__main__":
    b = load(sys.argv[1]); ind = indicators(*b)
    tx = None
    if len(sys.argv) > 2:
        tb = load(sys.argv[2]); tx = (tb, indicators(*tb))
    for nm, p in (("v3", V3), ("v3 保守版", V3_SAFE), ("v2", V2)):
        tr = [x for x in run(b, ind, twii_revised, "13:00", p, start=IS_START) if x["t"] >= TWO]
        sim(tr, f"{nm} 近 2 年")
        if tx:
            sim(run(tx[0], tx[1], tx_revised, "12:45", p), f"{nm} 台指期 1 個月")
    buy_hold(b, TWO, "對照：同期 1 口微台買進持有")
