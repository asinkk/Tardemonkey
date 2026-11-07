"""v3 改「進場後 5 小時平倉」的回測，跟原本 v3 比較。

規則：進場條件、初始停損跟 v3 完全一樣（v3.run 的訊號、追太高不做、0.5–1.5 ATR 停損）。
出場改成：
  - 初始停損仍有效，5 小時內碰到就先停損出場（不移保本、不移動停利）
  - 沒碰到停損，進場後滿 5 小時平倉
5 小時的算法（hours_mode）：
  - "trading"（預設）：只算開盤時間。台指期 60 分K 每根都是 60 分鐘，滿 5 根就平倉（第 5 根收盤價）；
    加權指數 13:00 那根只有 30 分鐘（13:00–13:30），累計滿 300 分鐘的那根收盤平倉，跨日時會變 5.5 小時
  - "calendar"：時鐘時間。進場時間 + 5 小時，若那時沒開盤（台指期 13:45–15:00、05:00–08:45、週末），
    就在下一根 K 開盤平倉
選項 be=True：1R 獲利後停損移到成本（v3 的保本），仍然 5 小時平倉。
用法：python strategy/v3_5h.py history/twii_60min.csv history/tx_1min [輸出交易紀錄 csv]
"""
import csv, datetime as dt, sys
from backtest import load, indicators, twii_revised, tx_revised, COST
from build_bars import build
import v3

CAP, MULT, FEE = 200000, 10, 25
TWO = dt.datetime(2024, 11, 7)   # 跟 v3 交易紀錄 CSV 同一段「近 2 年」


def dur_twii(t):
    return 30 if t.hour == 13 else 60


def dur_tx(t):
    return 60


def run(bars, ind, allowed, last_bar, dur, hours=5, hours_mode="trading", be_r=None, p=v3.V3, start=None):
    """訊號與初始停損逐字照 v3.run；只換出場。回傳格式跟 v3.run 一樣，另加 why（出場原因）。"""
    T, O, H, L, C = bars
    n = len(C)
    E20, E60, hist = ind["E20"], ind["E60"], ind["hist"]
    MID, SD, ATR = ind["MID"], ind["SD"], ind["ATR"]
    limit = hours * 60
    trades, i = [], 100
    while i < n - 1:
        sig = None
        if (start is None or T[i] >= start) and allowed(T[i]):
            for s in (1, -1):
                if (E20[i] - E60[i]) * s <= 0:
                    continue
                body, rng = (C[i] - O[i]) * s, H[i] - L[i]
                if body > 0 and body >= 0.5 * rng and body >= p["body_atr"] * ATR[i - 1] \
                        and (C[i] - (MID[i] + 2 * SD[i] * s)) * s > 0 and hist[i] * s > 0:
                    sig = (s, "突破"); break
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
        if raw > (p["stop_atr"] + 0.5) * ATR[i]:
            i += 1; continue
        stop = e - max(min(raw, p["stop_atr"] * ATR[i]), 0.5 * ATR[i]) * side
        R = abs(e - stop)
        stop0, et = stop, T[i + 1]
        target = et + dt.timedelta(hours=hours)
        held, j, be, still_open = 0, i + 1, False, False
        while j < n:
            if hours_mode == "calendar" and T[j] >= target:   # 目標時間沒開盤，下一根開盤平倉
                xt, xp, why = T[j], O[j], "滿5小時(下一根開盤)"; break
            worst = L[j] if side > 0 else H[j]
            if (worst - stop) * side <= 0:
                xp = stop if (O[j] - stop) * side > 0 else O[j]
                xt, why = T[j], "保本" if be else "停損"; break
            held += dur(T[j])
            end_t = T[j] + dt.timedelta(minutes=dur(T[j]))
            if (hours_mode == "trading" and held >= limit) or (hours_mode == "calendar" and end_t >= target):
                xt, xp, why = end_t, C[j], "滿5小時"; break
            if be_r and not be and (C[j] - e) * side >= be_r * R:
                be, stop = True, e + COST * side
            if j + 1 >= n:
                xt, xp, why, still_open = end_t, C[j], "持倉中", True; break
            j += 1
        net = (xp - e) * side - COST
        trades.append(dict(t=T[i], side=side, kind=kind, R=R, r=net / R, bars=j - i, entry=e, entry_t=et,
                           stop0=stop0, stop=stop, exit_t=xt, exit=xp, open=still_open, be=be, why=why))
        i = j
    return trades


def money(tr):
    eq = pk = CAP; dd = worst = 0; rows = []
    for x in tr:
        pts = x["r"] * x["R"] + COST          # 引擎扣的 COST 點加回，改扣實際費用
        fee = FEE if x["open"] else 2 * FEE
        twd = pts * MULT - fee
        eq += twd; pk = max(pk, eq); dd = min(dd, eq - pk); worst = min(worst, twd)
        rows.append(dict(x, pts=pts, fee=fee, twd=twd, eq=eq))
    return dict(eq=eq, dd=dd, worst=worst, fees=sum(r["fee"] for r in rows), rows=rows,
                unreal=sum(r["twd"] for r in rows if r["open"]))


def line(name, tr):
    s, m = v3.stats(tr), money(tr)
    hold = sum((x["exit_t"] - x["entry_t"]).total_seconds() for x in tr) / 3600 / len(tr) if tr else 0
    stops = sum(x.get("why") == "停損" for x in tr)
    top2 = sum(sorted(x["r"] for x in tr)[-2:]) if tr else 0
    print(f"{name:26s} {s['n']:3d} 筆 勝率 {s['win']:.0%} 合計 {s['tot']:+6.1f}R 回落 {s['dd']:5.1f}R PF {s['pf']:.2f} t {s['t']:.2f} "
          f"前2筆 {top2:+5.1f}R 停損 {stops:2d} 筆 平均持有 {hold:5.1f} 時 | 期末 {m['eq']:>9,.0f}（未實現 {m['unreal']:+,.0f}）"
          f" 回落 {m['dd']:>8,.0f} 單筆最大虧 {m['worst']:>7,.0f} 費用 {m['fees']:,.0f}")
    return m


if __name__ == "__main__":
    hb = load(sys.argv[1]); hind = indicators(*hb)
    rows = build(sys.argv[2], verbose=False)
    tb = ([r[0] for r in rows], *([r[k] for r in rows] for k in range(1, 5))); tind = indicators(*tb)

    tw = {
        "v3（原版）": [x for x in v3.run(hb, hind, twii_revised, "13:00", v3.V3, start=v3.IS_START)],
        "5h 交易時間": run(hb, hind, twii_revised, "13:00", dur_twii, start=v3.IS_START),
        "5h 時鐘時間": run(hb, hind, twii_revised, "13:00", dur_twii, hours_mode="calendar", start=v3.IS_START),
        "5h 交易時間＋1R保本": run(hb, hind, twii_revised, "13:00", dur_twii, be_r=1.0, start=v3.IS_START),
    }
    for title, cut, end in (("== 加權指數 3 年（2023/11 起，約 2024/1 起有訊號）", v3.IS_START, None),
                            (f"== 加權指數近 2 年 {TWO:%Y/%m/%d} 起", TWO, None),
                            (f"== 加權指數挑參數段（{v3.SPLIT:%Y/%m/%d} 前）", v3.IS_START, v3.SPLIT),
                            (f"== 加權指數驗證段 {v3.SPLIT:%Y/%m/%d} 起", v3.SPLIT, None)):
        print(title)
        for k, tr in tw.items():
            line(k, [x for x in tr if x["t"] >= cut and (end is None or x["t"] < end)])

    print(f"== 台指期日夜盤 60 分K {tb[0][0]:%Y/%m/%d}–{tb[0][-1]:%m/%d}（history/tx_1min，前 100 根暖機）")
    txr = {
        "v3（原版）": v3.run(tb, tind, tx_revised, "12:45", v3.V3),
        "5h 交易時間": run(tb, tind, tx_revised, "12:45", dur_tx),
        "5h 時鐘時間": run(tb, tind, tx_revised, "12:45", dur_tx, hours_mode="calendar"),
        "5h 交易時間＋1R保本": run(tb, tind, tx_revised, "12:45", dur_tx, be_r=1.0),
    }
    for k, tr in txr.items():
        line(k, tr)

    print("== 持有時數敏感度（交易時間，不保本）")
    for h in (2, 3, 4, 5, 6, 8, 10):
        a = [x for x in run(hb, hind, twii_revised, "13:00", dur_twii, hours=h, start=v3.IS_START) if x["t"] >= TWO]
        b = run(tb, tind, tx_revised, "12:45", dur_tx, hours=h)
        sa, sb = v3.stats(a), v3.stats(b)
        print(f"  {h:2d} 小時  加權近2年 {sa['n']:3d} 筆 {sa['tot']:+6.1f}R 期末 {money(a)['eq']:>9,.0f} 回落 {money(a)['dd']:>8,.0f}"
              f" | 台指期 {sb['n']:2d} 筆 {sb['tot']:+5.1f}R 期末 {money(b)['eq']:>9,.0f}")

    print("== 台指期 5h（交易時間）明細")
    for x in money(txr["5h 交易時間"])["rows"]:
        print(f"  {x['entry_t']:%m/%d %H:%M} {'多' if x['side'] > 0 else '空'}・{x['kind']} 進 {x['entry']:.0f} 停損 {x['R']:.0f} 點 "
              f"出 {x['exit_t']:%m/%d %H:%M} {x['exit']:.0f} {x['why']} {x['r']:+.2f}R {x['twd']:+,.0f} 元")
    print("== 台指期 v3 原版明細")
    for x in money(txr["v3（原版）"])["rows"]:
        print(f"  {x['entry_t']:%m/%d %H:%M} {'多' if x['side'] > 0 else '空'}・{x['kind']} 進 {x['entry']:.0f} 停損 {x['R']:.0f} 點 "
              f"出 {x['exit_t']:%m/%d %H:%M} {x['exit']:.0f} {'持倉中' if x['open'] else ''} {x['r']:+.2f}R {x['twd']:+,.0f} 元")

    if len(sys.argv) > 3:
        with open(sys.argv[3], "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["資料", "序號", "類型", "方向", "訊號K", "進場時間", "進場價", "初始停損", "停損點數", "出場時間", "出場價", "出場原因",
                        "損益點(毛)", "費用(元)", "損益(元)", "累計帳戶(元)", "R倍數"])
            for src, tr in (("台指期日夜盤（實際交易依據）", txr["5h 交易時間"]),
                            ("加權指數近2年（回測用）", [x for x in tw["5h 交易時間"] if x["t"] >= TWO])):
                for k, x in enumerate(money(tr)["rows"], 1):
                    w.writerow([src, k, x["kind"], "多" if x["side"] > 0 else "空", f"{x['t']:%Y-%m-%d %H:%M}", f"{x['entry_t']:%Y-%m-%d %H:%M}",
                                f"{x['entry']:.0f}", f"{x['stop0']:.0f}", f"{x['R']:.0f}", f"{x['exit_t']:%Y-%m-%d %H:%M}", f"{x['exit']:.0f}",
                                x["why"], f"{x['pts']:.0f}", x["fee"], f"{x['twd']:.0f}", f"{x['eq']:.0f}", f"{x['r']:.2f}"])
        print("輸出", sys.argv[3])
