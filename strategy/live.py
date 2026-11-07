"""即時策略看板：台指期 v3（60 分K）與 D2（日K）目前的訊號、部位、停損與進場條件。

資料：
  - history/tx_1min/：每天台北 06:30 存的期交所每筆成交（轉成 1 分K）
  - 期交所 MIS 即時 1 分K（getChartData1M）：補上目前或最近一個日盤、夜盤還沒存進 history 的部分
  - history/twii_60min.csv：加權指數 60 分K，只給 D2 日K 指標暖機用
D2 是用加權指數日盤日K 調的參數，所以台指期這邊也只用日盤（08:45–13:45）合成日K，夜盤不算。
策略規則完全沿用 v3.run（V3）與 daily.run（D2），這裡只負責接資料、判斷最後一根 K 是否已收完、
把結果整理成網頁。條件明細（✓/✗）是顯示用的重算，判斷進出場以 run() 的結果為準。

用法：python strategy/live.py [--src history/tx_1min] [--twii history/twii_60min.csv] [--out docs]
                             [--record history/forward_d2_trades.csv] [--offline]
"""
import argparse, bisect, csv, datetime as dt, glob, html, json, os, shutil, sys, tempfile, urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_bars import build
from backtest import load, indicators, tx_revised, us_dst, COST
import v3, daily
import trades_page

MIS = "https://mis.taifex.com.tw/futures/api/"
MONTHS = "ABCDEFGHIJKL"
FORWARD_START = dt.datetime(2026, 11, 9, 8, 45)   # 兩套規則都在 11/7 定案，11/9 起算前測
CAP, MULT, FEE = 200000, 10, 25                   # 本金 20 萬、1 口微台（每點 10 元）、每次下單 25 元


def now_tpe():
    return dt.datetime.now(dt.timezone.utc).replace(tzinfo=None, microsecond=0) + dt.timedelta(hours=8)


# ---------- 即時 1 分K ----------
def post(ep, body):
    req = urllib.request.Request(MIS + ep, json.dumps(body).encode(),
                                 {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0 (txo-dashboard)"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def contract_code(sym, year_now):
    """TXFK6-F → 202611（月份字母 A–L，年份個位數）。"""
    y = year_now // 10 * 10 + int(sym[4])
    if y < year_now - 1:
        y += 10
    return f"{y}{MONTHS.index(sym[3]) + 1:02d}"


def fetch_live(now):
    """回傳 (sessions, info)。sessions：[(開始時間, {契約: {datetime: [o,h,l,c,v]}})]，每個盤別一筆；
    每個盤別抓成交量前兩大的契約，換月時 build_bars 才算得出新舊月價差。"""
    sessions, info = [], {}
    for mt, name in (("0", "日盤"), ("1", "夜盤")):
        try:
            q = post("getQuoteList", {"MarketType": mt, "SymbolType": "F", "KindID": "1", "CID": "TXF", "ExpireMonth": "",
                                      "RowSize": "全部", "PageNo": "", "SortColumn": "", "AscDesc": "A"})["RtData"]["QuoteList"]
        except Exception as e:
            info[name] = {"error": f"報價清單抓取失敗：{e}"}
            continue
        futs = [x for x in q if x.get("SymbolID", "").startswith("TXF") and len(x["SymbolID"]) > 4 and x["SymbolID"][3] in MONTHS]
        futs.sort(key=lambda x: -int(float(x.get("CTotalVolume") or 0)))
        by_contract, start = {}, None
        for k, f in enumerate(futs[:2]):
            try:
                d = post("getChartData1M", {"SymbolID": f["SymbolID"]})["RtData"]
            except Exception as e:
                info.setdefault(name, {"error": f"1 分K 抓取失敗：{e}"})
                continue
            base = dt.datetime.strptime(d["Quote"]["CDate"], "%Y%m%d")
            bars = {}
            for t, o, h, l, c, v in d.get("Ticks") or []:
                hh, mm = int(t[:2]), int(t[2:4])
                day = base + dt.timedelta(days=1) if mt == "1" and hh < 15 else base
                ts = day + dt.timedelta(minutes=hh * 60 + mm - 1)   # MIS 標的是這分鐘的結束時間
                o, h, l, c, v = float(o), float(h), float(l), float(c), int(float(v))
                b = bars.get(ts)
                if b is None:
                    bars[ts] = [o, h, l, c, v]
                else:
                    b[1] = max(b[1], h); b[2] = min(b[2], l); b[3] = c; b[4] += v
            if not bars:
                continue
            by_contract[contract_code(f["SymbolID"], now.year)] = bars
            if k == 0:
                start = min(bars)
                info[name] = {"symbol": f["SymbolID"], "name": f.get("DispCName"), "last": float(d["Quote"]["CLastPrice"] or 0),
                              "ref": float(d["Quote"]["CRefPrice"] or 0), "date": d["Quote"]["CDate"],
                              "last_minute": f"{max(bars):%Y-%m-%d %H:%M}", "minutes": len(bars)}
        if by_contract and start:
            sessions.append((start, by_contract))
    sessions.sort(key=lambda s: s[0])
    return sessions, info


def merged_dir(src, sessions):
    """把 history 的檔案連到暫存資料夾，再把 history 還沒有的即時分K 依盤別各寫成一個檔案，接在最後。"""
    files = sorted(glob.glob(os.path.join(src, "Daily_*.csv")))
    if not files:
        sys.exit(f"{src} 沒有 Daily_*.csv")
    last = None
    for f in files[-2:]:
        for r in csv.DictReader(open(f, encoding="utf-8")):
            t = dt.datetime.strptime(r["date"] + r["time"], "%Y%m%d%H%M")
            last = t if last is None or t > last else last
    tmp = tempfile.mkdtemp(prefix="txlive_")
    for f in files:
        shutil.copy(f, tmp)
    added, stem = 0, os.path.basename(files[-1])[:-4]
    for k, (_, by_contract) in enumerate(sessions, 1):
        rows = [(c, t, b) for c, bars in by_contract.items() for t, b in bars.items() if t > last]
        if not rows:
            continue
        rows.sort(key=lambda x: (x[0], x[1]))
        with open(os.path.join(tmp, f"{stem}_live{k}.csv"), "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["date", "time", "contract", "open", "high", "low", "close", "volume"])
            for c, t, (o, h, l, cl, v) in rows:
                w.writerow([f"{t:%Y%m%d}", f"{t:%H%M}", c, f"{o:g}", f"{h:g}", f"{l:g}", f"{cl:g}", v])
        front = max(by_contract, key=lambda x: len(by_contract[x]))
        added += sum(1 for c, *_ in rows if c == front)
    return tmp, last, added


# ---------- K 棒 ----------
def with_dummy(bars, step):
    """在最後面多接一根「開高低收都等於最後收盤」的假 K，讓 run() 也會檢查最後一根 K 的訊號與出場。"""
    T, O, H, L, C = bars
    c = C[-1]
    return (T + [T[-1] + step], O + [c], H + [c], L + [c], C + [c])


def evaluate(bars, runner, step):
    ind = indicators(*bars)
    trades = runner(bars, ind)
    b2 = with_dummy(bars, step)
    trades2 = runner(b2, indicators(*b2))
    pos = next((x for x in trades if x["open"]), None)
    sig = next((x for x in trades2 if x["t"] == bars[0][-1]), None)
    pending_exit = None
    if pos:
        y = next((x for x in trades2 if x["t"] == pos["t"]), None)
        if y and not y["open"]:
            pending_exit = y
    return dict(ind=ind, trades=trades, pos=pos, sig=sig, pending_exit=pending_exit)


# ---------- 進出場時間（精確到分鐘） ----------
def exit_reason_v3(x):
    why = "移動停利" if (x["stop"] - x["entry"]) * x["side"] > COST else "保本停損" if x["be"] else "初始停損"
    return why + ("（跳空成交）" if round(x["exit"]) != round(x["stop"]) else "")


def first_touch(mT, m, start, end, stop, side):
    """[start, end) 之間第一根碰到停損的 1 分K 時間。"""
    i = bisect.bisect_left(mT, start)
    while i < len(m) and mT[i] < end:
        if (((m[i][3] if side > 0 else m[i][2]) - stop) * side) <= 0:
            return mT[i]
        i += 1
    return None


def refine_times(trades, strat, m, day_open=dt.time(8, 45), day_close=dt.time(13, 45)):
    """run() 的時間是 K 棒開始時間（日K 只有日期）。這裡改成實際進出場的時點：
    - D2 進場：訊號當天收盤（台指期 13:45；加權指數 13:30）
    - 跌破/站上 EMA 出場：隔天開盤
    - 停損出場：用更細的 K 棒（台指期 1 分K、加權指數 60 分K）找第一次碰到停損的時間
    m 是 [(開始時間, 開, 高, 低, 收, ...)]，價格要跟 trades 同一套（換月調整後）。"""
    mT = [r[0] for r in m]
    for x in trades:
        if strat == "d2":
            x["entry_t"] = dt.datetime.combine(x["entry_t"].date(), day_close)
        else:
            x.setdefault("why", "" if x["open"] else exit_reason_v3(x))
        if x["open"]:
            continue
        if strat == "d2":
            d = x["exit_t"].date()
            if x["why"].startswith("跌破"):
                x["exit_t"] = dt.datetime.combine(d, day_open)
                continue
            start, end = dt.datetime.combine(d, day_open), dt.datetime.combine(d, day_close) + dt.timedelta(minutes=1)
        else:
            start, end = x["exit_t"], x["exit_t"] + dt.timedelta(minutes=60)
        t = first_touch(mT, m, start, end, x["stop"], x["side"])
        if t:
            x["exit_t"] = t
    return trades


# ---------- 條件明細（顯示用） ----------
def ck(name, ok, detail=""):
    return {"name": name, "ok": bool(ok), "detail": detail}


def proj_stop(e, ext, atr, side, stop_atr, max_atr):
    raw = abs(e - (ext - 0.1 * atr * side))
    stop = e - max(min(raw, stop_atr * atr), 0.5 * atr) * side
    return raw <= max_atr * atr, stop, raw


def v3_checks(bars, ind, i):
    T, O, H, L, C = bars
    E20, E60, hist, MID, SD, ATR = (ind[k] for k in ("E20", "E60", "hist", "MID", "SD", "ATR"))
    p, out = v3.V3, {}
    for s, side in ((1, "多"), (-1, "空")):
        body, rng = (C[i] - O[i]) * s, H[i] - L[i]
        strong = body > 0 and body >= 0.5 * rng
        bb = MID[i] + 2 * SD[i] * s
        prev = H[i - 1] if s > 0 else L[i - 1]
        gt = ">" if s > 0 else "<"
        ok_stop, stop, raw = proj_stop(C[i], L[i] if s > 0 else H[i], ATR[i], s, p["stop_atr"], p["stop_atr"] + 0.5)
        out[side] = {
            "common": [ck("可進場時段", tx_revised(T[i]), f"{T[i]:%H:%M} 這根"),
                       ck(f"趨勢 EMA20 {gt} EMA60", (E20[i] - E60[i]) * s > 0, f"{E20[i]:,.0f} / {E60[i]:,.0f}")],
            "突破": [ck(f"{'紅' if s > 0 else '黑'}K 實體 ≥ 高低差一半", strong, f"實體 {(C[i] - O[i]):+,.0f}、高低差 {rng:,.0f}"),
                     ck("實體 ≥ 0.8 ATR", body >= p["body_atr"] * ATR[i - 1], f"需 {p['body_atr'] * ATR[i - 1]:,.0f} 點"),
                     ck(f"收盤{'站上布林上軌' if s > 0 else '跌破布林下軌'}", (C[i] - bb) * s > 0, f"{'上' if s > 0 else '下'}軌 {bb:,.0f}"),
                     ck(f"MACD 柱{'翻正' if s > 0 else '翻負'}", hist[i] * s > 0, f"{hist[i]:+,.1f}")],
            "拉回": [ck(f"{'紅' if s > 0 else '黑'}K 實體 ≥ 高低差一半", strong, ""),
                     ck(f"{'低點碰到' if s > 0 else '高點碰到'} EMA20", ((L[i] if s > 0 else H[i]) - E20[i]) * s <= 0, f"EMA20 {E20[i]:,.0f}"),
                     ck(f"收盤{'站回' if s > 0 else '壓回'} EMA20", (C[i] - E20[i]) * s > 0, ""),
                     ck(f"收盤{'過前一根高點' if s > 0 else '破前一根低點'}", (C[i] - prev) * s > 0, f"{prev:,.0f}"),
                     ck(f"EMA60 比 5 根前{'高' if s > 0 else '低'}", (E60[i] - E60[i - 5]) * s > 0, f"{E60[i] - E60[i - 5]:+,.0f}")],
            "stop": {"ok": ok_stop, "stop": stop, "raw": raw},
        }
    return out


def d2_checks(bars, ind, i):
    T, O, H, L, C = bars
    E20, E60, ATR = ind["E20"], ind["E60"], ind["ATR"]
    p, out = daily.D2, {}
    N = p["entry"]
    for s, side in ((1, "多"), (-1, "空")):
        body, rng = (C[i] - O[i]) * s, H[i] - L[i]
        lvl = max(C[i - N:i]) if s > 0 else min(C[i - N:i])
        gt = ">" if s > 0 else "<"
        ok_stop, stop, raw = proj_stop(C[i], L[i] if s > 0 else H[i], ATR[i], s, p["stop_atr"], p.get("max_atr", p["stop_atr"] + 0.5))
        out[side] = {
            "common": [ck(f"趨勢 日EMA20 {gt} EMA60", (E20[i] - E60[i]) * s > 0, f"{E20[i]:,.0f} / {E60[i]:,.0f}")],
            "突破": [ck(f"收盤創 {N} 日新{'高' if s > 0 else '低'}", (C[i] - lvl) * s > 0, f"需{'高於' if s > 0 else '低於'} {lvl:,.0f}"),
                     ck(f"{'紅' if s > 0 else '黑'}K 實體 ≥ 高低差一半", body > 0 and body >= 0.5 * rng,
                        f"開 {O[i]:,.0f}、實體 {(C[i] - O[i]):+,.0f}、高低差 {rng:,.0f}"),
                     ck("停損距離 ≤ 1.5 ATR（不追太遠）", ok_stop, f"{raw:,.0f} 點 / 上限 {1.5 * ATR[i]:,.0f}")],
            "stop": {"ok": ok_stop, "stop": stop, "raw": raw},
            "level": lvl,
        }
    return out


# ---------- 輸出 ----------
def twd(x):
    return (x["r"] * x["R"] + COST) * MULT - (FEE if x["open"] else 2 * FEE)


def why_text(x):
    """daily.run 的出場原因一律寫「跌破」，空單其實是收盤站上 EMA。"""
    w = x.get("why", "")
    return w.replace("跌破", "站上") if x["side"] < 0 else w


def trade_row(x, last_c=None):
    """last_c：持倉中的交易用這個價格算未實現損益（台指期用現價；None 表示用 run() 最後一根收盤）。"""
    if x["open"] and last_c is not None:
        pts = (last_c - x["entry"]) * x["side"]
        x = dict(x, r=(pts - COST) / x["R"])
    return {"signal": f"{x['t']:%Y-%m-%d %H:%M}", "kind": x["kind"], "side": "多" if x["side"] > 0 else "空",
            "entry_t": f"{x['entry_t']:%Y-%m-%d %H:%M}", "entry": round(x["entry"]), "stop0": round(x["stop0"]),
            "stop": round(x["stop"]), "exit_t": "" if x["open"] else f"{x['exit_t']:%Y-%m-%d %H:%M}",
            "exit": "" if x["open"] else round(x["exit"]), "r": round(x["r"], 2), "pts": round(x["r"] * x["R"] + COST),
            "twd": round(twd(x)), "open": x["open"], "be": x["be"], "why": why_text(x),
            "phase": "前測" if x["t"] >= FORWARD_START.replace(hour=0, minute=0) else "回測"}


def write_record(path, trades, last_c):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["階段", "訊號K", "類型", "方向", "進場時間", "進場價", "初始停損", "目前/出場停損", "出場時間", "出場價",
                    "出場原因", "R", "損益點", "損益元(1口微台)", "狀態"])
        for x in trades:
            r = trade_row(x, last_c)
            w.writerow([r["phase"], r["signal"], r["kind"], r["side"], r["entry_t"], r["entry"], r["stop0"], r["stop"],
                        r["exit_t"], r["exit"], r["why"], f"{r['r']:+.2f}", f"{r['pts']:+d}", f"{r['twd']:+d}",
                        "持倉中" if r["open"] else "已出場"])


def main():
    ap = argparse.ArgumentParser()
    here = os.path.dirname(os.path.abspath(__file__))
    ap.add_argument("--src", default=os.path.join(here, "..", "history", "tx_1min"))
    ap.add_argument("--twii", default=os.path.join(here, "..", "history", "twii_60min.csv"))
    ap.add_argument("--out", default=os.path.join(here, "..", "docs"))
    ap.add_argument("--record", help="另存 D2 交易紀錄 CSV 的路徑")
    ap.add_argument("--offline", action="store_true", help="不抓即時分K，只用 history")
    ap.add_argument("--now", help="測試用：假裝現在是台北時間 YYYY-mm-ddTHH:MM")
    a = ap.parse_args()

    now = dt.datetime.strptime(a.now, "%Y-%m-%dT%H:%M") if a.now else now_tpe()
    sessions, info = ([], {}) if a.offline else fetch_live(now)
    tmp, hist_last, added = merged_dir(a.src, sessions)
    try:
        rows = build(tmp, verbose=False)
        m1 = build(tmp, 1, verbose=False)   # 換月調整後的 1 分K，找停損觸價時間用
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    last_min = max(r[0] for r in rows)  # 60 分K 的開始時間；真正最後一分鐘看 info
    T = [r[0] for r in rows]
    hb = (T, *([r[k] for r in rows] for k in range(1, 5)))
    last_c = hb[4][-1]

    # v3：最後一根 60 分K 收完了沒
    v3_end = T[-1] + dt.timedelta(minutes=60)
    v3_forming = now < v3_end
    hv = evaluate(hb, lambda b, ind: v3.run(b, ind, tx_revised, "12:45", v3.V3), dt.timedelta(minutes=60))
    v3c = v3_checks(hb, hv["ind"], len(T) - 1)

    # D2：台指期「只用日盤」的日K（跟調參用的加權指數日K 一樣是 1 個日盤一根），接在加權指數日K 後面暖機
    txd = daily.tx_daily(rows, night=False)
    tw = daily.twii_daily(load(a.twii))
    sp, i0, off = daily.splice(tw, txd)
    dlast = sp[0][-1]
    d_close = dt.datetime.combine(dlast.date(), dt.time(13, 45))
    d_forming = now < d_close
    hd = evaluate(sp, lambda b, ind: daily.run(b, ind, daily.D2, sig_from=i0), dt.timedelta(days=1))
    d2c = d2_checks(sp, hd["ind"], len(sp[0]) - 1)

    refine_times(hv["trades"], "v3", m1)
    refine_times(hd["trades"], "d2", m1)

    if a.record:
        write_record(a.record, hd["trades"], last_c)

    def pos_info(res, strat):
        x = res["pos"]
        if not x:
            return None
        pts = (last_c - x["entry"]) * x["side"]
        R = x["R"]
        d = {"side": "多" if x["side"] > 0 else "空", "kind": x["kind"], "entry_t": f"{x['entry_t']:%Y-%m-%d %H:%M}",
             "entry": round(x["entry"]), "stop0": round(x["stop0"]), "stop": round(x["stop"]), "be": x["be"], "R": round(R),
             "pts": round(pts), "twd": round(pts * MULT - FEE), "to_stop": round((last_c - x["stop"]) * x["side"]),
             "be_at": round(x["entry"] + 1.0 * R * x["side"])}
        ind = res["ind"]
        if strat == "v3":
            d["trail_note"] = (f"保本後停損改成「進場後{'最高收盤 − 5 ATR' if x['side'] > 0 else '最低收盤 + 5 ATR'}」，只往有利方向移動；"
                               f"目前 ATR {ind['ATR'][-1]:,.0f}")
        else:
            d["trail_note"] = (f"保本後，收盤{'跌破' if x['side'] > 0 else '站上'}日 EMA20 就隔天開盤出場；"
                               f"目前 EMA20 {ind['E20'][-1]:,.0f}")
        if res["pending_exit"]:
            y = res["pending_exit"]
            d["pending_exit"] = f"{why_text(y) or '出場條件成立'}：下一根 K 開盤出場"
        return d

    def sig_info(res):
        x = res["sig"]
        if not x:
            return None
        return {"side": "多" if x["side"] > 0 else "空", "kind": x["kind"], "entry": round(x["entry"]),
                "stop": round(x["stop0"]), "R": round(x["R"])}

    def rec(res):
        tr = [trade_row(x, last_c) for x in res["trades"]]
        fwd = [r for r in tr if r["phase"] == "前測"]
        bt = [r for r in tr if r["phase"] == "回測"]
        def summ(rs):
            closed = [r for r in rs if not r["open"]]
            return {"n": len(closed), "win": sum(r["r"] > 0.1 for r in closed), "R": round(sum(r["r"] for r in closed), 1),
                    "twd": sum(r["twd"] for r in rs)}
        return {"forward": fwd, "backtest": bt, "fsum": summ(fwd), "bsum": summ(bt)}

    data = {
        "generated": f"{now:%Y-%m-%d %H:%M}",
        "last_price": last_c, "history_last": f"{hist_last:%Y-%m-%d %H:%M}", "live_minutes_added": added, "live": info,
        "v3": {"bar": f"{T[-1]:%Y-%m-%d %H:%M}", "bar_end": f"{v3_end:%H:%M}", "forming": v3_forming,
               "pos": pos_info(hv, "v3"), "sig": sig_info(hv), "checks": v3c, "record": rec(hv),
               "levels": {k: round(hv["ind"][k][-1]) for k in ("E20", "E60", "ATR")} |
                         {"BBU": round(hv["ind"]["MID"][-1] + 2 * hv["ind"]["SD"][-1]),
                          "BBL": round(hv["ind"]["MID"][-1] - 2 * hv["ind"]["SD"][-1])}},
        "d2": {"bar": f"{dlast:%Y-%m-%d}", "forming": d_forming,
               "pos": pos_info(hd, "d2"), "sig": sig_info(hd), "checks": d2c, "record": rec(hd),
               "levels": {k: round(hd["ind"][k][-1]) for k in ("E20", "E60", "ATR")},
               "bar_ohlc": [round(sp[k][-1]) for k in range(1, 5)]},
    }
    os.makedirs(a.out, exist_ok=True)
    json.dump(data, open(os.path.join(a.out, "live.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    open(os.path.join(a.out, "live.html"), "w", encoding="utf-8").write(render(data, now))
    tw60 = load(a.twii)
    tws = trades_page.twii_backtests(tw60, tw, refine_times)
    tx = {"v3": [trade_row(x, last_c) for x in hv["trades"]], "d2": [trade_row(x, last_c) for x in hd["trades"]]}
    open(os.path.join(a.out, "trades.html"), "w", encoding="utf-8").write(
        trades_page.render(tx, tws, now, last_c, f"{T[0]:%Y-%m-%d}", (CAP, MULT, FEE, COST), trade_row))
    print(f"現價 {last_c:,.0f}；history 到 {hist_last}，補上即時 {added} 分鐘；"
          f"v3 {'持倉' if hv['pos'] else '空手'}{'、訊號' if hv['sig'] else ''}；D2 {'持倉' if hd['pos'] else '空手'}{'、訊號' if hd['sig'] else ''}")


# ---------- 網頁 ----------
def session_state(now):
    wd, t = now.weekday(), now.time()
    if wd < 5 and dt.time(8, 45) <= t < dt.time(13, 45):
        return "日盤交易中"
    if (wd < 5 and t >= dt.time(15, 0)) or (1 <= wd <= 5 and t < dt.time(5, 0)):
        return "夜盤交易中"
    return "休市"


def money(v):
    cls = "up" if v > 0 else "dn" if v < 0 else ""
    return f'<span class="{cls}">{v:+,.0f}</span>'


def esc(s):
    return html.escape(str(s))


def checks_html(chk):
    def rows(items):
        return "".join(f'<li class="{"ok" if c["ok"] else "no"}"><span class="mk">{"✓" if c["ok"] else "✗"}</span>'
                       f'<span>{esc(c["name"])}</span><span class="dt">{esc(c["detail"])}</span></li>' for c in items)
    out = []
    for side in ("多", "空"):
        c = chk[side]
        groups = [g for g in ("突破", "拉回") if g in c]
        all_ok = {g: all(x["ok"] for x in c["common"] + c[g]) for g in groups}
        badge = "、".join(g for g in groups if all_ok[g])
        head = f'<span class="pill {"go" if badge else ""}">{esc(badge + "條件全到") if badge else "未成立"}</span>'
        body = f'<ul class="ck">{rows(c["common"])}</ul>'
        for g in groups:
            title = f"{g}進場" if len(groups) > 1 else "進場條件"
            body += f'<p class="gt">{title}</p><ul class="ck">{rows(c[g])}</ul>'
        st = c["stop"]
        body += f'<p class="note">若現在進場，停損約 {st["stop"]:,.0f}（{abs(st["raw"]):,.0f} 點前的端點，最多 1–1.5 ATR）</p>'
        out.append(f'<details class="side"{" open" if badge else ""}><summary>做{side} {head}</summary>{body}</details>')
    return "".join(out)


def strat_card(name, sub, s, kind, last_c, now):
    pos, sig = s["pos"], s["sig"]
    if kind == "v3":
        bar_txt = (f"最後一根 60 分K：{s['bar'][5:]} 起{'，形成中（' + s['bar_end'] + ' 收）' if s['forming'] else '，已收完'}")
    else:
        bar_txt = f"最後一根日K（只算日盤）：{s['bar'][5:]}" + ("，形成中（13:45 收）" if s["forming"] else "，已收盤")
    if pos:
        st_cls = "long" if pos["side"] == "多" else "short"
        status = f'<span class="state {st_cls}">持{pos["side"]}單</span>'
        p = pos
        be_txt = "已保本" if p["be"] else f"收盤到 {p['be_at']:,} 就把停損移到進場價（保本）"
        pend = f'<p class="alert">⚠ {esc(p["pending_exit"])}{"（以現價計，收K時再確認）" if s["forming"] else ""}</p>' if p.get("pending_exit") else ""
        body = f"""{pend}<div class="kv">
<div><span>進場</span><b>{p['entry']:,}</b><em>{p['entry_t'][5:]}・{esc(p['kind'])}</em></div>
<div><span>目前停損</span><b>{p['stop']:,}</b><em>距現價 {p['to_stop']:,} 點</em></div>
<div><span>未實現</span><b>{money(p['pts'])} 點</b><em>{money(p['twd'])} 元（1 口微台）</em></div>
<div><span>初始停損</span><b>{p['stop0']:,}</b><em>風險 {p['R']:,} 點</em></div>
</div><p class="note">{esc(be_txt)}。{esc(p['trail_note'])}。</p>"""
    else:
        status = '<span class="state flat">空手</span>'
        body = ""
    if sig:
        if kind == "v3":
            when = "下一根 60 分K 開盤進場"
        else:
            when = "13:45 收盤前進場" if s["forming"] else "應已在 13:45 收盤前進場"
        tag = "形成中：若以現價收K" if s["forming"] else "訊號成立"
        body = (f'<p class="alert sig">● {tag}，做{sig["side"]}（{esc(sig["kind"])}）→ {when}，'
                f'參考價 {sig["entry"]:,}、停損 {sig["stop"]:,}（{sig["R"]:,} 點）</p>') + body
    lv = s["levels"]
    if kind == "v3":
        lv_txt = f"EMA20 {lv['E20']:,}・EMA60 {lv['E60']:,}・布林 {lv['BBL']:,}–{lv['BBU']:,}・ATR {lv['ATR']:,}"
    else:
        o, h, l, c = s["bar_ohlc"]
        lv_txt = f"日EMA20 {lv['E20']:,}・EMA60 {lv['E60']:,}・日ATR {lv['ATR']:,}・最後日K 開{o:,} 高{h:,} 低{l:,} 收{c:,}"
    chk_title = "進場條件（看最後一根 K" + ("，以現價計）" if s["forming"] else "）")
    if pos:
        chk = f'<details><summary class="sub">進場條件明細（持倉中，不會再開新倉）</summary>{checks_html(s["checks"])}</details>'
    else:
        chk = f'<h3>{chk_title}</h3>{checks_html(s["checks"])}'
    return f"""<section class="card">
<div class="ch"><div><h2>{esc(name)}</h2><p class="sub">{esc(sub)}</p></div>{status}</div>
<p class="sub">{esc(bar_txt)}</p>
{body}
<p class="lv">{esc(lv_txt)}</p>
{chk}
</section>"""


def record_html(title, r):
    def table(rows):
        if not rows:
            return '<p class="note">還沒有交易。</p>'
        head = "".join(f"<th>{h}</th>" for h in ("進場", "方向", "進場價", "出場", "出場價", "R", "元"))
        body = "".join(
            "<tr>" + "".join(f"<td>{c}</td>" for c in (
                esc(x["entry_t"][5:]), f'{x["side"]}・{esc(x["kind"])}', f'{x["entry"]:,}',
                "持倉中" if x["open"] else esc(x["exit_t"][5:]), "" if x["open"] else f'{x["exit"]:,}',
                f'{x["r"]:+.2f}', money(x["twd"]))) + "</tr>" for x in reversed(rows))
        return f'<div class="scroll"><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'
    fs, bs = r["fsum"], r["bsum"]
    return f"""<section class="card"><h2>{esc(title)}</h2>
<p class="sub">前測（{FORWARD_START:%Y/%m/%d} 起）已出場 {fs['n']} 筆、賺 {fs['win']} 筆、合計 {fs['R']:+.1f}R、{money(fs['twd'])} 元（含未實現）</p>
{table(r['forward'])}
<p class="sub">前測開始前（台指期累積資料回測）：{bs['n']} 筆 {bs['R']:+.1f}R、{money(bs['twd'])} 元・<a href="trades.html">完整歷史成交紀錄 →</a></p>
</section>"""


def render(d, now):
    sess = session_state(now)
    live = d["live"]
    src = "・".join(f"{k} {v['last_minute'][5:]}" for k, v in live.items() if "last_minute" in v) or "沒有即時資料"
    errs = "".join(f'<p class="alert">⚠ {esc(k)}：{esc(v["error"])}</p>' for k, v in live.items() if "error" in v)
    dst = us_dst(now)
    night_win = "21:00–23:00" if dst else "22:00–00:00"
    return f"""<!doctype html>
<html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="600">
<title>台指策略即時訊號</title>
<style>
:root {{ --bg:#f6f7f9; --panel:#fff; --fg:#1b1f24; --muted:#5d6673; --line:#dfe3e8; --up:#c2412d; --dn:#2f7d5b;
  --accent:#2f5fd0; --warn:#a86500; --warnbg:#fff4e0; --sigbg:#e8efff;
  --font:"Noto Sans TC","PingFang TC","Microsoft JhengHei",system-ui,sans-serif; }}
@media (prefers-color-scheme: dark) {{ :root {{ --bg:#14171b; --panel:#1c2026; --fg:#e8ebef; --muted:#9aa3ae; --line:#2c323a;
  --up:#ef7a63; --dn:#5cc296; --accent:#7ea2ff; --warn:#f0b35a; --warnbg:#3a2c14; --sigbg:#1e2a48; color-scheme:dark; }} }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:var(--fg); font-family:var(--font); font-size:15px; line-height:1.5; }}
main {{ max-width:720px; margin:0 auto; padding:16px; }}
h1 {{ font-size:20px; margin:0; }} h2 {{ font-size:17px; margin:0; }} h3 {{ font-size:13px; color:var(--muted); font-weight:500; margin:14px 0 4px; }}
.sub, .note, .asof {{ color:var(--muted); font-size:12.5px; margin:4px 0; }}
.top {{ display:flex; justify-content:space-between; align-items:flex-end; gap:8px; flex-wrap:wrap; }}
.price {{ font-size:28px; font-variant-numeric:tabular-nums; margin:0; }}
.card {{ background:var(--panel); border:1px solid var(--line); border-radius:10px; padding:14px; margin-top:12px; }}
.ch {{ display:flex; justify-content:space-between; align-items:flex-start; gap:8px; }}
.state {{ font-weight:600; padding:4px 10px; border-radius:999px; white-space:nowrap; font-size:14px; border:1px solid var(--line); }}
.state.long {{ color:#fff; background:var(--up); border-color:var(--up); }}
.state.short {{ color:#fff; background:var(--dn); border-color:var(--dn); }}
.state.flat {{ color:var(--muted); }}
.kv {{ display:grid; grid-template-columns:1fr 1fr; gap:8px; margin:10px 0 4px; }}
.kv div {{ border:1px solid var(--line); border-radius:8px; padding:8px; }}
.kv > div > span {{ display:block; font-size:12px; color:var(--muted); }}
.kv b {{ display:block; font-size:19px; font-weight:600; font-variant-numeric:tabular-nums; }}
.kv em {{ display:block; font-style:normal; font-size:12px; color:var(--muted); }}
.alert {{ background:var(--warnbg); color:var(--warn); border-radius:8px; padding:8px 10px; margin:8px 0; font-size:14px; }}
.alert.sig {{ background:var(--sigbg); color:var(--accent); font-weight:600; }}
.lv {{ font-size:12.5px; color:var(--muted); border-top:1px dashed var(--line); padding-top:8px; margin:10px 0 0; font-variant-numeric:tabular-nums; }}
details.side {{ border:1px solid var(--line); border-radius:8px; margin:6px 0; padding:0 10px; }}
details.side summary {{ padding:8px 0; cursor:pointer; font-weight:600; }}
.pill {{ font-size:12px; font-weight:500; color:var(--muted); border:1px solid var(--line); border-radius:999px; padding:1px 8px; margin-left:6px; }}
.pill.go {{ color:#fff; background:var(--accent); border-color:var(--accent); }}
.gt {{ font-size:12.5px; color:var(--muted); margin:8px 0 2px; }}
ul.ck {{ list-style:none; padding:0; margin:0 0 6px; }}
ul.ck li {{ display:grid; grid-template-columns:20px 1fr auto; gap:6px; padding:3px 0; font-size:14px; }}
ul.ck li.ok .mk {{ color:var(--accent); font-weight:700; }} ul.ck li.no {{ color:var(--muted); }}
.dt {{ font-size:12px; color:var(--muted); text-align:right; font-variant-numeric:tabular-nums; }}
.scroll {{ overflow-x:auto; margin-top:8px; }}
table {{ border-collapse:collapse; width:100%; font-variant-numeric:tabular-nums; font-size:13px; }}
th, td {{ padding:5px 6px; text-align:right; border-bottom:1px solid var(--line); white-space:nowrap; }}
th {{ color:var(--muted); font-weight:500; }} th:nth-child(-n+2), td:nth-child(-n+2) {{ text-align:left; }}
.up {{ color:var(--up); }} .dn {{ color:var(--dn); }}
a {{ color:var(--accent); }}
footer {{ margin:20px 0 40px; }}
</style></head><body><main>
<div class="top"><div><h1>台指策略即時訊號</h1>
<p class="asof">{esc(sess)}・網頁產生於台北 {esc(d['generated'])}</p></div>
<div style="text-align:right"><p class="price">{d['last_price']:,.0f}</p><p class="asof">台指期近月・資料到 {esc(src)}</p></div></div>
{errs}
{strat_card("v3・60 分K", f"60 分K 收完才看訊號；日盤 08:45–11:45、夜盤 {night_win} 開始的 K 才開新倉", d["v3"], "v3", d["last_price"], now)}
{strat_card("D2・日K", "收盤創 20 日新高/新低，13:45 收盤前進場；日K 只看日盤，夜盤不算", d["d2"], "d2", d["last_price"], now)}
{record_html("v3 模擬交易紀錄", d["v3"]["record"])}
{record_html("D2 模擬交易紀錄", d["d2"]["record"])}
<footer>
<p class="note">GitHub Actions 在交易時段大約每 30 分鐘更新一次：日盤台北 08:48–13:48（:18、:48，另加 13:35 給 D2 收盤前看），
夜盤 15:04–05:04（:04、:34）。GitHub 的排程是盡力而為，常晚 5–30 分鐘，偶爾會漏跑，所以請看上面的「資料到」時間，不要拿這頁當下單依據。</p>
<p class="note">「形成中」表示那根 K 還沒收完，條件是用現價先算，收K前都可能變。v3 的訊號要等 60 分K 收完、下一根開盤才進場；D2 是日K 收盤前進場。
試算：本金 20 萬、1 口微台（每點 10 元）、每次下單 25 元；R 是以初始停損點數為 1 的倍數。</p>
<p class="note">規則：<code>strategy/v3.py</code> 的 V3、<code>strategy/daily.py</code> 的 D2。資料：期交所每日成交（history 到 {esc(d['history_last'])}）＋期交所即時 1 分K。
<a href="trades.html">歷史成交紀錄</a>・<a href="strategy.html">v3 前測（每日）</a>・<a href="index.html">選擇權盤面</a></p>
</footer>
</main></body></html>
"""


if __name__ == "__main__":
    main()
