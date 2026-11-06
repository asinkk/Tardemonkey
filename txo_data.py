#!/usr/bin/env python3
"""Build TXO weekly-option dashboard datasets from TAIFEX public data.

Usage: python3 -I build_txo_dashboard.py OUT_DIR
Writes summary.json, strikes.json, pain.json, pc_history.json into OUT_DIR.
Sources:
  - https://www.taifex.com.tw/cht/3/optDataDown  (daily per-contract CSV: OI, settlement)
  - https://www.taifex.com.tw/cht/3/pcRatioDown  (market P/C ratio history)
  - https://mis.taifex.com.tw/futures/api/getQuoteListOption  (live quotes, night or day session)
"""
import csv, io, json, math, os, sys, datetime as dt, urllib.request, urllib.parse

TW = dt.timezone(dt.timedelta(hours=8))
UA = {"User-Agent": "Mozilla/5.0 (txo-dashboard)"}


def post(url, data, js=False):
    if js:
        body = json.dumps(data).encode(); hdr = {**UA, "Content-Type": "application/json"}
    else:
        body = urllib.parse.urlencode(data).encode(); hdr = {**UA, "Content-Type": "application/x-www-form-urlencoded"}
    req = urllib.request.Request(url, data=body, headers=hdr, method="POST")
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def num(s):
    try:
        s = str(s).strip()
        return float(s) if s not in ("", "-") else None
    except ValueError:
        return None


def N(x):
    return 0.5 * (1 + math.erf(x / math.sqrt(2)))


def bs(F, K, T, s, cp):
    d1 = (math.log(F / K) + 0.5 * s * s * T) / (s * math.sqrt(T)); d2 = d1 - s * math.sqrt(T)
    return F * N(d1) - K * N(d2) if cp == "C" else K * N(-d2) - F * N(-d1)


def implied(px, F, K, T, cp):
    intr = max(F - K, 0) if cp == "C" else max(K - F, 0)
    if px is None or T <= 0 or px <= intr + 0.05:
        return None
    lo, hi = 0.01, 5.0
    for _ in range(100):
        m = (lo + hi) / 2
        if bs(F, K, T, m, cp) > px: hi = m
        else: lo = m
    return m if 0.011 < m < 4.99 else None


def main(out):
    now = dt.datetime.now(TW)
    start = (now - dt.timedelta(days=10)).strftime("%Y/%m/%d"); end = now.strftime("%Y/%m/%d")
    raw = post("https://www.taifex.com.tw/cht/3/optDataDown",
               {"down_type": "1", "commodity_id": "TXO", "commodity_id2": "", "queryStartDate": start, "queryEndDate": end})
    rows = [r for r in csv.reader(io.StringIO(raw.decode("big5", "replace"))) if len(r) > 20 and r[1] == "TXO"]
    trade_date = max(r[0] for r in rows)
    day = [r for r in rows if r[0] == trade_date]
    prev_dates = sorted({r[0] for r in rows if r[0] < trade_date})
    prev_date = prev_dates[-1] if prev_dates else None
    prev_oi = {}
    for r in rows:
        if r[0] == prev_date and r[17] == "一般":
            prev_oi[(r[2].strip(), int(float(r[3])), "C" if r[4] == "買權" else "P")] = int(r[11] or 0)
    td = dt.datetime.strptime(trade_date, "%Y/%m/%d").date()

    # contracts still alive after trade date
    exps = {}
    for r in day:
        code = r[2].strip(); ed = dt.datetime.strptime(r[20].strip(), "%Y%m%d").date()
        if ed > td: exps[code] = ed
    weekly = sorted((v, k) for k, v in exps.items() if not k.isdigit())
    monthly = sorted((v, k) for k, v in exps.items() if k.isdigit())
    chosen = [k for _, k in weekly[:2]] + ([monthly[0][1]] if monthly else [])
    chosen = sorted(set(chosen), key=lambda k: exps[k])

    summary, strikes_out, pain_out = [], [], []
    for code in chosen:
        ed = exps[code]
        oi, vol, settle = {}, {}, {}
        for r in day:
            if r[2].strip() != code: continue
            k = (int(float(r[3])), "C" if r[4] == "買權" else "P")
            vol[k] = vol.get(k, 0) + int(r[9] or 0)
            if r[17] == "一般":
                oi[k] = int(r[11] or 0); settle[k] = num(r[10])
        # live quotes: try night (1) then day (0)
        live, quote_time, session = {}, None, None
        hm = now.hour * 100 + now.minute
        day_first = 845 <= hm < 1400  # 日盤時段先用日盤報價，其餘先用夜盤
        order = (("0", "日盤"), ("1", "夜盤")) if day_first else (("1", "夜盤"), ("0", "日盤"))
        for mt, name in order:
            try:
                d = json.loads(post("https://mis.taifex.com.tw/futures/api/getQuoteListOption",
                                    {"MarketType": mt, "SymbolType": "O", "KindID": "1", "CID": "TXO", "ExpireMonth": code,
                                     "RowSize": "全部", "PageNo": "", "SortColumn": "", "AscDesc": "A"}, js=True))
                q = d.get("RtData", {}).get("QuoteList", []) or []
            except Exception:
                q = []
            good = 0; tmp = {}; times = []
            for x in q:
                k = (int(float(x["StrikePrice"])), x["CP"])
                b, a = num(x["CBidPrice1"]), num(x["CAskPrice1"])
                mid = (b + a) / 2 if b and a and b > 0 and a > 0 and a >= b else None
                if mid: good += 1
                tmp[k] = {"mid": mid, "vol": int(x.get("CTotalVolume") or 0)}
                if x.get("CTime") and x.get("CDate") == now.strftime("%Y%m%d"): times.append(x["CDate"] + x["CTime"])
            if good >= 20:
                live, session = tmp, name
                quote_time = max(times) if times else None
                break
        price = {}
        for k in set(oi) | set(live):
            p = live.get(k, {}).get("mid")
            price[k] = p if p else settle.get(k)
        if not session: session = "結算價"
        qt = dt.datetime.strptime(quote_time, "%Y%m%d%H%M%S").replace(tzinfo=TW) if quote_time else now
        Ks = sorted({k[0] for k in price})
        best = None
        for K in Ks:
            c, p = price.get((K, "C")), price.get((K, "P"))
            if c and p and (best is None or abs(c - p) < best[0]): best = (abs(c - p), K, c, p)
        if not best: continue
        _, K0, c0, p0 = best
        F = K0 + c0 - p0
        expiry_dt = dt.datetime.combine(ed, dt.time(13, 30), TW)
        T = max((expiry_dt - qt).total_seconds(), 3600) / (365 * 86400)
        atm_iv = implied(c0, F, K0, T, "C") or implied(p0, F, K0, T, "P")
        straddle = c0 + p0

        tc = sum(v for k, v in oi.items() if k[1] == "C"); tp = sum(v for k, v in oi.items() if k[1] == "P")
        lv_c = sum(v["vol"] for k, v in live.items() if k[1] == "C"); lv_p = sum(v["vol"] for k, v in live.items() if k[1] == "P")
        dv_c = sum(v for k, v in vol.items() if k[1] == "C"); dv_p = sum(v for k, v in vol.items() if k[1] == "P")
        near = [k for k in oi if abs(k[0] - F) <= 2500]
        nc = sum(oi[k] for k in near if k[1] == "C"); np_ = sum(oi[k] for k in near if k[1] == "P")
        allK = sorted({k[0] for k in oi})
        def payout(S):
            return sum(v * max(S - k[0], 0) if k[1] == "C" else v * max(k[0] - S, 0) for k, v in oi.items())
        max_pain = min(allK, key=payout) if allK else None
        label = ("週選 " + code[-2:] if not code.isdigit() else "月選") + f"（{ed.month}/{ed.day}）"
        call_wall = max(((v, k[0]) for k, v in oi.items() if k[1] == "C" and abs(k[0] - F) <= 3000), default=(0, None))
        put_wall = max(((v, k[0]) for k, v in oi.items() if k[1] == "P" and abs(k[0] - F) <= 3000), default=(0, None))
        summary.append({
            "expiry": code, "label": label, "expiry_date": ed.isoformat(), "days_left": round(T * 365, 2),
            "trade_date": trade_date.replace("/", "-"), "quote_time": qt.strftime("%Y-%m-%d %H:%M"), "price_basis": session,
            "forward": round(F, 1), "atm_strike": K0, "atm_call": c0, "atm_put": p0, "straddle": round(straddle, 1),
            "move_pct": round(straddle / F, 4), "range_low": round(F - straddle), "range_high": round(F + straddle),
            "atm_iv": round(atm_iv, 4) if atm_iv else None,
            "sigma_pts": round(F * atm_iv * math.sqrt(T)) if atm_iv else None,
            "call_oi": tc, "put_oi": tp, "pc_oi": round(tp / tc, 3) if tc else None,
            "pc_oi_near": round(np_ / nc, 3) if nc else None,
            "pc_vol_day": round(dv_p / dv_c, 3) if dv_c else None,
            "pc_vol_live": round(lv_p / lv_c, 3) if lv_c else None,
            "max_pain": max_pain, "call_wall": call_wall[1], "call_wall_oi": call_wall[0],
            "put_wall": put_wall[1], "put_wall_oi": put_wall[0],
            "oi_prev_date": prev_date.replace("/", "-") if prev_date else None,
        })
        for K in Ks:
            if abs(K - F) > 3000: continue
            cm, pm = price.get((K, "C")), price.get((K, "P"))
            civ, piv = implied(cm, F, K, T, "C"), implied(pm, F, K, T, "P")
            otm = civ if K >= F else piv
            strikes_out.append({
                "expiry": code, "strike": K, "call_oi": oi.get((K, "C"), 0), "put_oi": oi.get((K, "P"), 0),
                "call_price": cm, "put_price": pm,
                "call_iv": round(civ, 4) if civ else None, "put_iv": round(piv, 4) if piv else None,
                "otm_iv": round(otm, 4) if otm else None,
                "call_vol": vol.get((K, "C"), 0), "put_vol": vol.get((K, "P"), 0),
                "call_oi_chg": oi.get((K, "C"), 0) - prev_oi.get((code, K, "C"), 0) if prev_oi else None,
                "put_oi_chg": oi.get((K, "P"), 0) - prev_oi.get((code, K, "P"), 0) if prev_oi else None,
            })
        for S in allK:
            if abs(S - F) <= 3000:
                pain_out.append({"expiry": code, "settle": S, "payout": round(payout(S) * 50 / 1e6, 2)})  # 百萬元 (每點50元)

    pcs = (now - dt.timedelta(days=30)).strftime("%Y/%m/%d")
    pr = post("https://www.taifex.com.tw/cht/3/pcRatioDown", {"queryStartDate": pcs, "queryEndDate": end}).decode("big5", "replace")
    pc_hist = []
    for r in list(csv.reader(io.StringIO(pr)))[1:]:
        if len(r) >= 7 and r[0][:2] == "20":
            pc_hist.append({"date": r[0].replace("/", "-"), "pc_vol": round(float(r[3]) / 100, 4), "pc_oi": round(float(r[6]) / 100, 4),
                            "put_oi": int(r[4]), "call_oi": int(r[5])})
    pc_hist.sort(key=lambda x: x["date"])

    os.makedirs(out, exist_ok=True)
    for name, data in (("summary", summary), ("strikes", strikes_out), ("pain", pain_out), ("pc_history", pc_hist)):
        with open(os.path.join(out, name + ".json"), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)


if __name__ == "__main__":
    main(sys.argv[1])
