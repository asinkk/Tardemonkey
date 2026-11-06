"""盤面延伸指標：三大法人與大額交易人選擇權部位、台指期籌碼與正逆價差、VIX/HV、美股與 VIX。

extra(store_dir, atm_iv) 回傳 dict；任一區塊抓不到就留空字典並記在 errors，不影響主頁面。
資料來源：期交所（taifex.com.tw / mis.taifex.com.tw）與 Yahoo Finance 公開圖表 API。
"""
import csv, io, json, math, os, datetime as dt, urllib.request, urllib.parse

TW = dt.timezone(dt.timedelta(hours=8))
UA = {"User-Agent": "Mozilla/5.0 (txo-dashboard)"}


def _req(url, data=None, js=False, timeout=40):
    hdr = dict(UA)
    body = None
    if data is not None:
        if js:
            body = json.dumps(data).encode(); hdr["Content-Type"] = "application/json"
        else:
            body = urllib.parse.urlencode(data).encode(); hdr["Content-Type"] = "application/x-www-form-urlencoded"
    req = urllib.request.Request(url, data=body, headers=hdr, method="POST" if body else "GET")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def _csv(url, data):
    # 結束日若還沒有資料（例如凌晨夜盤），期交所會回錯誤頁 → 結束日往前退
    end = dt.datetime.strptime(data["queryEndDate"], "%Y/%m/%d")
    for back in range(6):
        d = dict(data, queryEndDate=(end - dt.timedelta(days=back)).strftime("%Y/%m/%d"))
        txt = _req(url, d).decode("big5", "replace")
        rows = [r for r in csv.reader(io.StringIO(txt)) if r and r[0][:2] == "20"]
        if rows: return rows
    return []


def _i(x):
    try: return int(float(str(x).strip()))
    except ValueError: return 0


def _d(s):  # 2026/11/06 -> 2026-11-06
    return s.strip().replace("/", "-")


# ---------- 1. 三大法人選擇權（買權/賣權分計） ----------
def inst_options(now):
    rows = _csv("https://www.taifex.com.tw/cht/3/callsAndPutsDateDown",
                {"queryStartDate": (now - dt.timedelta(days=30)).strftime("%Y/%m/%d"),
                 "queryEndDate": now.strftime("%Y/%m/%d"), "commodityId": "TXO"})
    who_map = {"外資及陸資": "foreign", "外資": "foreign", "自營商": "dealer", "投信": "trust"}
    by = {}
    for r in rows:
        if len(r) < 16: continue
        who = who_map.get(r[3].strip())
        if not who: continue
        rec = by.setdefault(_d(r[0]), {"date": _d(r[0])})
        cp = "call" if r[2].strip().upper() == "CALL" else "put"
        rec[f"{who}_{cp}_oi"] = _i(r[14])            # 未平倉口數淨額
        rec[f"{who}_{cp}_amt"] = round(_i(r[15]) / 1e5, 2)  # 未平倉金額淨額（億元）
        rec[f"{who}_{cp}_trade"] = _i(r[8])          # 當日交易口數淨額
    hist = [by[k] for k in sorted(by)]
    for h in hist:  # 外資方向指標：買權淨金額 − 賣權淨金額（億元），正值偏多
        if "foreign_call_amt" in h and "foreign_put_amt" in h:
            h["foreign_bias_amt"] = round(h["foreign_call_amt"] - h["foreign_put_amt"], 2)
    return hist


# ---------- 2. 大額交易人（前五/前十大、特定法人），全部月份合計 ----------
def large_traders(now):
    rows = _csv("https://www.taifex.com.tw/cht/3/largeTraderOptDown",
                {"queryStartDate": (now - dt.timedelta(days=10)).strftime("%Y/%m/%d"),
                 "queryEndDate": now.strftime("%Y/%m/%d")})
    by = {}
    for r in rows:
        if len(r) < 11 or r[1].strip() != "TXO" or r[4].strip() != "999999": continue
        cp = "call" if r[3].strip() == "買權" else "put"
        kind = "all" if r[5].strip() == "0" else "inst"   # 0 全部交易人、1 特定法人
        rec = by.setdefault(_d(r[0]), {"date": _d(r[0])})
        rec[f"{kind}_{cp}_top5_net"] = _i(r[6]) - _i(r[7])
        rec[f"{kind}_{cp}_top10_net"] = _i(r[8]) - _i(r[9])
        rec[f"{kind}_{cp}_top10_buy"] = _i(r[8]); rec[f"{kind}_{cp}_top10_sell"] = _i(r[9])
        rec[f"{cp}_market_oi"] = _i(r[10])
    return [by[k] for k in sorted(by)]


# ---------- 3. 台指期：三大法人淨部位 + 即時價格與正逆價差 ----------
def futures(now):
    rows = _csv("https://www.taifex.com.tw/cht/3/futContractsDateDown",
                {"queryStartDate": (now - dt.timedelta(days=30)).strftime("%Y/%m/%d"),
                 "queryEndDate": now.strftime("%Y/%m/%d"), "commodityId": "TXF"})
    who_map = {"外資及陸資": "foreign", "外資": "foreign", "自營商": "dealer", "投信": "trust"}
    by = {}
    for r in rows:
        if len(r) < 15: continue
        who = who_map.get(r[2].strip())
        if not who: continue
        rec = by.setdefault(_d(r[0]), {"date": _d(r[0])})
        rec[f"{who}_net_oi"] = _i(r[13]); rec[f"{who}_net_trade"] = _i(r[7])
    hist = [by[k] for k in sorted(by)]

    live = {}
    for mt, name in (("0", "日盤"), ("1", "夜盤")):
        try:
            d = json.loads(_req("https://mis.taifex.com.tw/futures/api/getQuoteList",
                                {"MarketType": mt, "SymbolType": "F", "KindID": "1", "CID": "TXF", "ExpireMonth": "",
                                 "RowSize": "全部", "PageNo": "", "SortColumn": "", "AscDesc": "A"}, js=True))
            q = d.get("RtData", {}).get("QuoteList", []) or []
        except Exception:
            q = []
        spot = next((x for x in q if "現貨" in x.get("DispCName", "")), None)
        futs = [x for x in q if x.get("SymbolID", "").startswith("TXF") and "現貨" not in x.get("DispCName", "")]
        if not futs: continue
        f = futs[0]
        last = float(f["CLastPrice"]) if f.get("CLastPrice") else None
        ref = float(f["CRefPrice"]) if f.get("CRefPrice") else None
        sp = None
        if spot:
            sp = float(spot["CLastPrice"]) if spot.get("CLastPrice") else (float(spot["CRefPrice"]) if spot.get("CRefPrice") else None)
        live[name] = {
            "name": f.get("DispCName"), "last": last, "ref": ref,
            "chg": round(last - ref, 1) if last and ref else None,
            "time": (f.get("CDate", "") + " " + f.get("CTime", "")).strip(),
            "oi": _i(f.get("OpenInterest") or 0),
            "spot": sp,
            # 夜盤時現貨已收盤，價差無意義，只算日盤
            "basis": round(last - sp, 2) if last and sp and name == "日盤" else None,
        }
    return {"history": hist, "live": live}


# ---------- 4. 台指 VIX 歷史（期交所只留近 3 個月，存檔累積成一年）+ HV ----------
def _yahoo(sym, rng="5d", interval="1d"):
    last_err = None
    for host in ("query1", "query2"):
        try:
            d = json.loads(_req(f"https://{host}.finance.yahoo.com/v8/finance/chart/{urllib.parse.quote(sym)}?range={rng}&interval={interval}"))
            return d["chart"]["result"][0]
        except Exception as e:
            last_err = e
    raise last_err


def vol_regime(now, store_dir, atm_iv):
    path = os.path.join(store_dir, "vix_daily.csv") if store_dir else None
    vix = {}
    if path and os.path.exists(path):
        for r in csv.DictReader(open(path, encoding="utf-8")):
            vix[r["date"]] = float(r["vix"])
    m = now.replace(day=1)
    months = []
    for _ in range(3):
        months.append(m.strftime("%Y%m")); m = (m - dt.timedelta(days=1)).replace(day=1)
    for ym in months:
        try:
            txt = _req(f"https://www.taifex.com.tw/file/taifex/Dailydownload/vix/log2data/{ym}new.txt").decode("big5", "replace")
        except Exception:
            continue
        for line in txt.splitlines():
            p = line.split()
            if len(p) >= 3 and p[0].isdigit() and len(p[0]) == 8:
                try: vix[f"{p[0][:4]}-{p[0][4:6]}-{p[0][6:]}"] = float(p[2])
                except ValueError: pass
    cutoff = (now - dt.timedelta(days=366)).strftime("%Y-%m-%d")
    vix = {k: v for k, v in vix.items() if k >= cutoff}
    if path and vix:
        os.makedirs(store_dir, exist_ok=True)
        with open(path, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f); w.writerow(["date", "vix"])
            for k in sorted(vix): w.writerow([k, vix[k]])
    vh = [{"date": k, "vix": vix[k]} for k in sorted(vix)]

    # 加權指數日收盤 → 20 日歷史波動（年化）
    hv = []
    try:
        r = _yahoo("^TWII", "2y", "1d")
        ts, cl = r["timestamp"], r["indicators"]["quote"][0]["close"]
        px = [(dt.datetime.fromtimestamp(t, TW).strftime("%Y-%m-%d"), c) for t, c in zip(ts, cl) if c]
        rets = [math.log(px[i][1] / px[i - 1][1]) for i in range(1, len(px))]
        for i in range(20, len(rets) + 1):
            w = rets[i - 20:i]; mu = sum(w) / 20
            sd = math.sqrt(sum((x - mu) ** 2 for x in w) / 19)
            hv.append({"date": px[i][0], "hv20": round(sd * math.sqrt(252), 4), "close": round(px[i][1], 2)})
        hv = [h for h in hv if h["date"] >= cutoff]
    except Exception:
        hv = []

    def rank(series, cur):
        if not series or cur is None: return None, None
        lo, hi = min(series), max(series)
        r = (cur - lo) / (hi - lo) if hi > lo else None
        pct = sum(1 for x in series if x <= cur) / len(series)
        return (round(r, 3) if r is not None else None), round(pct, 3)

    vs = [x["vix"] for x in vh]
    cur_vix = vs[-1] if vs else None
    vix_rank, vix_pct = rank(vs, cur_vix)
    hs = [x["hv20"] for x in hv]
    cur_hv = hs[-1] if hs else None
    hv_rank, hv_pct = rank(hs, cur_hv)
    return {
        "vix_history": vh, "hv_history": hv,
        "vix": cur_vix, "vix_date": vh[-1]["date"] if vh else None,
        "vix_min": min(vs) if vs else None, "vix_max": max(vs) if vs else None,
        "vix_rank": vix_rank, "vix_pct": vix_pct,
        "vix_days": len(vs), "vix_from": vh[0]["date"] if vh else None,
        "hv20": cur_hv, "hv_rank": hv_rank, "hv_pct": hv_pct,
        "atm_iv": atm_iv,
        "iv_hv_ratio": round(atm_iv / cur_hv, 2) if atm_iv and cur_hv else None,
        "vix_hv_ratio": round(cur_vix / 100 / cur_hv, 2) if cur_vix and cur_hv else None,
    }


# ---------- 5. 美股期貨與 VIX ----------
GLOBAL = [("^VIX", "美國 VIX"), ("ES=F", "S&P 500 期貨"), ("NQ=F", "那斯達克 100 期貨"),
          ("YM=F", "道瓊期貨"), ("^SOX", "費城半導體"), ("TSM", "台積電 ADR")]


def global_mkts():
    out = []
    for sym, name in GLOBAL:
        try:
            r = _yahoo(sym, "1mo", "1d")
            m = r["meta"]
            last = m.get("regularMarketPrice")
            closes = [c for c in r["indicators"]["quote"][0]["close"] if c]
            prev = m.get("chartPreviousClose") if len(closes) < 2 else None
            # 前一交易日收盤：以日線倒數第二根為準（最後一根是今天或最新一天）
            prev = closes[-2] if len(closes) >= 2 else prev
            t = dt.datetime.fromtimestamp(m.get("regularMarketTime", 0), TW).strftime("%m-%d %H:%M")
            out.append({"symbol": sym, "name": name, "last": round(last, 2) if last else None,
                        "prev": round(prev, 2) if prev else None,
                        "chg_pct": round(last / prev - 1, 4) if last and prev else None,
                        "time": t, "spark": [round(c, 2) for c in closes[-20:]]})
        except Exception:
            out.append({"symbol": sym, "name": name, "last": None})
    return out


def extra(store_dir=None, atm_iv=None):
    now = dt.datetime.now(TW)
    res, errors = {}, []
    for key, fn in (("inst_options", lambda: inst_options(now)), ("large_traders", lambda: large_traders(now)),
                    ("futures", lambda: futures(now)), ("vol", lambda: vol_regime(now, store_dir, atm_iv)),
                    ("global", global_mkts)):
        try:
            res[key] = fn()
        except Exception as e:
            res[key] = None; errors.append(f"{key}: {e}")
    res["errors"] = errors
    return res


if __name__ == "__main__":
    import sys
    print(json.dumps(extra(sys.argv[1] if len(sys.argv) > 1 else None, 0.45), ensure_ascii=False, indent=1)[:6000])
