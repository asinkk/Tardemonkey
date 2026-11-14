"""週選即時訊號：v3 60 分K（限時段）出訊號時，買同方向、最近到期、權利金約 30 點的台指週選，分批停利、不設停損。

規則（2026-11-14 用 10/1–11/13 逐筆成交回測後定案，樣本只有 9 筆，見 docs/options.html 的說明）：
  - 多單買買權（C）、空單買賣權（P）；到期日最近的一檔，距到期不到 22.6 小時（＝訊號所屬交易日當天到期）就改買下一個到期。
  - 履約價取當時價格最接近 30 點的那一檔。
  - 出場分四批：3 倍、5 倍、10 倍限價停利，以及「漲過 5 倍後，從最高價回落 30%」出場；沒成交的放到結算。不設停損。
    口數怎麼分在網頁上算（30%／30%／20%／20%），這裡只記每一批每口的結果。

資料：
  - history/txo_ticks/：期交所選擇權逐筆成交（每天 06:30 存前一個交易日的夜盤＋日盤）。有逐筆的部分一律以逐筆為準。
  - 期交所 MIS 即時報價與 1 分K：逐筆還沒存到的部分（當天的盤）先用這個，隔天逐筆進來後重算校正。
狀態存在 history/option_trades.json（每筆訊號一筆紀錄，以 v3 進場時間當 id）。
只用 Python 標準函式庫。由 live.py 呼叫 update()。
"""
import csv, datetime as dt, glob, json, lzma, os, re

TARGET = 30                        # 想買的權利金（點）
MIN_LEFT = dt.timedelta(hours=22.6)  # 距到期不到這麼久＝當天到期，改買下一期
MULTS = (3, 5, 10)                 # 限價停利倍數
TRAIL_ARM, TRAIL_BACK = 5, 0.30    # 漲過 5 倍後，從最高價回落 30% 出場
COST = 1.0                         # 每口來回成本（點）：手續費＋交易稅
POINT = 50                         # 1 點 50 元
START = dt.datetime(2026, 10, 1, 8, 45)        # 逐筆資料從這天起
FORWARD_START = dt.datetime(2026, 11, 16)      # 規則 11/14 定案，之後的訊號算前測
F = "%Y-%m-%d %H:%M:%S"
MONTHS_C, MONTHS_P = "ABCDEFGHIJKL", "MNOPQRSTUVWX"


def tick_size(p):
    return 0.1 if p < 10 else 0.5 if p < 50 else 1 if p < 500 else 5 if p < 1000 else 10


def round_tick(p):
    t = tick_size(p)
    return round(round(p / t) * t, 1)


def stamp(t):
    return t.strftime("%Y%m%d%H%M%S")


def ts(s):
    return dt.datetime.strptime(s, "%Y%m%d%H%M%S")


# ---------- 到期日 ----------
def nth_weekday(year, month, wd, n):
    first = dt.date(year, month, 1)
    return first + dt.timedelta(days=(wd - first.weekday()) % 7 + 7 * (n - 1))


def expiry_time(code, trading_days=()):
    """202611 → 第 3 個週三；202611W2 → 第 2 個週三；202611F3 → 第 3 個週五，13:30 到期。
    trading_days（已有逐筆檔的日期，排序）用來處理遇假日順延：名目到期日沒開盤就順延到下一個交易日。"""
    y, m = int(code[:4]), int(code[4:6])
    d = nth_weekday(y, m, 2, 3) if len(code) == 6 else nth_weekday(y, m, 2 if code[6] == "W" else 4, int(code[7:]))
    if trading_days and d not in trading_days and d < trading_days[-1]:
        d = next(x for x in trading_days if x > d)
    return dt.datetime.combine(d, dt.time(13, 30))


# ---------- 逐筆檔 ----------
class Ticks:
    """history/txo_ticks 的讀取。檔名是交易日：前一晚 15:00 起的夜盤＋當天日盤。"""

    def __init__(self, folder):
        self.files = {}
        for f in sorted(glob.glob(os.path.join(folder, "Daily_*.csv.xz"))):
            y, m, d = os.path.basename(f)[6:16].split("_")
            self.files[dt.date(int(y), int(m), int(d))] = f
        self.days = sorted(self.files)
        self._cache = {}

    def trading_day(self, t):
        """t 屬於哪個交易日的檔案（夜盤算下一個交易日）；還沒有那個檔案回傳 None。"""
        d = t.date() + dt.timedelta(days=1) if t.time() >= dt.time(14, 0) else t.date()
        return next((x for x in self.days if x >= d), None)

    def load(self, day):
        """回傳 ({(到期, 履約價, C/P): [(yyyymmddHHMMSS 字串, 價)]}, 這個檔最後一筆成交時間字串)。只留最近讀的 3 個檔。"""
        if day not in self._cache:
            by, last = {}, ""
            with lzma.open(self.files[day], "rt", encoding="utf-8", newline="") as f:
                r = csv.reader(f)
                next(r)
                for row in r:
                    ts = row[0] + row[1].zfill(6)
                    by.setdefault((row[2], int(float(row[3])), row[4]), []).append((ts, float(row[5])))
                    if ts > last:
                        last = ts
            for v in by.values():
                v.sort(key=lambda x: x[0])
            while len(self._cache) >= 3:
                self._cache.pop(next(iter(self._cache)))
            self._cache[day] = (by, last)
        return self._cache[day]

    def covered_to(self):
        """逐筆資料涵蓋到什麼時候（最後一個檔的最後一筆）。"""
        return ts(self.load(self.days[-1])[1]) if self.days else None

    def path(self, key, after, until):
        """某一檔在 (after, until] 的逐筆成交。"""
        out = []
        d0 = self.trading_day(after)
        for day in self.days:
            if d0 is None or day < d0 or day > until.date():
                continue
            a, u = stamp(after), stamp(until)
            out += [(ts(t), p) for t, p in self.load(day)[0].get(key, ()) if a < t <= u]
        return out

    def pick(self, t, side):
        """訊號時間 t 當下（逐筆）要買哪一檔：回傳 dict 或 None。"""
        day = self.trading_day(t)
        if day is None:
            return None
        by, last = self.load(day)
        if ts(last) < t:
            return None                                   # 這個檔還沒收到訊號之後的資料（例如只有週五夜盤的週一檔）
        cp = "C" if side > 0 else "P"
        exps = sorted({(expiry_time(k[0], self.days), k[0]) for k in by if k[2] == cp})
        exps = [(e, c) for e, c in exps if e - t >= MIN_LEFT]
        if not exps:
            return None
        exp_t, code = exps[0]
        ref, lo, hi = {}, stamp(t - dt.timedelta(minutes=3)), stamp(t + dt.timedelta(seconds=5))
        for (c, strike, p), rows in by.items():
            if c == code and p == cp:
                px = [x[1] for x in rows if lo <= x[0] < hi]
                if px:
                    ref[strike] = px[-1]
        if not ref:
            return None
        strike = min(ref, key=lambda s: (abs(ref[s] - TARGET), s))
        first = next(((ts(x), p) for x, p in by[(code, strike, cp)] if x >= stamp(t)), None)
        if first is None:
            return None
        ks = sorted(ref)
        i = ks.index(strike)
        alts = [{"strike": s, "px": ref[s]} for s in ks[max(0, i - 1):i + 2] if s != strike]
        return {"code": code, "strike": strike, "cp": cp, "exp_t": exp_t, "entry": first[1], "entry_at": first[0],
                "alts": alts, "src": "逐筆"}


# ---------- 期交所 MIS 即時 ----------
def parse_symbol(q, year_now):
    """MIS 報價 → (到期代碼, 履約價, C/P, 到期時間)；不是台指選擇權個別契約回傳 None。
    SymbolID 例：TXO47000K6-O（月選 11 月買權）、TXX44000W6-N（F3 11 月賣權，夜盤）。"""
    m = re.match(r"^TX.(\d+)([A-X])(\d)-[A-Z]$", q.get("SymbolID", ""))
    if not m:
        return None
    letter = m.group(2)
    cp = "C" if letter in MONTHS_C else "P"
    month = (MONTHS_C if cp == "C" else MONTHS_P).index(letter) + 1
    y = year_now // 10 * 10 + int(m.group(3))
    if y < year_now - 1:
        y += 10
    name = q.get("DispCName", "")
    w = re.search(r"([WF]\d)\s*\((\d{4})/(\d\d)/(\d\d)\)", name)
    if w:
        code = f"{y}{month:02d}{w.group(1)}"
        exp_t = dt.datetime(int(w.group(2)), int(w.group(3)), int(w.group(4)), 13, 30)
    else:
        code = f"{y}{month:02d}"
        exp_t = expiry_time(code)
    return code, int(m.group(1)), cp, exp_t


def is_day(now):
    return dt.time(8, 45) <= now.time() < dt.time(13, 46)


class Mis:
    def __init__(self, post, now):
        self.post, self.now, self._q = post, now, None

    def quotes(self):
        """目前盤別（日盤／夜盤）所有台指選擇權報價：{(到期, 履約價, C/P): dict}。"""
        if self._q is None:
            self._q = {}
            raw = self.post("getQuoteList", {"MarketType": "0" if is_day(self.now) else "1", "SymbolType": "O", "KindID": "1",
                                             "CID": "TXO", "ExpireMonth": "", "RowSize": "全部", "PageNo": "",
                                             "SortColumn": "", "AscDesc": "A"})["RtData"]["QuoteList"]
            for q in raw:
                k = parse_symbol(q, self.now.year)
                if not k:
                    continue
                f = lambda name: float(q.get(name) or 0)
                bid, ask, last = f("CBidPrice1"), f("CAskPrice1"), f("CLastPrice")
                px = (bid + ask) / 2 if bid > 0 and ask > 0 and ask <= max(bid * 1.5, bid + 3) else last
                self._q[k[:3]] = {"symbol": q["SymbolID"], "exp_t": k[3], "bid": bid, "ask": ask, "last": last, "px": px,
                                  "vol": f("CTotalVolume")}
        return self._q

    def pick(self, t, side):
        cp = "C" if side > 0 else "P"
        q = {k: v for k, v in self.quotes().items() if k[2] == cp and v["px"] > 0 and v["exp_t"] > self.now}
        exps = sorted({(v["exp_t"], k[0]) for k, v in q.items() if v["exp_t"] - t >= MIN_LEFT})
        if not exps:
            return None
        exp_t, code = exps[0]
        ref = {k[1]: v for k, v in q.items() if k[0] == code and (v["vol"] > 0 or v["bid"] > 0)}
        if not ref:
            return None
        strike = min(ref, key=lambda s: (abs(ref[s]["px"] - TARGET), s))
        ks = sorted(ref)
        i = ks.index(strike)
        alts = [{"strike": s, "px": round_tick(ref[s]["px"])} for s in ks[max(0, i - 1):i + 2] if s != strike]
        v = ref[strike]
        return {"code": code, "strike": strike, "cp": cp, "exp_t": exp_t, "entry": round_tick(v["px"]), "entry_at": self.now,
                "alts": alts, "src": "即時", "bid": v["bid"], "ask": v["ask"]}

    def bars(self, key):
        """某一檔目前日盤、夜盤的 1 分K：[(時間, 開, 高, 低, 收)]。抓不到回傳空的。"""
        q = self.quotes().get(tuple(key))
        if not q:
            return []
        out = {}
        for suffix, night in (("O", False), ("N", True)):
            try:
                d = self.post("getChartData1M", {"SymbolID": q["symbol"][:-1] + suffix})["RtData"]
                base = dt.datetime.strptime(d["Quote"]["CDate"], "%Y%m%d")
            except Exception:
                continue
            for t, o, h, l, c, _v in d.get("Ticks") or []:
                hh, mm = int(t[:2]), int(t[2:4])
                day = base + dt.timedelta(days=1) if night and hh < 15 else base
                end = day + dt.timedelta(minutes=hh * 60 + mm)      # MIS 標的是這分鐘的結束時間
                out[end] = (float(o), float(h), float(l), float(c))
        return [(t, *out[t]) for t in sorted(out)]


# ---------- 模擬 ----------
def simulate(entry, seq):
    """seq：進場後依時間排序的 (時間, 價)。回傳各批每口的結果。"""
    tg = {m: round_tick(entry * m) for m in MULTS}
    arm = round_tick(entry * TRAIL_ARM)
    fills, trail, peak, peak_t, last, last_t = {}, None, entry, None, entry, None
    for t, p in seq:
        for m in MULTS:
            if m not in fills and p >= tg[m]:
                fills[m] = t
        if p > peak:
            peak, peak_t = p, t
        if trail is None and peak >= arm and p <= peak * (1 - TRAIL_BACK):
            trail = (t, p)
        last, last_t = p, t
    return {"targets": {str(m): tg[m] for m in MULTS}, "arm": arm,
            "fills": {str(m): fills[m].strftime(F) for m in fills},
            "trail": {"t": trail[0].strftime(F), "px": trail[1]} if trail else None,
            "peak": peak, "peak_t": peak_t.strftime(F) if peak_t else "",
            "last": last, "last_t": last_t.strftime(F) if last_t else ""}


def refresh(rec, ticks, mis, now):
    """重算一筆還沒定案的紀錄：逐筆涵蓋的部分用逐筆，之後用即時 1 分K（存在 rec['bars']，隔天被逐筆取代）。"""
    key = (rec["code"], rec["strike"], rec["cp"])
    exp_t = dt.datetime.strptime(rec["exp_t"], F)
    entry_at = dt.datetime.strptime(rec["entry_at"], F)
    cover = ticks.covered_to()
    seq = ticks.path(key, entry_at, exp_t) if cover and cover > entry_at else []
    cut = max(cover, entry_at) if cover else entry_at
    bars = {b[0]: b for b in rec.get("bars", [])}
    if mis and now < exp_t + dt.timedelta(hours=1):
        try:
            for t, o, h, l, c in mis.bars(key):
                bars[t.strftime(F)] = [t.strftime(F), o, h, l, c]
        except Exception as e:
            rec["warn"] = f"即時 1 分K 抓取失敗：{e}"
        else:
            rec.pop("warn", None)
    keep = sorted(b for b in bars.values() if dt.datetime.strptime(b[0], F) > cut)
    for b in keep:
        t = dt.datetime.strptime(b[0], F)
        if t <= exp_t:
            seq += [(t, b[1]), (t, b[2]), (t, b[3]), (t, b[4])]       # 1 分K 內先高後低（保守估回吐）
    rec.update(simulate(rec["entry"], seq))
    rec["final"] = bool(cover and cover >= exp_t)
    rec["closed"] = rec["final"] or now >= exp_t
    if rec["final"]:
        rec.pop("bars", None)
    else:
        rec["bars"] = keep
    q = None
    if mis and not rec["closed"]:
        try:
            q = mis.quotes().get(key)
        except Exception:
            pass
    rec["quote"] = {"bid": q["bid"], "ask": q["ask"], "last": q["last"]} if q else None
    return rec


def update(signals, now, root, post=None):
    """signals：v3 的每筆進場 [{'entry_t': datetime, 'side': ±1, 'kind': str}]（含剛成立、下一根 K 才進場的訊號）。
    post：呼叫期交所 MIS 的函式（None 表示離線，只用逐筆）。回傳給網頁用的 dict，並把狀態存回 history/option_trades.json。"""
    path = os.path.join(root, "history", "option_trades.json")
    try:
        state = json.load(open(path, encoding="utf-8"))
    except Exception:
        state = {"trades": []}
    recs = {x["id"]: x for x in state["trades"]}
    ticks = Ticks(os.path.join(root, "history", "txo_ticks"))
    mis = Mis(post, now) if post else None
    notes = []
    for s in signals:
        t = s["entry_t"]
        sid = f"{t:%Y-%m-%d %H:%M}"
        if t < START or t > now + dt.timedelta(minutes=1):
            continue
        rec = recs.get(sid)
        late = rec is not None and rec.get("src") == "即時" and rec.get("lag", 0) > 10 and not rec.get("repriced")
        if rec is None or late:
            p = None
            try:
                p = ticks.pick(t, s["side"])
            except Exception as e:
                notes.append(f"{sid} 逐筆選檔失敗：{e}")
            if p and late:
                # 網頁晚了超過 10 分鐘才更新到的訊號：逐筆進來後，同一檔改用訊號時間後第一筆成交當買進價
                first = next(iter(ticks.path((rec["code"], rec["strike"], rec["cp"]), t - dt.timedelta(seconds=1),
                                             dt.datetime.strptime(rec["exp_t"], F))), None)
                if first:
                    rec.update(entry=first[1], entry_at=first[0].strftime(F), repriced=True)
                continue
            if p is None and rec is None and mis:
                try:
                    p = mis.pick(t, s["side"])
                except Exception as e:
                    notes.append(f"{sid} 即時報價抓取失敗：{e}")
            if p is None:
                if rec is None:
                    notes.append(f"{sid} 的訊號還選不到合約（沒有逐筆也沒有即時報價），下次更新再試")
                continue
            recs[sid] = {"id": sid, "side": s["side"], "kind": s["kind"], "code": p["code"], "strike": p["strike"],
                         "cp": p["cp"], "exp_t": p["exp_t"].strftime(F), "entry": p["entry"],
                         "entry_at": p["entry_at"].strftime(F), "alts": p["alts"], "src": p["src"],
                         "lag": max(0, round((p["entry_at"] - t).total_seconds() / 60)),
                         "phase": "前測" if t >= FORWARD_START else "回測"}
    for rec in recs.values():
        if not rec.get("final"):
            try:
                refresh(rec, ticks, mis, now)
            except Exception as e:
                rec["warn"] = f"更新失敗：{e}"
    state = {"trades": [recs[k] for k in sorted(recs)]}
    os.makedirs(os.path.dirname(path), exist_ok=True)
    json.dump(state, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    cover = ticks.covered_to()
    return {"generated": f"{now:%Y-%m-%d %H:%M}", "ticks_to": f"{cover:%Y-%m-%d %H:%M}" if cover else "", "notes": notes,
            "rules": {"target": TARGET, "mults": MULTS, "arm": TRAIL_ARM, "back": TRAIL_BACK, "cost": COST, "point": POINT},
            "trades": [{k: v for k, v in r.items() if k != "bars"} for r in state["trades"]]}
