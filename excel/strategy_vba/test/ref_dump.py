"""Python 參考結果：用 strategy/live.py 同一套程式算，輸出跟 VBA DumpAll 一樣格式的文字檔。
用法：python ref_dump.py 資料夾(含 tx_1min、tx_daily.csv、twii_60min.csv) "YYYY-MM-DD HH:MM" 測試資料夾或空字串 輸出檔
"""
import datetime as dt, json, os, shutil, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "..", "strategy"))
import live, v3, daily
from build_bars import build
from backtest import load, indicators, tx_revised

data, now_s, fix, out = sys.argv[1:5]
now = dt.datetime.strptime(now_s, "%Y-%m-%d %H:%M")
src = os.path.join(data, "tx_1min")
if fix:
    def post(ep, body):
        key = f"quotes_{body['MarketType']}" if ep == "getQuoteList" else f"chart_{body['SymbolID']}"
        return json.load(open(os.path.join(fix, key + ".json"), encoding="utf-8"))
    live.post = post
    sessions, info = live.fetch_live(now)
else:
    sessions, info = [], {}
gap = live.find_gap(live.history_last(src), sessions, now)
tmp, hist_last, added = live.merged_dir(src, sessions)
try:
    rows = build(tmp, verbose=False); m1 = build(tmp, 1, verbose=False); raw = live.read_raw_minutes(tmp)
finally:
    shutil.rmtree(tmp, ignore_errors=True)
T = [r[0] for r in rows]
hb = (T, *([r[k] for r in rows] for k in range(1, 5)))
hv = live.evaluate(hb, lambda b, ind: v3.run(b, ind, tx_revised, "12:45", v3.V3), dt.timedelta(minutes=60))
v3c = live.v3_checks(hb, hv["ind"], len(T) - 1)
sp, i0, md2, prev = live.tx_daily_bars(os.path.join(data, "tx_daily.csv"), raw)
hd = live.evaluate(sp, lambda b, ind: daily.run(b, ind, daily.D2, sig_from=i0), dt.timedelta(days=1))
d2c = live.d2_checks(sp, hd["ind"], len(sp[0]) - 1)
live.refine_times(hv["trades"], "v3", m1)
live.refine_times(hd["trades"], "d2", md2)
tw60 = load(os.path.join(data, "twii_60min.csv"))
tws = live.trades_page.twii_backtests(tw60, daily.twii_daily(tw60), live.refine_times)

f6 = lambda x: repr(round(x, 6))
tm = lambda t: f"{t:%Y-%m-%d %H:%M}"
L = [f"N60\t{len(T)}\t{tm(T[0])}\t{tm(T[-1])}"]
L += [f"B60\t{tm(t)}\t{f6(o)}\t{f6(h)}\t{f6(l)}\t{f6(c)}" for t, o, h, l, c, v in rows]
L.append(f"M1\t{len(m1)}")
L.append(f"ND\t{len(sp[0])}\t{i0}")
L += [f"BD\t{t:%Y-%m-%d}\t{f6(o)}\t{f6(h)}\t{f6(l)}\t{f6(c)}" for t, o, h, l, c in zip(*sp)]
iv, idd = hv["ind"], hd["ind"]
L.append("IND60\t" + "\t".join(f6(iv[k][-1]) for k in ("E20", "E60", "hist", "MID", "SD", "ATR")))
L.append("INDD\t" + "\t".join(f6(idd[k][-1]) for k in ("E20", "E60", "ATR")))
def trs(nm, trades):
    for x in trades:
        L.append("\t".join(["TR", nm, tm(x["t"]), str(x["side"]), x["kind"], f6(x["R"]), f6(x["r"]), f6(x["entry"]), tm(x["entry_t"]),
                            f6(x["stop0"]), f6(x["stop"]), "" if x["open"] else tm(x["exit_t"]), f6(x["exit"]),
                            str(int(x["open"])), str(int(x["be"])), live.why_text(x)]))
trs("v3", hv["trades"]); trs("d2", hd["trades"]); trs("tw_v3", tws["v3"]); trs("tw_d2", tws["d2"])
for nm, res in (("v3", hv), ("d2", hd)):
    L.append(f"POS\t{nm}\t{res['trades'].index(res['pos']) if res['pos'] else -1}")
for nm, res in (("v3", hv), ("d2", hd)):
    if res["sig"]:
        x = res["sig"]; L.append(f"SIG\t{nm}\t{x['side']}\t{x['kind']}\t{f6(x['entry'])}\t{f6(x['stop0'])}")
for nm, res in (("v3", hv), ("d2", hd)):
    if res["pending_exit"]:
        L.append(f"PEND\t{nm}\t{live.why_text(res['pending_exit']) or '出場條件成立'}")
for strat, chk in ((1, v3c), (2, d2c)):
    for sd, side in ((1, "多"), (2, "空")):
        c = chk[side]
        for g in ("common", "突破", "拉回"):
            for x in c.get(g, []):
                L.append(f"CK\t{strat}\t{sd}\t{g}\t{int(x['ok'])}\t{x['name']}\t{x['detail']}")
for sd, side in ((1, "多"), (2, "空")):
    for strat, chk in ((1, v3c), (2, d2c)):
        st = chk[side]["stop"]
        L.append(f"STOP\t{strat}\t{sd}\t{int(st['ok'])}\t{f6(st['stop'])}\t{f6(st['raw'])}")
L += [f"GAP\t{tm(g)}" for g in gap]
open(out, "w", encoding="utf-8").write("\n".join(L) + "\n")
print("ok", len(L), "lines; live", {k: v.get("last_minute", v.get("error")) for k, v in info.items()})
