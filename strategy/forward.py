"""v3 前測：用 history/tx_1min 累積的台指期日夜盤 60 分K 跑 v3，輸出模擬交易紀錄與網頁。
FORWARD_START 之後的訊號才算前測（策略規則在這之前就定好，沒有偷看）。
試算條件：本金 20 萬、固定 1 口微台（每點 10 元）、每次下單 25 元。
資料每天台北 06:30 更新到前一晚夜盤，所以網頁是事後紀錄，不是即時訊號。
用法：python strategy/forward.py [1分K資料夾] [輸出 html] [輸出 csv]
"""
import csv, datetime as dt, html, os, sys
from build_bars import build
from backtest import indicators, tx_revised, COST
from v3 import run, V3, stats

FORWARD_START = dt.datetime(2026, 11, 9, 8, 45)   # 規則定案後第一個交易日
CAP, MULT, FEE = 200000, 10, 25

src = sys.argv[1] if len(sys.argv) > 1 else os.path.join("history", "tx_1min")
out_html = sys.argv[2] if len(sys.argv) > 2 else os.path.join("docs", "strategy.html")
out_csv = sys.argv[3] if len(sys.argv) > 3 else os.path.join("history", "forward_trades.csv")

rows = build(src, verbose=False)
T = [r[0] for r in rows]
bars = (T, [r[1] for r in rows], [r[2] for r in rows], [r[3] for r in rows], [r[4] for r in rows])
ind = indicators(*bars)
trades = run(bars, ind, tx_revised, "12:45", V3)
for x in trades:
    x["twd"] = (x["r"] * x["R"] + COST) * MULT - (FEE if x["open"] else 2 * FEE)
    x["phase"] = "前測" if x["t"] >= FORWARD_START else "回測"

with open(out_csv, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["階段", "訊號K", "類型", "方向", "進場時間", "進場價", "初始停損", "目前/出場停損", "出場時間", "出場價",
                "R", "損益點", "損益元(1口微台)", "狀態"])
    for x in trades:
        w.writerow([x["phase"], f"{x['t']:%Y-%m-%d %H:%M}", x["kind"], "多" if x["side"] > 0 else "空",
                    f"{x['entry_t']:%Y-%m-%d %H:%M}", f"{x['entry']:.0f}", f"{x['stop0']:.0f}", f"{x['stop']:.0f}",
                    "" if x["open"] else f"{x['exit_t']:%Y-%m-%d %H:%M}", "" if x["open"] else f"{x['exit']:.0f}",
                    f"{x['r']:+.2f}", f"{x['r'] * x['R'] + COST:+.0f}", f"{x['twd']:+.0f}", "持倉中" if x["open"] else "已出場"])

fwd = [x for x in trades if x["phase"] == "前測"]
bt = [x for x in trades if x["phase"] == "回測"]
last_t, last_c = T[-1], bars[4][-1]
open_pos = next((x for x in trades if x["open"]), None)


def money(v):
    cls = "up" if v > 0 else "dn" if v < 0 else ""
    return f'<span class="{cls}">{v:+,.0f}</span>'


def table(ts):
    if not ts:
        return '<p class="note">還沒有交易。</p>'
    body = []
    for x in reversed(ts):
        body.append("<tr>" + "".join(f"<td>{c}</td>" for c in (
            f"{x['entry_t']:%m/%d %H:%M}", x["kind"], "多" if x["side"] > 0 else "空", f"{x['entry']:,.0f}",
            f"{x['stop0']:,.0f}", "持倉中" if x["open"] else f"{x['exit_t']:%m/%d %H:%M}",
            f"{last_c:,.0f}" if x["open"] else f"{x['exit']:,.0f}", f"{x['r']:+.2f}R", money(x["twd"])
        )) + "</tr>")
    head = "".join(f"<th>{h}</th>" for h in ("進場", "類型", "方向", "進場價", "初始停損", "出場", "出場價", "R", "損益（元）"))
    return f'<div class="scroll"><table><thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table></div>'


def summary_tiles(ts):
    s = stats([x for x in ts if not x["open"]])
    total = sum(x["twd"] for x in ts)
    eq = CAP + total
    return f"""<div class="tiles">
  <div class="tile"><h3>帳戶（20 萬起）</h3><p class="big">{eq:,.0f}</p><p class="sub">{money(total)} 元（{total / CAP:+.1%}）</p></div>
  <div class="tile"><h3>已出場筆數</h3><p class="big">{s['n']}</p><p class="sub">勝率 {s['win']:.0%}</p></div>
  <div class="tile"><h3>合計 R</h3><p class="big">{s['tot']:+.1f}R</p><p class="sub">最大回落 {s['dd']:.1f}R</p></div>
</div>"""


if open_pos:
    x = open_pos
    pos = (f"<p><b>{'多單' if x['side'] > 0 else '空單'}</b>（{x['kind']}），{x['entry_t']:%m/%d %H:%M} 進場 {x['entry']:,.0f}。"
           f"目前停損價 <b>{x['stop']:,.0f}</b>{'（已保本）' if x['be'] else ''}，"
           f"未實現 {money(x['twd'])} 元。</p>")
else:
    pos = "<p><b>空手</b>，等待下一個訊號。</p>"

page = f"""<!doctype html>
<html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>台指 v3 前測</title>
<style>
:root {{ --bg:#f6f7f9; --panel:#fff; --fg:#1b1f24; --muted:#5d6673; --line:#dfe3e8; --up:#c2412d; --dn:#2f7d5b;
  --font:"Noto Sans TC","PingFang TC","Microsoft JhengHei",system-ui,sans-serif; }}
@media (prefers-color-scheme: dark) {{ :root {{ --bg:#14171b; --panel:#1c2026; --fg:#e8ebef; --muted:#9aa3ae; --line:#2c323a;
  --up:#ef7a63; --dn:#5cc296; color-scheme:dark; }} }}
body {{ margin:0; background:var(--bg); color:var(--fg); font-family:var(--font); font-size:14px; }}
main {{ max-width:960px; margin:0 auto; padding:16px; }}
h1 {{ font-size:20px; margin:0; }} h2 {{ font-size:16px; margin:24px 0 8px; }}
.asof, .note, .sub {{ color:var(--muted); font-size:12px; margin:4px 0; }}
.tiles {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:8px; }}
.tile, .card {{ background:var(--panel); border:1px solid var(--line); border-radius:8px; padding:12px; }}
.tile h3 {{ margin:0; font-size:12px; font-weight:500; color:var(--muted); }}
.big {{ font-size:22px; margin:4px 0 0; font-variant-numeric:tabular-nums; }}
.scroll {{ overflow-x:auto; background:var(--panel); border:1px solid var(--line); border-radius:8px; }}
table {{ border-collapse:collapse; width:100%; font-variant-numeric:tabular-nums; }}
th, td {{ padding:6px 8px; text-align:right; border-bottom:1px solid var(--line); white-space:nowrap; }}
th {{ color:var(--muted); font-weight:500; }} th:nth-child(-n+3), td:nth-child(-n+3) {{ text-align:left; }}
.up {{ color:var(--up); }} .dn {{ color:var(--dn); }}
a {{ color:inherit; }}
</style></head><body><main>
<h1>台指 60 分K 策略 v3 前測</h1>
<p class="asof">資料到 {last_t:%Y-%m-%d %H:%M} 那根 K（收盤 {last_c:,.0f}）。每天台北 06:30 更新，所以這是事後模擬紀錄，不是即時訊號。</p>
<div class="card"><h3 style="margin:0 0 4px;font-size:12px;color:var(--muted);font-weight:500">目前部位（模擬）</h3>{pos}</div>
<h2>前測：{FORWARD_START:%Y/%m/%d} 起</h2>
<p class="note">規則在這天之前就定好，之後的結果才算真正的樣本外驗證。試算：本金 20 萬、1 口微台、每次下單 25 元。</p>
{summary_tiles(fwd)}
<div style="height:8px"></div>
{table(fwd)}
<h2>前測開始前（回測，參考用）</h2>
<p class="note">{trades[0]['t'] if trades else ''} 起的累積資料。價格已換算成目前近月合約。</p>
{summary_tiles(bt)}
<div style="height:8px"></div>
{table(bt)}
<p class="note" style="margin-top:16px">規則：見 strategy/v3.py 的 V3。<a href="index.html">回選擇權盤面</a></p>
</main></body></html>
"""
os.makedirs(os.path.dirname(out_html), exist_ok=True)
open(out_html, "w", encoding="utf-8").write(page)
print("輸出", out_html, out_csv, "；前測", len(fwd), "筆，回測", len(bt), "筆，持倉", bool(open_pos))
