#!/usr/bin/env python3
"""台指週選盤面：抓期交所資料 → 計算 → 輸出單一 HTML 圖表頁。只用 Python 標準函式庫。

用法：
    python3 txo_report.py               # 輸出 docs/index.html，並把當次資料存到 history/
    python3 txo_report.py -o 報表.html   # 指定輸出檔
"""
import argparse, datetime as dt, json, os, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import txo_data  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--out", default=os.path.join(HERE, "docs", "index.html"))
    ap.add_argument("--history", default=os.path.join(HERE, "history"), help="資料存檔資料夾，給空字串則不存")
    ap.add_argument("--history-at-close", action="store_true",
                    help="只在日盤收盤後（13:40–15:00）與夜盤收盤前後（04:30–05:30）存檔，避免每半小時都存")
    a = ap.parse_args()

    with tempfile.TemporaryDirectory() as tmp:
        txo_data.main(tmp)
        data = {k: json.load(open(os.path.join(tmp, k + ".json"), encoding="utf-8"))
                for k in ("summary", "strikes", "pain", "pc_history")}
    if not data["summary"]:
        sys.exit("沒有抓到任何到期的資料，未更新頁面。")
    now = dt.datetime.now(dt.timezone(dt.timedelta(hours=8)))
    data["generated"] = now.strftime("%Y-%m-%d %H:%M")

    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    html = open(os.path.join(HERE, "template.html"), encoding="utf-8").read()
    html = html.replace("/*__DATA__*/null", payload)
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as f:
        f.write(html)

    hm = now.hour * 100 + now.minute
    at_close = 1340 <= hm < 1500 or 430 <= hm < 530
    if a.history and (at_close or not a.history_at_close):
        os.makedirs(a.history, exist_ok=True)
        with open(os.path.join(a.history, now.strftime("%Y-%m-%d_%H%M") + ".json"), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    s = data["summary"][0]
    print(f"已輸出 {a.out}")
    print(f"{s['label']}：合成期貨 {s['forward']:,.0f}，預期波動 ±{s['straddle']:,.0f}（{s['range_low']:,}–{s['range_high']:,}），"
          f"ATM IV {s['atm_iv']:.1%}，P/C(OI) {s['pc_oi']}，最大痛點 {s['max_pain']:,}")


if __name__ == "__main__":
    main()
