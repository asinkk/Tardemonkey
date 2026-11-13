"""把期交所「前 30 個交易日期貨每筆成交資料」整理成台指期 1 分K，存到 history/tx_1min/。

期交所只保留最近 30 個交易日，所以要每天跑一次才能慢慢累積歷史。
每個檔案 Daily_YYYY_MM_DD 含前一晚的夜盤加上當天日盤。
只用 Python 標準函式庫。

用法：python strategy/collect_ticks.py [輸出資料夾]
"""
import csv, io, os, re, sys, urllib.request, zipfile

LIST_URL = "https://www.taifex.com.tw/cht/3/dlFutPrevious30DaysSalesData"
FILE_URL = "https://www.taifex.com.tw/file/taifex/Dailydownload/DailydownloadCSV/{}"
PRODUCTS = {"TX"}          # 只收台指期（大台）；小台/微台價格相同
REFRESH_NEWEST = 3         # 最新幾個檔案每次都重抓，避免存到還沒更新完整的檔案
UA = {"User-Agent": "Mozilla/5.0 (tardemonkey)"}


def get(url):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read()


def to_1min(raw):
    """回傳 {(日期, 分鐘, 契約): [開, 高, 低, 收, 量]}"""
    bars = {}
    text = raw.decode("big5", errors="replace")
    for row in csv.reader(io.StringIO(text)):
        if len(row) < 6 or row[1].strip() not in PRODUCTS:
            continue
        contract = row[2].strip()
        if "/" in contract:  # 價差單
            continue
        try:
            price = float(row[4]); vol = int(row[5])
        except ValueError:
            continue
        key = (row[0].strip(), row[3].strip().zfill(6)[:4], contract)
        b = bars.get(key)
        if b is None:
            bars[key] = [price, price, price, price, vol]
        else:
            b[1] = max(b[1], price); b[2] = min(b[2], price); b[3] = price; b[4] += vol
    return bars


def collect(out, refresh=REFRESH_NEWEST):
    """下載 out 裡還沒有的檔案，並重抓最新 refresh 個。回傳更新的檔案數。live.py 補缺口時用 refresh=0。"""
    os.makedirs(out, exist_ok=True)
    names = sorted(set(re.findall(r"Daily_\d{4}_\d{2}_\d{2}\.zip", get(LIST_URL).decode("utf-8", "replace"))))
    if not names:
        raise RuntimeError("期交所頁面沒有列出任何檔案，可能改版了")
    newest = set(names[-refresh:]) if refresh else set()
    done = 0
    for name in names:
        dest = os.path.join(out, name.replace(".zip", ".csv"))
        if os.path.exists(dest) and name not in newest:
            continue
        try:
            z = zipfile.ZipFile(io.BytesIO(get(FILE_URL.format(name))))
            raw = z.read(z.namelist()[0])
        except Exception as e:  # 單一檔案失敗不影響其他檔案
            print("略過", name, e)
            continue
        bars = to_1min(raw)
        if not bars:
            print("沒有台指期資料", name)
            continue
        tmp = dest + ".tmp"
        with open(tmp, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["date", "time", "contract", "open", "high", "low", "close", "volume"])
            for (d, t, c), (o, h, l, cl, v) in sorted(bars.items(), key=lambda kv: (kv[0][2], kv[0][0], kv[0][1])):
                w.writerow([d, t, c, f"{o:g}", f"{h:g}", f"{l:g}", f"{cl:g}", v])
        os.replace(tmp, dest)
        done += 1
        print("已存", dest, len(bars), "筆 1 分K")
    print("完成，更新", done, "個檔案")
    return done


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join("history", "tx_1min")
    try:
        collect(out)
    except RuntimeError as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()
