"""把期交所「前 30 個交易日選擇權每筆成交資料」裡的台指選擇權（TXO，含週選）存到 history/txo_ticks/。

期交所只保留最近 30 個交易日，所以要每天跑一次才能慢慢累積歷史。
每個檔案 Daily_YYYY_MM_DD 是一個交易日：前一晚 15:00 起的夜盤（到隔天 05:00）加上當天日盤，
週五夜盤算在下週一那天（所以週六就會看到下週一的檔案，週一日盤收完再重抓補齊）。

存法（為了 repo 不要長太快）：
- 只留 TXO（月選＋週三週選 W＋週五週選 F），其他個股／ETF 選擇權不要。
- 逐筆照原樣保留（不合併、不轉 K 棒），只去掉期交所檔案裡的空白欄位補齊。
- 用 xz 壓縮：原始 CSV 一天約 170 MB，壓完約 0.3–1.5 MB。
  Python 直接讀：pandas.read_csv("Daily_2026_11_13.csv.xz") 或 lzma.open(..., "rt")。

欄位：date（成交日期）, time（HHMMSS）, expiry（到期月份(週別)，例 202611、202611W4、202611F2）,
      strike（履約價）, cp（C/P）, price, qty（成交數量）, auction（開盤集合競價為 *）
排序和期交所原檔相同：依契約（履約價、到期、買賣權）再依時間。

只用 Python 標準函式庫。
用法：python strategy/collect_option_ticks.py [輸出資料夾]
"""
import csv, io, lzma, os, re, sys, urllib.request, zipfile

LIST_URL = "https://www.taifex.com.tw/cht/3/dlOptPrevious30DaysSalesData"
FILE_URL = "https://www.taifex.com.tw/file/taifex/Dailydownload/OptionsDailydownloadCSV/Options{}"
PRODUCTS = {"TXO"}
REFRESH_NEWEST = 3         # 最新幾個檔案每次都重抓，避免存到還沒更新完整的檔案（例如只有週五夜盤的週一檔）
HEADER = ["date", "time", "expiry", "strike", "cp", "price", "qty", "auction"]
UA = {"User-Agent": "Mozilla/5.0 (tardemonkey)"}


def get(url):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=180) as r:
        return r.read()


def to_rows(raw):
    """期交所原檔 → [[date, time, expiry, strike, cp, price, qty, auction], ...]"""
    rows = []
    text = raw.decode("big5", errors="replace")
    for r in csv.reader(io.StringIO(text)):
        if len(r) < 8 or r[1].strip() not in PRODUCTS:
            continue
        r = [x.strip() for x in r]
        if not r[7].isdigit():
            continue
        rows.append([r[0], r[5], r[3], r[2], r[4], r[6], r[7], r[8] if len(r) > 8 else ""])
    return rows


def collect(out, refresh=REFRESH_NEWEST):
    """下載 out 裡還沒有的檔案，並重抓最新 refresh 個。回傳更新的檔案數。"""
    os.makedirs(out, exist_ok=True)
    page = get(LIST_URL).decode("utf-8", "replace")
    names = sorted(set(re.findall(r"OptionsDaily_(\d{4}_\d{2}_\d{2})\.zip", page)))
    if not names:
        raise RuntimeError("期交所頁面沒有列出任何檔案，可能改版了")
    newest = set(names[-refresh:]) if refresh else set()
    done = 0
    for day in names:
        dest = os.path.join(out, f"Daily_{day}.csv.xz")
        if os.path.exists(dest) and day not in newest:
            continue
        try:
            z = zipfile.ZipFile(io.BytesIO(get(FILE_URL.format(f"Daily_{day}.zip"))))
            raw = z.read(z.namelist()[0])
        except Exception as e:  # 單一檔案失敗不影響其他檔案
            print("略過", day, e)
            continue
        rows = to_rows(raw)
        if not rows:
            print("沒有台指選擇權資料", day)
            continue
        buf = io.StringIO()
        w = csv.writer(buf, lineterminator="\n")
        w.writerow(HEADER)
        w.writerows(rows)
        data = buf.getvalue().encode("utf-8")
        if os.path.exists(dest):
            with lzma.open(dest) as f:
                if f.read() == data:  # 內容沒變就不重寫，避免 git 多一份一樣的檔案
                    print("沒變", dest)
                    continue
        tmp = dest + ".tmp"
        with open(tmp, "wb") as f:
            f.write(lzma.compress(data, preset=9))
        os.replace(tmp, dest)
        done += 1
        print("已存", dest, len(rows), "筆", f"{os.path.getsize(dest) / 1e6:.2f} MB")
    print("完成，更新", done, "個檔案")
    return done


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join("history", "txo_ticks")
    try:
        collect(out)
    except RuntimeError as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()
