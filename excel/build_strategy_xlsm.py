#!/usr/bin/env python3
"""產生 台指策略訊號.xlsm（即時訊號＋成交紀錄的 Excel 獨立版），也可以順便打包成 zip（附 history 的歷史資料）。

用法：
    python3 excel/build_strategy_xlsm.py                    # 輸出 excel/台指策略訊號.xlsm
    python3 excel/build_strategy_xlsm.py --zip 輸出.zip      # 另外打包 xlsm + strategy_data（history 的 1 分K、每日行情、加權指數）
    python3 excel/build_strategy_xlsm.py --test 輸出.xlsm   # 測試版：多放 STest 模組（LibreOffice 比對用）
VBA 原始碼在 excel/strategy_vba/*.bas；活頁簿用 openpyxl 建，vbaProject.bin 沿用 build_xlsm.py 的寫法自己組。
cp950 存不了的字元（≥、–、✓ 等）會自動改成 ChrW()。
"""
import argparse, io, os, sys, uuid, zipfile

import openpyxl
from openpyxl.styles import Alignment, Font

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import build_xlsm as B  # noqa: E402

SRC = os.path.join(HERE, "strategy_vba")
MODULES = ["SCalc", "SNet", "SView"]
OUT = os.path.join(HERE, "台指策略訊號.xlsm")

THISWORKBOOK_CODE = '''Option Explicit

Private Sub Workbook_Open()
    OnOpen
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    OnClose
End Sub
'''

HELP = [
    ("台指策略訊號（Excel 獨立版）", "h"),
    ("網頁版 live.html（即時訊號）與 trades.html（成交紀錄）的 Excel 版。Excel 直接向期交所抓資料、在 Excel 裡算 v3 與 D2，不靠 GitHub、不需要 Python。", None),
    ("", None),
    ("第一次使用", "b"),
    ("1. 把整個資料夾解壓縮到電腦上（台指策略訊號.xlsm 和 strategy_data 資料夾要放在一起）。", None),
    ("2. 在 台指策略訊號.xlsm 上按右鍵 →「內容」→ 最下面勾「解除封鎖」→ 確定。沒做這步，Excel 會擋掉巨集。", None),
    ("3. 打開檔案，按上方黃色列的「啟用內容」。打開時會自動更新一次，之後預設每 5 分鐘自動更新（休市時每小時）。", None),
    ("4. 要馬上更新就按「即時訊號」右上角的「立即更新」；「自動更新」按鈕可以開關自動更新。", None),
    ("5. 如果 strategy_data 資料夾不見了，第一次更新會從期交所重新下載最近 30 個交易日的每筆成交（約 3–10 分鐘）。", None),
    ("", None),
    ("資料（存在 strategy_data 資料夾，Excel 只要開著就會自己累積，不用存檔）", "b"),
    ("tx_1min\\Daily_YYYY_MM_DD.csv：台指期 1 分K。來源是期交所「前 30 個交易日期貨每筆成交資料」，每個檔案是前一晚夜盤加當天日盤。期交所只留 30 天，所以每小時檢查一次有沒有新檔，有就下載。", None),
    ("目前這一盤：期交所即時 1 分K（mis.taifex.com.tw），每次更新都重抓日盤、夜盤成交量最大的兩個契約，補上還沒存進檔案的部分。", None),
    ("tx_daily.csv：期交所台指期每日行情（2024 年起），只給 D2 日K 指標暖機用，每天更新一次。", None),
    ("twii_60min.csv：加權指數 60 分K（Yahoo），只給「成交紀錄」裡的加權指數回測用，每天更新一次；抓不到不影響訊號。", None),
    ("state.txt：上次更新時間、自動更新設定（auto=1 開 / 0 關，auto_min=間隔分鐘）。", None),
    ("隔了很多天沒開 Excel 也沒關係：只要不超過 30 個交易日，期交所的檔案都還補得回來。超過的話，那幾天 v3 會缺 60 分K（畫面會提示資料缺口）。", None),
    ("", None),
    ("規則（跟網頁版同一套算法，程式從 strategy/v3.py、strategy/daily.py、strategy/live.py 改寫）", "b"),
    ("v3（60 分K，台指期近月日盤＋夜盤）：EMA20/60 趨勢；布林突破（實體 ≥ 高低差一半、≥ 0.8 ATR、收盤過布林軌、MACD 柱同向）或拉回（碰 EMA20 後收回、過前一根高/低點、EMA60 比 5 根前同向）。", None),
    ("　只在日盤 08:45–11:45、夜盤 21:00–23:00（美國冬令 22:00–00:00）開始的 K 開新倉，下一根開盤進場。停損 1.5 ATR（最少 0.5 ATR，端點離進場 > 2 ATR 不做），收盤到 1R 保本，保本後最高收盤 − 5 ATR 移動停利。", None),
    ("D2（日K，台指期近月只算日盤 08:45–13:45）：日 EMA20 > EMA60（空單相反）、收盤創 20 日新高（新低）、實體 ≥ 高低差一半，當天 13:45 收盤前進場。", None),
    ("　停損 1 ATR（最少 0.5 ATR，端點離進場 > 1.5 ATR 不做），收盤到 1R 保本，保本後收盤跌破（空單站上）日 EMA20 隔天開盤出場。", None),
    ("換月：每個交易日用成交量最大的契約當近月，換月時把舊資料加上新舊月價差（回溯調整），所以早期價格會跟當時報價差一個價差。", None),
    ("試算：本金 20 萬、固定 1 口微台（每點 10 元）、每次下單 25 元；R 另扣 3 點滑價成本。「前測」是 2026/11/9 規則定案之後的交易。", None),
    ("加權指數只當回測用的歷史資料；即時訊號、部位、停損一律以台指期為準。", None),
    ("", None),
    ("限制", "b"),
    ("只能用 Windows 版 Excel（建議 Microsoft 365）；Mac 版與手機版不能執行巨集。電腦時間請設成台北時間。", None),
    ("Excel 要開著才會自動更新；關掉後不會在背景跑。要手機也能看，請用網頁版 https://asinkk.github.io/txo-dashboard/live.html 。", None),
    ("每次更新會重畫「即時訊號」「成交紀錄」兩張工作表，自己加的內容請放在別的工作表。", None),
    ("不知道國定假日：連假會被當成資料缺口提示，等假日過後就會消失。", None),
    ("", None),
    ("資料來源：台灣期貨交易所（www.taifex.com.tw、mis.taifex.com.tw）、Yahoo Finance。僅供參考，不構成投資建議。", None),
    ("VBA 原始碼：GitHub asinkk/txo-dashboard 的 excel/strategy_vba/；重新產生活頁簿：python3 excel/build_strategy_xlsm.py。", None),
]


def _ok(ch):
    try:
        ch.encode("cp950")
        return True
    except UnicodeEncodeError:
        return False


def vba_cp950(src):
    """把 cp950 存不了的字元改成 ChrW()：字串裡拆開接 ChrW，註解裡換成 ?。"""
    out = []
    for line in src.split("\n"):
        res, i, n = [], 0, len(line)
        while i < n:
            c = line[i]
            if c == '"':
                j, buf = i + 1, ""
                while j < n:
                    if line[j] == '"':
                        if j + 1 < n and line[j + 1] == '"':
                            buf += '""'
                            j += 2
                            continue
                        break
                    buf += line[j]
                    j += 1
                parts, cur = [], ""
                for ch in buf:
                    if _ok(ch):
                        cur += ch
                    else:
                        if cur:
                            parts.append('"' + cur + '"')
                        parts.append("ChrW(&H%04X)" % ord(ch))
                        cur = ""
                if cur or not parts:
                    parts.append('"' + cur + '"')
                res.append("(" + " & ".join(parts) + ")" if len(parts) > 1 else parts[0])
                i = j + 1
            elif c == "'":
                res.append("".join(ch if _ok(ch) else "?" for ch in line[i:]))
                i = n
            else:
                res.append(c if _ok(c) else "?")
                i += 1
        out.append("".join(res))
    return "\n".join(out)


def build_vba(sheet_codenames, modules):
    mods = [{"name": "ThisWorkbook", "doc": True,
             "src": B.DOC_ATTR.format(name="ThisWorkbook", guid=B.WB_GUID) + B.crlf(THISWORKBOOK_CODE)}]
    for cn in sheet_codenames:
        mods.append({"name": cn, "doc": True, "src": B.DOC_ATTR.format(name=cn, guid=B.WS_GUID)})
    for name in modules:
        src = open(os.path.join(SRC, name + ".bas"), encoding="utf-8").read()
        mods.append({"name": name, "doc": False, "src": B.crlf(f'Attribute VB_Name = "{name}"\n' + vba_cp950(src))})
    pid = "{" + str(uuid.uuid4()).upper() + "}"
    kids = [B.Node("_VBA_PROJECT", data=bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00])),
            B.Node("dir", data=B.ovba_compress(B.build_dir(mods, B.REFERENCES)))]
    for m in mods:
        kids.append(B.Node(m["name"], data=B.ovba_compress(m["src"].encode(B.ENC))))
    return B.write_cfb([B.Node("VBA", children=kids), B.Node("PROJECT", data=B.build_project_stream(mods, pid)),
                        B.Node("PROJECTwm", data=B.build_projectwm(mods))])


def build_workbook(path, test=False):
    wb = openpyxl.Workbook()
    wb.code_name = "ThisWorkbook"
    ws = wb.active
    ws.title = "即時訊號"
    ws.sheet_properties.codeName = "Sheet1"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 30
    ws["B1"] = "台指策略即時訊號"
    ws["B1"].font = Font(size=18, bold=True)
    ws["B3"] = "請按上方黃色列的「啟用內容」，或按右上角「立即更新」，就會開始抓資料並計算 v3 與 D2。"
    ws["B4"] = "如果沒有看到「啟用內容」而是紅色的「安全性風險」，請先關閉檔案，在檔案上按右鍵 →「內容」→ 勾「解除封鎖」。"
    ws["B5"] = "第一次使用、資料夾與規則說明見「說明」工作表。"
    t = wb.create_sheet("成交紀錄")
    t.sheet_properties.codeName = "Sheet2"
    t.sheet_view.showGridLines = False
    t["B1"] = "台指策略成交紀錄"
    t["B1"].font = Font(size=18, bold=True)
    t["B3"] = "按「即時訊號」的「立即更新」後，這裡會列出 v3 與 D2 的每一筆模擬交易。"
    h = wb.create_sheet("說明")
    h.sheet_properties.codeName = "Sheet3"
    h.sheet_view.showGridLines = False
    h.column_dimensions["A"].width = 130
    for i, (txt, st) in enumerate(HELP, start=1):
        c = h.cell(row=i, column=1, value=txt)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if st == "h":
            c.font = Font(size=16, bold=True)
        elif st == "b":
            c.font = Font(size=12, bold=True)
    buf = io.BytesIO()
    wb.save(buf)
    vba = build_vba(["Sheet1", "Sheet2", "Sheet3"], MODULES + (["STest"] if test else []))
    src = zipfile.ZipFile(io.BytesIO(buf.getvalue()))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                s = data.decode("utf-8")
                s = s.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                              "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
                s = s.replace("<Default ", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ', 1)
                data = s.encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                s = data.decode("utf-8")
                s = s.replace("</Relationships>",
                              '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                              'Target="vbaProject.bin"/></Relationships>')
                data = s.encode("utf-8")
            z.writestr(item, data)
        z.writestr("xl/vbaProject.bin", vba)


def package(xlsm, out_zip):
    """xlsm + strategy_data（history 的台指期 1 分K、每日行情、加權指數 60 分K）。"""
    hist = os.path.join(ROOT, "history")
    top = "台指策略Excel"
    with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(xlsm, f"{top}/{os.path.basename(xlsm)}")
        for f in sorted(os.listdir(os.path.join(hist, "tx_1min"))):
            if f.startswith("Daily_") and f.endswith(".csv"):
                z.write(os.path.join(hist, "tx_1min", f), f"{top}/strategy_data/tx_1min/{f}")
        for f in ("tx_daily.csv", "twii_60min.csv"):
            z.write(os.path.join(hist, f), f"{top}/strategy_data/{f}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("out", nargs="?", default=OUT)
    ap.add_argument("--zip", help="另外打包成 zip（附歷史資料）")
    ap.add_argument("--test", action="store_true", help="多放 STest 測試模組")
    a = ap.parse_args()
    build_workbook(a.out, a.test)
    print("已輸出", a.out)
    if a.zip:
        package(a.out, a.zip)
        print("已打包", a.zip)
