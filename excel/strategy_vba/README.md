# 台指策略訊號（Excel 獨立版）VBA 原始碼

`excel/台指策略訊號.xlsm` 的巨集。網頁 `docs/live.html`（即時訊號）與 `docs/trades.html`（成交紀錄）的 Excel 版：
Excel 直接向期交所抓資料、在 Excel 裡算 v3 與 D2，不靠 GitHub、不需要 Python。

| 模組 | 內容 |
| --- | --- |
| `SCalc.bas` | 計算：1 分K → 60 分K／日K（換月回溯調整）、指標、v3.run、daily.run（D2）、進出場時間到分鐘、條件明細。照 `strategy/live.py`、`build_bars.py`、`backtest.py`、`v3.py`、`daily.py` 改寫 |
| `SNet.bas` | 抓資料：期交所 MIS 即時 1 分K、每筆成交 zip → `tx_1min/Daily_*.csv`、每日行情 → `tx_daily.csv`、Yahoo 加權指數 → `twii_60min.csv` |
| `SView.bas` | 更新流程、自動更新（Application.OnTime）、「即時訊號」「成交紀錄」工作表 |
| `STest.bas` | 測試入口（只放在測試版活頁簿） |

資料放在活頁簿旁邊的 `strategy_data/`，格式跟 repo 的 `history/` 一樣，可以直接把 `history/tx_1min`、`tx_daily.csv`、`twii_60min.csv` 複製過去。

產生活頁簿：`python3 excel/build_strategy_xlsm.py [--zip 輸出.zip]`（zip 會附上 history 的資料）。

## 驗證

`test/ref_dump.py` 用 `strategy/live.py` 的函式算出參考結果；`STest.RunOffline` 在 LibreOffice 裡用同一份資料跑 VBA，`test/compare.py` 逐行比對
（60 分K、日K、指標、v3／D2／加權指數回測每一筆交易、條件明細、停損、資料缺口）。`test/fixture_20261107/` 是 11/7 存的期交所即時 1 分K，
用來測「history 少了最後兩天、靠即時資料補」的情況。2026-11-07 兩種情況都 0 差異；成交紀錄四個帳戶摘要也跟網頁一樣。
網路下載（MSXML、解壓縮）只能在 Windows Excel 上測。
