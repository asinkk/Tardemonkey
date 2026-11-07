# 台指期 60 分K 策略回測

策略說明在 Claude 文件「台指期 60 分K 波段策略（修訂版）」。這個資料夾放資料累積與回測程式，只用 Python 標準函式庫。

| 檔案 | 用途 |
| --- | --- |
| `collect_ticks.py` | 下載期交所「前 30 個交易日期貨每筆成交資料」，把台指期（TX）轉成 1 分K，存到 `history/tx_1min/` |
| `build_bars.py` | 把 1 分K 合成近月連續的 60 分K（日盤 08:45 起、夜盤 15:00 起，換月回溯調整） |
| `backtest.py` | 回測草稿版與修訂版規則 |

期交所只留最近 30 個交易日，`.github/workflows/tx-ticks.yml` 每天台北 06:30 自動存一次，資料會越累積越長。

```
python strategy/build_bars.py history/tx_1min tx_60min.csv
python strategy/backtest.py tx_60min.csv tx
```

`build_bars.py` 第三個參數可以改 K 棒分鐘數，例如 `15`。
