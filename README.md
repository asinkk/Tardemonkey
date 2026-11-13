# Tardemonkey

## 台指期 v3 60 分K 即時訊號

從交易猴（txo-dashboard）複製過來的 v3 60 分K 策略，只用 Python 標準函式庫。
訊號、部位、停損只用台指期近月（TX）日盤＋夜盤 60 分K；加權指數只用在歷史成交紀錄頁的回測。

| 檔案 | 用途 |
| --- | --- |
| `strategy/v3.py` | v3 規則（拉回進場、5 ATR 移動停利、日線濾網選項） |
| `strategy/backtest.py` | 指標、開倉時段（`tx_revised`）、手續費等共用函式 |
| `strategy/build_bars.py` | 1 分K 合成近月連續 60 分K（日盤 08:45 起、夜盤 15:00 起，換月回溯調整） |
| `strategy/collect_ticks.py` | 下載期交所前 30 個交易日每筆成交，轉成 1 分K 存到 `history/tx_1min/` |
| `strategy/fetch_twii.py` | 抓加權指數 60 分K（Yahoo），存到 `history/twii_60min.csv`（回測用） |
| `strategy/live.py` | 即時訊號頁：history＋期交所即時 1 分K，算 v3 限時段與不限時段目前的部位、停損與進場條件，輸出 `docs/live.html`、`docs/live.json`、`docs/trades.html` |
| `strategy/trades_page.py` | 歷史成交紀錄頁 `docs/trades.html`（由 `live.py` 產生） |
| `strategy/forward.py` | v3 前測（2026/11/9 起算），輸出 `docs/strategy.html`、`history/forward_trades.csv` |

```
python strategy/live.py --out docs              # 抓即時資料
python strategy/live.py --out /tmp/x --offline  # 只用 history
```

排程（GitHub Actions，時間可能晚 5–30 分鐘或漏跑）：
- `strategy-live.yml`：交易時段每半小時更新即時訊號頁（日盤台北 :18/:48，夜盤 :04/:34）。
- `tx-ticks.yml`：每天台北 06:30 存期交所 1 分K（期交所只留 30 個交易日）、更新加權指數、跑前測與即時訊號頁。
