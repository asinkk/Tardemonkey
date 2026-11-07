# 台指週選盤面（獨立版）

不需要 Claude，也不耗 token。程式直接抓台灣期交所公開資料，計算預期波動、IV、未平倉牆、P/C、最大痛點，輸出一個 HTML 圖表頁。只用 Python 3 標準函式庫，不用安裝任何套件。

## 檔案

| 檔案 | 用途 |
| --- | --- |
| `txo_report.py` | 主程式：抓資料、計算、輸出 `docs/index.html`，並把當次資料存到 `history/` |
| `txo_data.py` | 抓資料與計算的邏輯 |
| `template.html` | 圖表頁樣板 |
| `run_txo.bat` / `run_txo.sh` | 給電腦排程器用的啟動檔 |
| `.github/workflows/txo.yml` | GitHub Actions 每日排程 |
| `strategy/` | 台指期 60 分K（v3）與日K（D2）策略的資料累積、回測與即時訊號頁 `docs/live.html`（見 `strategy/README.md`） |

## 手動跑

```
python3 txo_report.py
```

然後用瀏覽器打開 `docs/index.html`。頁面的圖表函式庫從 cdnjs 載入，所以打開時需要網路。

## 每日自動跑：選一種

### A. 在自己電腦上排程（最簡單）

電腦需要在排程時間開著。

- **Windows**：開「工作排程器」→「建立基本工作」→ 觸發程序選每週，勾週一到週五，時間 21:05 → 動作選「啟動程式」，程式填 `run_txo.bat` 的完整路徑。
- **macOS**：終端機執行 `crontab -e`，加入一行（路徑換成你的）：
  `5 21 * * 1-5 /Users/你的帳號/txo/run_txo.sh`

### B. GitHub Actions + GitHub Pages（電腦不用開，手機也能看）

1. 在 GitHub 建一個新的 repo（可設 private，但 private repo 的 Pages 需要付費方案；public 則免費）。
2. 把這個資料夾全部上傳到 repo 根目錄（包含 `.github` 資料夾）。
3. repo 的 Settings → Pages → Source 選「Deploy from a branch」，branch 選 `main`，資料夾選 `/docs`。
4. Actions 分頁 → 選「台指週選盤面」→ Run workflow，先手動跑一次。
5. 之後交易日每半小時自動更新：日盤 08:45–13:45、夜盤 15:15–隔天 04:45（台北時間）。網址是 `https://<你的帳號>.github.io/<repo 名稱>/`。history/ 只在日盤收盤後與夜盤收盤時各存一份。

GitHub 的排程常會延遲 5 到 30 分鐘。GitHub 主機在國外，如果期交所擋了海外連線就會失敗（目前測試可以連）。

## 注意

- 夜盤報價是即時的；未平倉量是最近一次日盤收盤的數字。
- 想要日盤收盤版，可以再多排一次 15:30。這時 OI 已更新，但夜盤剛開，報價會比較稀。
- 僅為資料整理，不構成投資建議。

## Excel 巨集版（不需要 Python、不靠 GitHub）

`excel/台指週選盤面.xlsm`：打開後按「更新資料」，Excel 直接向期交所與 Yahoo Finance 抓資料、計算並畫圖，內容和網頁一樣（各到期指標、OI 牆、OI 增減、IV 微笑、最大痛點、法人與大額交易人、台指期、VIX/HV、美股）。

1. 下載後在檔案上按右鍵 →「內容」→ 勾「解除封鎖」→ 確定（否則 Excel 會擋巨集）。
2. 打開檔案，按黃色列的「啟用內容」，之後每次打開都會自動更新。
3. 只能用 Windows 版 Excel；電腦時間請設台北時間。

巨集原始碼在 `excel/TXO.bas`；改完執行 `python3 excel/build_xlsm.py` 重新產生 .xlsm（需要 openpyxl）。
