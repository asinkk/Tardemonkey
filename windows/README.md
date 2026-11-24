# 在 Windows 電腦跑即時訊號

2026-11-24 起，交易時段每半小時的即時訊號改在這台電腦跑，不再用 GitHub Actions 排程
（原版 txo-dashboard 因為排程鬧鐘與高頻排程被 GitHub 以違反 ToS 停用）。
每天 06:32 的資料收集也在這台電腦跑（`run_live.ps1 -Daily`），GitHub 上已經沒有排程。

需要：Git for Windows（已登入、能 push 這個 repo）、Python 3（只用標準函式庫）、電腦時區是台北。

安裝（PowerShell）：

```powershell
cd C:\Users\mika\Desktop\claude
git clone https://github.com/asinkk/Tardemonkey.git tardemonkey-live
cd tardemonkey-live
powershell -NoProfile -ExecutionPolicy Bypass -File windows\run_live.ps1   # 試跑一次，看 windows\logs
powershell -NoProfile -ExecutionPolicy Bypass -File windows\install_task.ps1
```

注意：
- `tardemonkey-live` 資料夾專門給排程用，每次跑都會先對齊 GitHub 最新 main，本機修改會被蓋掉。
- 電腦要開著（睡眠時會嘗試喚醒；關機時錯過的那次開機後補跑一次）。
- 只有訊號、部位或週選紀錄有變才推上 GitHub（`should_push.py`），每個盤收盤那次（13:50、05:06）一定推一次；平常一天大約 3 次 commit。
- 紀錄在 `windows\logs\live-YYYY-MM.log`（每天收資料是 `daily-YYYY-MM.log`）。
