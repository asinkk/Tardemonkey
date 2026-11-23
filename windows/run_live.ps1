# 即時訊號：抓期交所即時 1 分K、算 v3（限時段），有變動就推回 GitHub。
# 由工作排程器每半小時叫一次（install_task.ps1 會設好）。手動試跑：
#   powershell -NoProfile -ExecutionPolicy Bypass -File windows\run_live.ps1
# 加 -Daily：每天 06:32 收資料（原本 GitHub 的 tx-ticks.yml）：存期交所 1 分K、台指選擇權逐筆、跑 v3 前測，再算即時訊號。
#   powershell -NoProfile -ExecutionPolicy Bypass -File windows\run_live.ps1 -Daily
# 這個資料夾專門給排程用：每次都會先對齊 GitHub 最新的 main，本機改的東西會被蓋掉。
param([switch]$Daily)
$ErrorActionPreference = 'Continue'
$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo
$logDir = Join-Path $PSScriptRoot 'logs'
New-Item -ItemType Directory -Force $logDir | Out-Null
$log = Join-Path $logDir ((@('live-', 'daily-')[[int][bool]$Daily]) + (Get-Date -Format 'yyyy-MM') + '.log')
function Log($m) { '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m | Out-File -Append -Encoding utf8 $log }
function Run([string]$exe, [string[]]$argv) {
    $out = & $exe @argv 2>&1
    foreach ($l in $out) { Log "  $l" }
    return $LASTEXITCODE
}

# 同時只跑一個（上一次卡住超過 40 分鐘就不管它）
$lock = Join-Path $logDir 'running.lock'
if ((Test-Path $lock) -and (Get-Item $lock).LastWriteTime -gt (Get-Date).AddMinutes(-40)) { Log '上一次還在跑，這次跳過'; exit 0 }
Set-Content $lock $PID

$py = 'python'; $pyArgs = @()
if (Get-Command py -ErrorAction SilentlyContinue) { $py = 'py'; $pyArgs = @('-3') }
$files = @('docs/live.html', 'docs/live.json', 'docs/trades.html', 'docs/options.html', 'history/option_trades.json', 'history/tx_1min')
if ($Daily) { $files += @('history/txo_ticks', 'history/forward_trades.csv', 'docs/strategy.html') }
$code = 1
try {
    Log '開始'
    if ((Run git @('fetch', '-q', 'origin', 'main')) -ne 0) { Log 'git fetch 失敗'; exit 1 }
    Run git @('reset', '-q', '--hard', 'origin/main') | Out-Null
    if ($Daily) {
        if ((Run $py ($pyArgs + @('strategy/collect_ticks.py', 'history/tx_1min'))) -ne 0) { Log 'collect_ticks.py 失敗'; exit 1 }
        # 選擇權失敗不擋期貨與策略頁，明天會再補（期交所留 30 個交易日）
        if ((Run $py ($pyArgs + @('strategy/collect_option_ticks.py', 'history/txo_ticks'))) -ne 0) { Log 'collect_option_ticks.py 失敗（略過）' }
        Push-Location strategy
        $rc = Run $py ($pyArgs + @('forward.py', '../history/tx_1min', '../docs/strategy.html', '../history/forward_trades.csv'))
        Pop-Location
        if ($rc -ne 0) { Log 'forward.py 失敗'; exit 1 }
    }
    if ((Run $py ($pyArgs + @('strategy/live.py', '--out', 'docs'))) -ne 0) { Log 'live.py 失敗'; exit 1 }
    Run git (@('add') + $files) | Out-Null
    & git diff --cached --quiet
    if ($LASTEXITCODE -eq 0) { Log '沒有變動'; $code = 0; exit 0 }
    $msg = '策略訊號 ' + (Get-Date -Format 'yyyy-MM-dd HH:mm') + '（本機）'
    if ($Daily) { $msg = '累積台指期 1 分K、台指選擇權逐筆 ' + (Get-Date -Format 'yyyy-MM-dd') + '（本機）' }
    Run git @('-c', 'user.name=txo-bot', '-c', 'user.email=txo-bot@users.noreply.github.com', 'commit', '-q', '-m', $msg) | Out-Null
    for ($i = 1; $i -le 5; $i++) {
        # 產生的網頁／json 衝突時用這次剛算的（rebase 時 -X theirs 指自己的 commit）
        if ((Run git @('pull', '-q', '--rebase', '-X', 'theirs', 'origin', 'main')) -eq 0 -and (Run git @('push', '-q', 'origin', 'HEAD:main')) -eq 0) {
            Log "已推上 GitHub：$msg"; $code = 0; exit 0
        }
        Run git @('rebase', '--abort') | Out-Null
        Start-Sleep -Seconds (5 * $i)
    }
    Log '推送失敗'
} finally {
    Remove-Item $lock -ErrorAction SilentlyContinue
    Log "結束（$code）"
}
exit $code
