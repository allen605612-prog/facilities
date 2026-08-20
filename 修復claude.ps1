<#
修復 npm 全域安裝的 claude-code 損毀問題（原生執行檔沒裝好，claude.exe 變成幾百 bytes 的錯誤 stub）。
症狀：cmd/PowerShell 執行 claude 出現「與您執行的 Windows 版本不相容」或
     「所指定的可執行檔不是這個作業系統平台的有效應用程式」。
用法：在 PowerShell 執行 .\修復claude.ps1
#>

$ErrorActionPreference = "Stop"

function Write-Step($msg) {
    Write-Host ""
    Write-Host "== $msg ==" -ForegroundColor Cyan
}

Write-Step "定位 npm 全域安裝路徑"
$npmRoot = (npm root -g).Trim()
$pkgDir = Join-Path $npmRoot "@anthropic-ai\claude-code"
$exePath = Join-Path $pkgDir "bin\claude.exe"
Write-Host "套件目錄: $pkgDir"

function Get-ExeStatus {
    if (-not (Test-Path $exePath)) { return "missing" }
    $size = (Get-Item $exePath).Length
    if ($size -lt 100KB) { return "stub" }
    return "ok"
}

$status = Get-ExeStatus
Write-Host "目前 claude.exe 狀態: $status"

if ($status -eq "ok") {
    Write-Host "claude.exe 看起來是正常的執行檔，先直接驗證是否能跑。" -ForegroundColor Yellow
    try {
        & $exePath --version
        Write-Host "claude 可正常執行，不需要修復。" -ForegroundColor Green
        exit 0
    } catch {
        Write-Host "檔案存在但無法執行，繼續走修復流程。" -ForegroundColor Yellow
    }
}

Write-Step "移除舊的 npm 全域安裝"
npm uninstall -g @anthropic-ai/claude-code 2>&1 | Out-Host

Write-Step "清除 npm cache"
npm cache clean --force 2>&1 | Out-Host

Write-Step "重新安裝 @anthropic-ai/claude-code"
npm install -g @anthropic-ai/claude-code 2>&1 | Out-Host

Write-Step "檢查原生執行檔是否安裝成功"
$status = Get-ExeStatus
Write-Host "安裝後狀態: $status"

if ($status -ne "ok") {
    Write-Host "原生 optional dependency 還是沒抓到，手動指定平台套件安裝。" -ForegroundColor Yellow

    $arch = $env:PROCESSOR_ARCHITECTURE
    $platformPkg = if ($arch -eq "ARM64") {
        "@anthropic-ai/claude-code-win32-arm64"
    } else {
        "@anthropic-ai/claude-code-win32-x64"
    }
    Write-Host "偵測到架構 $arch，安裝套件: $platformPkg"

    npm install -g $platformPkg 2>&1 | Out-Host

    Write-Step "重跑 postinstall 建立正確的 claude.exe"
    $installCjs = Join-Path $pkgDir "install.cjs"
    node $installCjs 2>&1 | Out-Host

    $status = Get-ExeStatus
    Write-Host "重試後狀態: $status"
}

if ($status -eq "ok") {
    Write-Step "驗證 claude 可正常執行"
    & $exePath --version
    Write-Host ""
    Write-Host "修復完成！請重新開一個新的終端機視窗，再輸入 claude 測試。" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "自動修復失敗，claude.exe 仍然異常（狀態: $status）。" -ForegroundColor Red
    Write-Host "建議改用官方安裝腳本：" -ForegroundColor Red
    Write-Host '  npm uninstall -g @anthropic-ai/claude-code' -ForegroundColor Red
    Write-Host '  irm https://claude.ai/install.ps1 | iex' -ForegroundColor Red
    exit 1
}

