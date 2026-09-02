# 產生競賽得獎匯入總表（雙擊執行，或在此資料夾按右鍵 → 使用 PowerShell 執行）
$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot
$env:PYTHONUTF8 = '1'
uvx --with openpyxl python "$PSScriptRoot\產生得獎匯入總表.py" @args
Write-Host ''
Read-Host '按 Enter 關閉'
