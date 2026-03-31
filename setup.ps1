# Exshell PowerShell 設定スクリプト
# $PROFILE に追記するか、手動で読み込んでください。
# 例: . D:\dev\exshell\setup.ps1

$ExshellBin = "$PSScriptRoot\src\Exshell\bin\Release\net8.0-windows\win-x64\publish\exshell.exe"

if (-not (Test-Path $ExshellBin)) {
    Write-Warning "exshell.exe が見つかりません: $ExshellBin"
    Write-Warning "先に 'dotnet publish' を実行してください。"
    return
}

function eopen { & $ExshellBin eopen @Args }
function els   { & $ExshellBin els   @Args }
function ecat  { & $ExshellBin ecat  @Args }
function cate  { & $ExshellBin cate  @Args }
function ediff { & $ExshellBin ediff @Args }
function einfo { & $ExshellBin einfo @Args }

# WSL 補助コマンド（仕様書 §10）
function uls   { wsl ls    @Args }
function ucat  { wsl cat   @Args }
function unl   { wsl nl    @Args }
function udiff { wsl diff  @Args }
function ugrep { wsl grep  @Args }
function used  { wsl sed   @Args }
function uawk  { wsl awk   @Args }
function usort { wsl sort  @Args }
function uuniq { wsl uniq  @Args }

Write-Host "Exshell loaded. Commands: eopen, els, ecat, cate, ediff, einfo"
