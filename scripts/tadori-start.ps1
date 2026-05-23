# ============================================================================
# Tadori 起動スクリプト (デスクトップショートカットから実行)
#   1. ローカルポートをチェックして中継サーバが未起動なら起動
#   2. ブラウザで指定の SharePoint サイト URL を開く
# ============================================================================
#
# 使い方:
#   1. tadori-start.bat をダブルクリック (PS Execution Policy を回避するため bat 経由を推奨)
#   2. もしくは直接: powershell -ExecutionPolicy Bypass -File .\tadori-start.ps1
#
# デスクトップショートカット作成:
#   - tadori-start.bat を右クリック → 送る → デスクトップ (ショートカット)
#   - もしくは tadori-start.bat 自体をデスクトップへドラッグでショートカット作成
#
# 起動先 SharePoint サイトの指定:
#   既定: tadori-ai-relay.env の TADORI_SITE_URL を読む
#   コマンドラインで上書き: .\tadori-start.ps1 -SiteUrl 'https://contoso.sharepoint.com/sites/xxx'
# ============================================================================

[CmdletBinding()]
param(
    [string]$SiteUrl,
    [int]$Port,
    [string]$EnvFile
)

$ErrorActionPreference = 'Stop'
$scriptDir = $PSScriptRoot

# 未捕捉エラーで即終了してしまうと bat 側でも一瞬で閉じるので、本体を try/catch で
# 包んでエラーメッセージを画面に残してから終了する。
trap {
    Write-Host ''
    Write-Host '[tadori-start] 予期しないエラーで終了します:' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.ScriptStackTrace) { Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray }
    Read-Host '何かキーを押して終了'
    exit 1
}

# ─── .env から既定値を読む (relay と共通の env ファイル) ─────────────────────
if (-not $EnvFile) { $EnvFile = Join-Path $scriptDir 'tadori-ai-relay.env' }
if (Test-Path -LiteralPath $EnvFile) {
    foreach ($raw in Get-Content -LiteralPath $EnvFile -Encoding UTF8) {
        $line = $raw.Trim()
        if (-not $line -or $line.StartsWith('#')) { continue }
        $eq = $line.IndexOf('=')
        if ($eq -lt 1) { continue }
        $key = $line.Substring(0, $eq).Trim()
        $val = $line.Substring($eq + 1).Trim()
        if (($val.StartsWith('"') -and $val.EndsWith('"')) -or ($val.StartsWith("'") -and $val.EndsWith("'"))) {
            $val = $val.Substring(1, $val.Length - 2)
        }
        if (-not [Environment]::GetEnvironmentVariable($key)) {
            [Environment]::SetEnvironmentVariable($key, $val)
        }
    }
}

if (-not $SiteUrl) { $SiteUrl = $env:TADORI_SITE_URL }
if (-not $Port)    { $Port = if ($env:TADORI_AI_PORT) { [int]$env:TADORI_AI_PORT } else { 18080 } }

if (-not $SiteUrl) {
    Write-Host '⚠ TADORI_SITE_URL が未設定です。tadori-ai-relay.env に追記するか、引数 -SiteUrl で指定してください。' -ForegroundColor Yellow
    Write-Host '  例: TADORI_SITE_URL=https://contoso.sharepoint.com/sites/xxx'
    Read-Host '何かキーを押して終了'
    exit 1
}

# ─── relay の起動チェック (ポート LISTEN) ───────────────────────────────────
function Test-PortListening {
    param([int]$Port)
    try {
        $conn = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction Stop
        return [bool]$conn
    } catch {
        # Get-NetTCPConnection が無い古い PS の場合は netstat
        try {
            $found = netstat -ano | Select-String -SimpleMatch ":${Port} " | Select-String -SimpleMatch 'LISTENING'
            return [bool]$found
        } catch { return $false }
    }
}

if (Test-PortListening -Port $Port) {
    Write-Host ("[tadori-start] 中継サーバはすでに起動しています (port {0})。" -f $Port) -ForegroundColor Green
} else {
    $relay = Join-Path $scriptDir 'tadori-ai-relay.ps1'
    if (-not (Test-Path -LiteralPath $relay)) {
        Write-Host "⚠ 中継サーバスクリプトが見つかりません: $relay" -ForegroundColor Red
        Read-Host '何かキーを押して終了'
        exit 1
    }
    Write-Host ("[tadori-start] 中継サーバを起動します (port {0})…" -f $Port)
    # 別ウィンドウで relay を起動 (閉じれば停止)。-NoExit でログを残す。
    Start-Process powershell -ArgumentList @(
        '-NoExit', '-ExecutionPolicy', 'Bypass', '-File', "`"$relay`""
    ) -WindowStyle Minimized
    # 立ち上がり待ち
    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 400
        if (Test-PortListening -Port $Port) { break }
    }
    if (Test-PortListening -Port $Port) {
        Write-Host '[tadori-start] 中継サーバが起動しました。' -ForegroundColor Green
    } else {
        Write-Host '⚠ 中継サーバの起動を確認できませんでした (タイムアウト)。最小化したウィンドウのログを確認してください。' -ForegroundColor Yellow
    }
}

# ─── SharePoint サイトをブラウザで開く ──────────────────────────────────────
Write-Host "[tadori-start] ブラウザで開く: $SiteUrl"
Start-Process $SiteUrl
