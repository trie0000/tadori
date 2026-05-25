# =============================================================================
# Tadori ワンクリック起動 (Windows)
# =============================================================================
#
# やること:
#   1) tadori-ai-relay.ps1 が起動していなければ、新しい PowerShell ウィンドウ
#      で立ち上げる (= このスクリプトを閉じてもリレーは生き続ける)
#   2) リレーの /tadori/health が 200 を返すまで最大 10 秒待機
#   3) SharePoint サイトを既定ブラウザで開く
#   4) 「ブックマークレットを押してください」の案内ダイアログを表示
#      (TopMost + DPI-aware で、ブラウザの後ろに隠れない)
#
# 設定:
#   tadori-ai-relay.env を読んで以下を解決:
#     TADORI_SITE_URL : SharePoint サイトの URL (任意。未設定なら起動時に入力)
#     TADORI_AI_PORT  : リレーの listen ポート (既定 18080)
#
#   .env に TADORI_SITE_URL=https://<tenant>.sharepoint.com/sites/<site>
#   を書いておけば毎回プロンプトを出さずに起動可能。
#
# 起動方法:
#   - エクスプローラで tadori-start.bat をダブルクリック
#   - もしくはタスクスケジューラの「ログオン時」トリガに登録して自動起動
# =============================================================================

[CmdletBinding()]
param(
    [string]$SiteUrl,
    [int]$Port,
    [string]$EnvFile
)

$ErrorActionPreference = 'Stop'

# 未捕捉エラーで即終了してしまうと bat 側でも一瞬で閉じるので、メッセージを
# 残してから終了する。
trap {
    Write-Host ''
    Write-Host '[tadori-start] 予期しないエラーで終了します:' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.ScriptStackTrace) { Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray }
    Read-Host '何かキーを押して終了'
    exit 1
}

# ─── .env 読込 (tadori-ai-relay.ps1 と同じ書式) ─────────────────────────────
if (-not $EnvFile) {
    $EnvFile = Join-Path $PSScriptRoot 'tadori-ai-relay.env'
}
if (Test-Path -LiteralPath $EnvFile) {
    try {
        foreach ($raw in (Get-Content -LiteralPath $EnvFile -Encoding UTF8)) {
            $line = $raw.Trim()
            if (-not $line) { continue }
            if ($line.StartsWith('#')) { continue }
            $eq = $line.IndexOf('=')
            if ($eq -lt 1) { continue }
            $key = $line.Substring(0, $eq).Trim()
            $val = $line.Substring($eq + 1).Trim()
            # インラインコメント (` # ...`) を削除 — クォート外のみ
            if ($val -notmatch '^["'']') {
                $hashIdx = $val.IndexOf(' #')
                if ($hashIdx -ge 0) { $val = $val.Substring(0, $hashIdx).TrimEnd() }
            }
            # 前後のクォートを剥がす
            if (($val.StartsWith('"') -and $val.EndsWith('"')) -or
                ($val.StartsWith("'") -and $val.EndsWith("'"))) {
                $val = $val.Substring(1, $val.Length - 2)
            }
            if (-not [Environment]::GetEnvironmentVariable($key)) {
                [Environment]::SetEnvironmentVariable($key, $val)
            }
        }
    } catch {
        Write-Warning ".env 読込失敗: $($_.Exception.Message)"
    }
}

# ─── SharePoint URL 解決 ───────────────────────────────────────────────────
if (-not $SiteUrl) { $SiteUrl = $env:TADORI_SITE_URL }
if (-not $SiteUrl) {
    Write-Host '' -NoNewline
    Write-Host 'SharePoint サイト URL が未設定です。' -ForegroundColor Yellow
    Write-Host 'tadori-ai-relay.env に下記を追記すると次回から自動になります:'
    Write-Host '  TADORI_SITE_URL=https://<tenant>.sharepoint.com/sites/<site>'
    Write-Host ''
    $SiteUrl = Read-Host 'SharePoint サイト URL を入力してください (空 Enter で中止)'
}
if (-not $SiteUrl) {
    Write-Host '[tadori-start] SP URL が未指定なので中止します' -ForegroundColor Yellow
    exit 1
}

# ─── ポート決定 + relay 起動状況確認 ────────────────────────────────────────
if (-not $Port) {
    $Port = if ($env:TADORI_AI_PORT) { [int]$env:TADORI_AI_PORT } else { 18080 }
}
$healthUrl = "http://127.0.0.1:$Port/tadori/health"

function Test-RelayUp {
    param([string]$Url, [int]$TimeoutSec = 1)
    try {
        $r = Invoke-WebRequest -Uri $Url -TimeoutSec $TimeoutSec -UseBasicParsing -ErrorAction Stop
        return ($r.StatusCode -eq 200)
    } catch { return $false }
}

$relayUp = Test-RelayUp -Url $healthUrl

if (-not $relayUp) {
    Write-Host "[tadori-start] リレーを新規起動 (ポート $Port)..." -ForegroundColor Cyan
    $relayPs1 = Join-Path $PSScriptRoot 'tadori-ai-relay.ps1'
    if (-not (Test-Path -LiteralPath $relayPs1)) {
        Write-Host "[tadori-start] エラー: $relayPs1 が見つかりません" -ForegroundColor Red
        exit 2
    }
    # 別ウィンドウで起動。-NoExit でリレーが止まってもウィンドウは残るので
    # 落ちた時にログが見れる。WorkingDirectory も渡して .env が見つかるように。
    Start-Process -FilePath 'powershell.exe' -ArgumentList @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-NoExit',
        '-File', "`"$relayPs1`""
    ) -WorkingDirectory $PSScriptRoot | Out-Null

    Write-Host "[tadori-start] /tadori/health の応答を待機中 (最大 10 秒)..." -ForegroundColor Cyan
    $waitMs = 0
    while ($waitMs -lt 10000) {
        Start-Sleep -Milliseconds 500
        $waitMs += 500
        if (Test-RelayUp -Url $healthUrl) {
            $relayUp = $true
            Write-Host "[tadori-start] リレー OK ($($waitMs) ms で起動完了)" -ForegroundColor Green
            break
        }
    }
    if (-not $relayUp) {
        Write-Host "[tadori-start] 警告: リレーの起動応答を確認できませんでした。SharePoint は開きますが、Tadori アプリは relay 未起動の警告が出るかもしれません。" -ForegroundColor Yellow
    }
} else {
    Write-Host "[tadori-start] リレーは既に起動済み (port $Port)" -ForegroundColor Green
}

# ─── SharePoint を既定ブラウザで開く ────────────────────────────────────────
Write-Host "[tadori-start] SharePoint を開く: $SiteUrl" -ForegroundColor Cyan
try {
    Start-Process $SiteUrl | Out-Null
} catch {
    Write-Host "[tadori-start] ブラウザ起動失敗: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "          手動で開いてください: $SiteUrl" -ForegroundColor Yellow
}

# ─── ブックマークレット案内ダイアログ ──────────────────────────────────────
# 普通の MessageBox はブラウザの後ろに隠れがちで、Windows DPI スケーリングで
# 文字がぼやけることもあるため、自前の Form を TopMost + DPI-aware + 標準
# ネイティブ ボタンで組み立てる (Spira と同じ作法)。
try {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    Add-Type -AssemblyName System.Drawing       | Out-Null

    # DPI スケーリングでのぼやけ対策。PerMonitorV2 が一番くっきり。
    # ※ Windows 10 1703 以降。古い OS は無視されるだけ。
    try { [System.Windows.Forms.Application]::SetHighDpiMode([System.Windows.Forms.HighDpiMode]::PerMonitorV2) | Out-Null } catch { }
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch { }
    try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

    $relayLine = if ($relayUp) {
        "リレー稼働中 (http://127.0.0.1:$Port)"
    } else {
        "リレー未確認 (起動に失敗している可能性あり)"
    }
    $msg = @"
Tadori の起動準備ができました。

  $relayLine
  SharePoint をブラウザで開きました

最後の手順:
  1) 開いた SharePoint ページの読込完了を待つ
  2) お気に入りバーの「Tadori」ブックマークレットをクリック
  3) アプリが画面右に開いてチャット/検索パネルが見えれば成功

このウィンドウは閉じて OK です。
リレーは別ウィンドウで動き続けます (閉じると Outlook 取り込み /
OneNote 連携 / 自動取り込みが止まります)。
"@

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Tadori 起動完了'
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.TopMost = $true                # ★ ブラウザより前
    $form.ShowInTaskbar = $true
    $form.AutoScaleMode = 'Dpi'
    $form.Font = New-Object System.Drawing.Font('Yu Gothic UI', 10)
    $form.ClientSize = New-Object System.Drawing.Size(540, 290)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $msg
    $label.AutoSize = $false
    $label.Dock = 'Fill'
    $label.Padding = New-Object System.Windows.Forms.Padding(18, 18, 18, 8)
    $label.TextAlign = 'TopLeft'
    $label.UseMnemonic = $false   # & を素直に表示
    $form.Controls.Add($label)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Bottom'
    $panel.Height = 48
    $form.Controls.Add($panel)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = 'OK'
    $btn.DialogResult = 'OK'
    $btn.Size = New-Object System.Drawing.Size(96, 30)
    $btn.Anchor = 'Right'
    $btn.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 96 - 18), 9)
    $panel.Controls.Add($btn)
    $form.AcceptButton = $btn
    $form.CancelButton = $btn

    # フォーカスを最前面に持ってくる (一部環境で TopMost だけだと弱いため)
    $form.Add_Shown({
        $form.Activate()
        $form.BringToFront()
        $btn.Focus() | Out-Null
    })

    [void]$form.ShowDialog()
    $form.Dispose()
} catch {
    # WinForms が使えない環境では console 出力で代替
    Write-Host ''
    Write-Host '[tadori-start] 準備完了。ブラウザの SharePoint ページで「Tadori」ブックマークレットを押してください。' -ForegroundColor Green
}
