# ============================================================================
# Tadori AI relay  (Pure PowerShell, no Python)
# ============================================================================
#
# ブラウザ (Tadori bookmarklet) から社内 AI ゲートウェイ / Azure OpenAI 互換
# エンドポイントを、オンプレ プロキシ経由で呼び出すための、ローカルで動く
# 小さな HTTP リレー。Spira の spira-ai-relay.ps1 から AI 中継部分のみを
# 流用（Outlook COM 操作は Tadori では不要なので除外）。
#
# なぜ必要か
# ----------
# ブラウザの fetch() は環境変数 HTTP_PROXY / HTTPS_PROXY を読まないし、
# Fetch API の仕様でプロキシを per-request で指定する方法もない。よって
# bookmarklet から「社内 AI ゲートウェイには必ず社内プロキシ経由で行く」
# というルーティングを直接表現できない。
#
#   Tadori (browser) --HTTP--> http://127.0.0.1:18080 --HTTPS via proxy--> gateway
#
# ブラウザは loopback には到達できる (プロキシ判定はループバックを除外)。
# PS 側は HttpClient で proxy を指定できる。
#
# 使い方
# ------
#   PS> Copy-Item tadori-ai-relay.env.example tadori-ai-relay.env
#   PS> notepad tadori-ai-relay.env       # 値を編集
#   PS> .\tadori-ai-relay.ps1
#
#   # または引数で個別に上書き:
#   PS> .\tadori-ai-relay.ps1 -Target 'https://...' -Proxy 'http://...:8080'
#
# 設定の優先順位:
#   1. コマンドライン引数 (-Target / -Proxy / -Port)
#   2. プロセス環境変数 (TADORI_AI_TARGET 等)
#   3. tadori-ai-relay.env ファイル (同じフォルダ)
#   4. デフォルト値 (port = 18080)
#
# 埋め込み呼び出し例 (Tadori runtime / PoC からはこの localhost を叩く):
#   POST http://localhost:18080/openai/deployments/text-embedding-3-small/embeddings?api-version=2024-02-01
#   ヘッダ: api-key: <サブスクリプションキー>
#
# 必要環境: Windows PowerShell 5.1 以上 / PowerShell 7+。Python 不要。
# 注: HttpListener が 127.0.0.1 で listen するので管理者権限は不要。
# ============================================================================

[CmdletBinding()]
param(
    [string]$Target,
    [string]$Proxy,
    [int]$Port,
    [switch]$NoProxy,
    # 社内ゲートウェイが自己署名証明書の場合 (要セキュリティ承認)
    [switch]$SkipCertCheck,
    # 環境設定ファイルのパス (既定: スクリプトと同じフォルダの tadori-ai-relay.env)
    [string]$EnvFile,
    # 開発者モードでローカル配信するバンドルの dist フォルダ (既定: ../dist)
    [string]$BundleDir
)

$ErrorActionPreference = 'Stop'

# 開発者モードのローカル配信フォルダ (tadori.bundle.js / version.txt をここから配る)。
$script:BundleDir = if ($BundleDir) { $BundleDir }
    elseif ($env:TADORI_BUNDLE_DIR) { $env:TADORI_BUNDLE_DIR }
    else { Join-Path $PSScriptRoot '..\dist' }

# ─── Load .env file ─────────────────────────────────────────────────────────
# 同じフォルダの `tadori-ai-relay.env` を読み、まだ設定されていない
# `$env:TADORI_AI_*` だけセットする。引数 / 既存 env を優先 (上書きしない)。
# `.env` 書式は KEY=VALUE。`#` 始まりと空行は無視。前後クォートは剥がす。

function Import-EnvFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $lines = Get-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop
    } catch {
        Write-Warning ".env ファイルを読めませんでした: $Path ($($_.Exception.Message))"
        return $false
    }
    foreach ($raw in $lines) {
        $line = $raw.Trim()
        if (-not $line) { continue }
        if ($line.StartsWith('#')) { continue }
        $eq = $line.IndexOf('=')
        if ($eq -lt 1) { continue }
        $key = $line.Substring(0, $eq).Trim()
        $val = $line.Substring($eq + 1).Trim()
        if ($val -notmatch '^["'']') {
            $hashIdx = $val.IndexOf(' #')
            if ($hashIdx -ge 0) { $val = $val.Substring(0, $hashIdx).TrimEnd() }
        }
        if (($val.StartsWith('"') -and $val.EndsWith('"')) -or
            ($val.StartsWith("'") -and $val.EndsWith("'"))) {
            $val = $val.Substring(1, $val.Length - 2)
        }
        if (-not [Environment]::GetEnvironmentVariable($key)) {
            [Environment]::SetEnvironmentVariable($key, $val)
        }
    }
    return $true
}

if (-not $EnvFile) {
    $EnvFile = Join-Path $PSScriptRoot 'tadori-ai-relay.env'
}
$loaded = Import-EnvFile -Path $EnvFile
if ($loaded) {
    Write-Host "[config] loaded: $EnvFile" -ForegroundColor DarkGray
}

# 引数 → 環境変数 → デフォルト の順で確定
if (-not $Target) { $Target = $env:TADORI_AI_TARGET }
if (-not $Proxy)  { $Proxy  = $env:TADORI_AI_PROXY }
if (-not $Port)   {
    $Port = if ($env:TADORI_AI_PORT) { [int]$env:TADORI_AI_PORT } else { 18080 }
}
if (-not $SkipCertCheck -and $env:TADORI_AI_SKIP_CERT_CHECK -eq '1') {
    $SkipCertCheck = [switch]$true
}

# ─── Pre-flight checks ──────────────────────────────────────────────────────

if (-not $Target) {
    Write-Host 'エラー: AI gateway URL (-Target / TADORI_AI_TARGET) が未指定です。' -ForegroundColor Red
    Write-Host '  Tadori relay は AI 中継専用です。tadori-ai-relay.env に TADORI_AI_TARGET を設定してください。'
    exit 1
}

if (-not $NoProxy -and -not $Proxy) {
    Write-Host '警告: プロキシが未指定です。直接接続を試みます (社内環境では失敗する可能性が高いです)。' -ForegroundColor Yellow
}

# ─── HttpClient setup ───────────────────────────────────────────────────────

Add-Type -AssemblyName System.Net.Http | Out-Null

$handler = New-Object System.Net.Http.HttpClientHandler
$handler.AllowAutoRedirect = $true
$handler.AutomaticDecompression = [System.Net.DecompressionMethods]::None  # SSE のため未解凍で流す

if (-not $NoProxy -and $Proxy) {
    $handler.Proxy = New-Object System.Net.WebProxy($Proxy, $true)  # bypassOnLocal=true
    $handler.UseProxy = $true
}
else {
    $handler.UseProxy = $false
}

if ($SkipCertCheck) {
    # ⚠ 本番運用では不可。検証用のみ。
    $handler.ServerCertificateCustomValidationCallback =
        [System.Net.Http.HttpClientHandler]::DangerousAcceptAnyServerCertificateValidator
}

$client = New-Object System.Net.Http.HttpClient($handler)
$client.Timeout = [TimeSpan]::FromMinutes(10)

# ─── Target URL parsing ─────────────────────────────────────────────────────

$Target = $Target.TrimEnd('/')
$targetUri = [Uri]$Target
$targetPath = $targetUri.AbsolutePath
if ($targetPath -eq '/') { $targetPath = '' }

# ─── HttpListener setup ─────────────────────────────────────────────────────

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:$Port/")
try {
    $listener.Start()
}
catch {
    Write-Host ''
    Write-Host "エラー: ポート $Port を listen できませんでした。" -ForegroundColor Red
    Write-Host "  - 別プロセスが同じポートを使っている可能性 (Get-NetTCPConnection -LocalPort $Port)"
    Write-Host '  - または別ポートを -Port 引数で指定してください'
    Write-Host "詳細: $($_.Exception.Message)"
    exit 1
}

$baseUrlShort  = "http://localhost:$Port"
$baseUrlMirror = if ($targetPath) { "$baseUrlShort$targetPath" } else { $baseUrlShort }

Write-Host ('-' * 72)
Write-Host '  Tadori AI relay (PowerShell)'
Write-Host ('-' * 72)
Write-Host "  listen  : http://127.0.0.1:$Port"
Write-Host "  target  : $Target"
Write-Host "  proxy   : $(if ($NoProxy -or -not $Proxy) { '(直接接続)' } else { $Proxy })"
if ($SkipCertCheck) { Write-Host '  SSL 検証スキップ中 (-SkipCertCheck)' -ForegroundColor Yellow }
Write-Host ('-' * 72)
Write-Host 'エンドポイント:'
Write-Host "  GET  $baseUrlShort/tadori/health           (死活確認)"
Write-Host "  GET  $baseUrlShort/tadori/outlook/import   (Outlook からメール読込: ?to=&cc=&since=&until=&max=)"
Write-Host "  GET  $baseUrlShort/tadori/outlook/open     (Internet-Message-Id でメール表示: ?id=)"
Write-Host "  GET  $baseUrlShort/tadori/onenote/hierarchy (OneNote ノートブック/セクション/ページ階層を取得)"
Write-Host "  GET  $baseUrlShort/tadori/onenote/pages     (OneNote 指定ページの本文抽出: ?ids=&since=&max=)"
Write-Host "  GET  $baseUrlShort/tadori/onenote/open      (OneNote 上でページ表示: ?id=)"
Write-Host "  GET  $baseUrlShort/tadori/tadori.bundle.js (開発: ローカル dist のバンドル配信)"
Write-Host "  GET  $baseUrlShort/tadori/version.txt      (開発: ローカル dist のバージョン配信)"
Write-Host "  $baseUrlShort/tadori/bundle-dir            (開発: 配信フォルダの確認 GET / 変更 POST)"
Write-Host "  *    $baseUrlShort/...                      (上記以外は target へ透過 forward)"
Write-Host "  bundle dir: $script:BundleDir"
Write-Host ''
Write-Host 'Tadori runtime / PoC のベース URL に下記を入力:'
Write-Host "  A: $baseUrlShort"
if ($baseUrlShort -ne $baseUrlMirror) {
    Write-Host "  B: $baseUrlMirror    (実 URL のパスを保ったまま localhost に置換)"
}
Write-Host ('-' * 72)
Write-Host 'Ctrl+C で終了' -ForegroundColor DarkGray
Write-Host ''

# ─── CORS helper ────────────────────────────────────────────────────────────

function Add-CorsHeaders {
    param([System.Net.HttpListenerResponse]$Response)
    # bookmarklet は SharePoint オリジンから動くので明示許可。listen は
    # loopback だけなので `*` でも外部から到達できず安全。
    $Response.Headers.Add('Access-Control-Allow-Origin', '*')
    $Response.Headers.Add('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    $Response.Headers.Add(
        'Access-Control-Allow-Headers',
        'Content-Type, api-key, Accept, Authorization, X-Requested-With, x-api-key'
    )
    # Chrome の Private Network Access (PNA): https の公開オリジンから
    # http://127.0.0.1 (private) へのリクエストはこのヘッダが無いと拒否される。
    $Response.Headers.Add('Access-Control-Allow-Private-Network', 'true')
    $Response.Headers.Add('Access-Control-Max-Age', '86400')
}

# ─── helpers ────────────────────────────────────────────────────────────────

function Send-Error {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [int]$Status,
        [string]$Code,
        [string]$Detail
    )
    $payload = (@{ error = @{ code = $Code; detail = $Detail } } | ConvertTo-Json -Compress -Depth 4)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
    try {
        $Response.StatusCode = $Status
        Add-CorsHeaders -Response $Response
        $Response.ContentType = 'application/json; charset=utf-8'
        $Response.ContentLength64 = $bytes.Length
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    }
    catch { }
    finally { try { $Response.OutputStream.Close() } catch { } }
}

function Send-Json {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [int]$Status,
        [object]$Body
    )
    $json  = ($Body | ConvertTo-Json -Compress -Depth 6)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    try {
        $Response.StatusCode    = $Status
        Add-CorsHeaders -Response $Response
        $Response.ContentType   = 'application/json; charset=utf-8'
        $Response.ContentLength64 = $bytes.Length
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } catch { }
    finally { try { $Response.OutputStream.Close() } catch { } }
}

# 静的ファイル配信 (開発者モードのローカルバンドル用)。
function Send-StaticFile {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [string]$FilePath,
        [string]$ContentType
    )
    if (-not (Test-Path -LiteralPath $FilePath)) {
        Send-Json -Response $Response -Status 404 -Body @{ ok = $false; error = @{ code = 'not_found'; detail = "File not found: $FilePath" } }
        return
    }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        Add-CorsHeaders -Response $Response
        $Response.StatusCode = 200
        $Response.ContentType = $ContentType
        $Response.Headers.Add('Cache-Control', 'no-store') # loader は ?v= でバスティングするが念のため
        $Response.ContentLength64 = $bytes.Length
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $Response.OutputStream.Close()
    } catch {
        Send-Json -Response $Response -Status 500 -Body @{ ok = $false; error = @{ code = 'read_error'; detail = $_.Exception.Message } }
    }
}

# ─── Outlook COM (メールインポート) ─────────────────────────────────────────
# 既存の受信済み ML メールを Outlook クライアントから読み出し、To/Cc 条件と
# 受信期間でフィルタして JSON で返す。Spira relay の Outlook COM 流用。
# ※ 読み取り専用 (.Display/.Send は呼ばない)。Windows + Outlook 必須。

function Get-OutlookOrNull {
    try {
        try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
        catch { return (New-Object -ComObject Outlook.Application) }
    } catch { return $null }
}

# Recipient / Sender から SMTP アドレスを取り出す (Exchange の EX 形式を SMTP へ)。
function Get-SmtpFromAddressEntry {
    param($AddressEntry, [string]$Fallback)
    try {
        if ($AddressEntry) {
            try { $ex = $AddressEntry.GetExchangeUser(); if ($ex -and $ex.PrimarySmtpAddress) { return $ex.PrimarySmtpAddress } } catch { }
            try {
                $smtp = $AddressEntry.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001F')
                if ($smtp) { return $smtp }
            } catch { }
        }
    } catch { }
    return $Fallback
}

# MailItem を JSON 化可能なハッシュへ。
function Read-MailItem {
    param($Item)
    $mid = ''
    try { $mid = $Item.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001F') } catch { }
    $from = ''
    try {
        $from = [string]$Item.SenderEmailAddress
        try { $from = Get-SmtpFromAddressEntry -AddressEntry $Item.Sender -Fallback $from } catch { }
    } catch { }
    $toList = @(); $ccList = @()
    try {
        foreach ($r in $Item.Recipients) {
            $smtp = Get-SmtpFromAddressEntry -AddressEntry $r.AddressEntry -Fallback ([string]$r.Address)
            if ($r.Type -eq 2) { $ccList += $smtp } else { $toList += $smtp }  # 1=To, 2=CC, 3=BCC
        }
    } catch { }
    $date = ''
    try { $date = $Item.ReceivedTime.ToString('o') } catch { }
    # HTML メールは HTMLBody をそのまま渡し isHtml=true (表示時にクライアントで
    # サニタイズ + HTML 描画)。それ以外はプレーン .Body。olFormatHTML=2。
    $isHtml = $false
    $body = ''
    try {
        if ([int]$Item.BodyFormat -eq 2) {
            $h = [string]$Item.HTMLBody
            if ($h) { $body = $h; $isHtml = $true } else { $body = [string]$Item.Body }
        } else {
            $body = [string]$Item.Body
        }
    } catch { try { $body = [string]$Item.Body } catch { } }
    $convId = ''
    try { $convId = [string]$Item.ConversationID } catch { }
    return @{
        messageId         = [string]$mid
        internetMessageId = [string]$mid
        conversationId    = [string]$convId
        subject   = [string]$Item.Subject
        from      = [string]$from
        to        = @($toList)
        cc        = @($ccList)
        date      = [string]$date
        body      = $body
        isHtml    = [bool]$isHtml
    }
}

function Invoke-OutlookImport {
    param([System.Net.HttpListenerContext]$Context)
    $request  = $Context.Request
    $response = $Context.Response
    $q = $request.QueryString

    $split = { param($s) if ($s) { ($s -split '[;,]') | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ } } else { @() } }
    $toAddrs = @(& $split $q['to'])
    $ccAddrs = @(& $split $q['cc'])
    $max = 1000
    if ($q['max']) { $tmp = 0; if ([int]::TryParse($q['max'], [ref]$tmp)) { $max = $tmp } }
    # 日付は「その日いっぱい」を含めたいので、since はその日の 0:00、until は
    # 翌日 0:00 とし、上限は < で比較する (同日指定でもその日が丸ごと範囲に入る)。
    $sinceDt = (Get-Date).Date.AddYears(-10)
    $untilDt = (Get-Date).Date.AddDays(1)
    if ($q['since']) { try { $sinceDt = [DateTime]::Parse($q['since']).Date } catch { } }
    if ($q['until']) { try { $untilDt = [DateTime]::Parse($q['until']).Date.AddDays(1) } catch { } }

    $ol = Get-OutlookOrNull
    if (-not $ol) {
        Send-Error -Response $response -Status 503 -Code 'no_outlook' -Detail 'Outlook を起動/接続できませんでした (Windows + Outlook が必要)'
        return
    }

    try {
        $ns = $ol.GetNamespace('MAPI')
        # ReceivedTime は Outlook ロケール依存。MM/dd/yyyy HH:mm が無難。
        $filter = "[ReceivedTime] >= '" + $sinceDt.ToString('MM/dd/yyyy HH:mm') + "' AND [ReceivedTime] < '" + $untilDt.ToString('MM/dd/yyyy HH:mm') + "'"

        $matches = {
            param($mail)
            if ($toAddrs.Count -eq 0 -and $ccAddrs.Count -eq 0) { return $true }
            $to = @(); $cc = @()
            try {
                foreach ($r in $mail.Recipients) {
                    $s = (Get-SmtpFromAddressEntry -AddressEntry $r.AddressEntry -Fallback ([string]$r.Address)).ToLower()
                    if ($r.Type -eq 2) { $cc += $s } else { $to += $s }
                }
            } catch { }
            foreach ($a in $toAddrs) { if ($to -contains $a) { return $true } }
            foreach ($a in $ccAddrs) { if ($cc -contains $a) { return $true } }
            return $false
        }

        # 除外フォルダ (送信済み=5 / 削除済み=3 / 迷惑メール=23) を全ストアぶん収集。
        # EntryID で一致判定し、そのフォルダ自身とサブフォルダを走査対象から外す。
        $skipIds = @{}
        try {
            foreach ($store in $ns.Stores) {
                foreach ($ft in @(3, 5, 23)) {
                    try { $f = $store.GetDefaultFolder($ft); if ($f) { $skipIds[$f.EntryID] = $true } } catch { }
                }
            }
        } catch { }

        $results = New-Object System.Collections.ArrayList
        # 全ストア (メインメールボックス / アーカイブ / オンラインアーカイブ / PST) の
        # ルートから全フォルダを幅優先で走査する。ローカルの「アーカイブ」は受信トレイの
        # 外、オンラインアーカイブは別ストアなので、Inbox 配下だけだと拾えない。
        $queue = New-Object System.Collections.Queue
        foreach ($store in $ns.Folders) { $queue.Enqueue($store) }
        $folders = 0
        while ($queue.Count -gt 0 -and $results.Count -lt $max -and $folders -lt 5000) {
            $folder = $queue.Dequeue(); $folders++
            try { if ($skipIds.ContainsKey($folder.EntryID)) { continue } } catch { }  # 送信済/削除済/迷惑メールは除外 (配下も辿らない)
            try { foreach ($sf in $folder.Folders) { $queue.Enqueue($sf) } } catch { }
            $items = $null
            try { $items = $folder.Items.Restrict($filter) } catch { continue }  # メール以外のフォルダは除外
            foreach ($it in $items) {
                if ($results.Count -ge $max) { break }
                try {
                    if ($it.Class -ne 43) { continue }  # olMail
                    if (& $matches $it) { [void]$results.Add((Read-MailItem -Item $it)) }
                } catch { }
            }
        }

        Write-Host ("[import] matched {0} mails / scanned {1} folders (to=[{2}] cc=[{3}] {4}..{5})" -f $results.Count, $folders, ($toAddrs -join ','), ($ccAddrs -join ','), $sinceDt.ToString('yyyy-MM-dd'), $untilDt.AddDays(-1).ToString('yyyy-MM-dd'))
        Send-Json -Response $response -Status 200 -Body @{ ok = $true; count = $results.Count; mails = @($results) }
    }
    catch {
        Send-Error -Response $response -Status 500 -Code 'outlook_error' -Detail $_.Exception.Message
    }
}

# Internet-Message-Id で Outlook 内のメールを探し、見つかればクライアント上に
# 表示 (インスペクタを開く)。受信トレイ等の全ストア・全サブフォルダを走査。
function Invoke-OutlookOpen {
    param([System.Net.HttpListenerContext]$Context)
    $response = $Context.Response
    $id = $Context.Request.QueryString['id']
    if (-not $id) {
        Send-Error -Response $response -Status 400 -Code 'bad_request' -Detail 'クエリ id (Internet-Message-Id) が必要です'
        return
    }

    $ol = Get-OutlookOrNull
    if (-not $ol) {
        Send-Error -Response $response -Status 503 -Code 'no_outlook' -Detail 'Outlook を起動/接続できませんでした (Windows + Outlook が必要)'
        return
    }

    try {
        $ns = $ol.GetNamespace('MAPI')
        # PR_INTERNET_MESSAGE_ID (0x1035001F) を DASL で等値フィルタ。値内の ' は '' へ。
        $idEsc = $id -replace "'", "''"
        $dasl  = '@SQL="http://schemas.microsoft.com/mapi/proptag/0x1035001F" = ''' + $idEsc + ''''

        $queue = New-Object System.Collections.Queue
        foreach ($store in $ns.Folders) { $queue.Enqueue($store) }

        $found = $null
        $folders = 0
        while ($queue.Count -gt 0 -and -not $found -and $folders -lt 2000) {
            $folder = $queue.Dequeue(); $folders++
            try { foreach ($sf in $folder.Folders) { $queue.Enqueue($sf) } } catch { }
            try {
                $items = $folder.Items.Restrict($dasl)
                foreach ($it in $items) {
                    try { if ($it.Class -eq 43) { $found = $it; break } } catch { }  # olMail
                }
            } catch { continue }
        }

        if ($found) {
            try { $found.Display() } catch { }                    # インスペクタを開く
            # 最前面化: メールのインスペクタを Activate、Outlook 本体も Activate。
            try { $found.GetInspector.Activate() } catch { }
            try { $ol.ActiveExplorer().Activate() } catch { }
            # それでも他ウィンドウに隠れるケース用に Win32 SetForegroundWindow を保険で叩く。
            try {
                if (-not ('Tadori.Native' -as [type])) {
                    Add-Type -Namespace Tadori -Name Native -MemberDefinition @'
                        [System.Runtime.InteropServices.DllImport("user32.dll")]
                        public static extern bool SetForegroundWindow(System.IntPtr hWnd);
                        [System.Runtime.InteropServices.DllImport("user32.dll")]
                        public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
'@
                }
                $hwnd = [System.IntPtr]$found.GetInspector.WindowHandle()
                if ($hwnd -ne [System.IntPtr]::Zero) {
                    [void][Tadori.Native]::ShowWindow($hwnd, 9)   # SW_RESTORE = 9
                    [void][Tadori.Native]::SetForegroundWindow($hwnd)
                }
            } catch { }
            Write-Host ("[open] displayed mail id={0}" -f $id)
            Send-Json -Response $response -Status 200 -Body @{ ok = $true; found = $true }
        } else {
            Write-Host ("[open] not found id={0} (scanned {1} folders)" -f $id, $folders)
            Send-Json -Response $response -Status 200 -Body @{ ok = $true; found = $false }
        }
    }
    catch {
        Send-Error -Response $response -Status 500 -Code 'outlook_error' -Detail $_.Exception.Message
    }
}

# ─── OneNote COM (ノートブック/セクション/ページ取り込み) ───────────────────
# 階層列挙 → 選択ページ群の本文抽出 → 既定アプリでページを開く。
# 必要環境: Windows + OneNote デスクトップ。Outlook と同じく COM 経由。

function Get-OneNoteOrNull {
    try {
        try { return [Runtime.InteropServices.Marshal]::GetActiveObject('OneNote.Application') }
        catch { return (New-Object -ComObject OneNote.Application) }
    } catch { return $null }
}

# OneNote XML から階層 JSON (ブラウザでツリー表示する用) を構築。
function Invoke-OneNoteHierarchy {
    param([System.Net.HttpListenerContext]$Context)
    $response = $Context.Response
    $one = Get-OneNoteOrNull
    if (-not $one) {
        Send-Error -Response $response -Status 503 -Code 'no_onenote' -Detail 'OneNote を起動/接続できませんでした (Windows + OneNote が必要)'
        return
    }
    try {
        $xml = ''
        # 4 = hsPages: ノートブック → セクション → ページ まで取得。
        $one.GetHierarchy('', 4, [ref]$xml)
        [xml]$doc = $xml
        $notebooks = New-Object System.Collections.ArrayList
        foreach ($nb in $doc.SelectNodes('//*[local-name()="Notebook"]')) {
            $sectionsList = New-Object System.Collections.ArrayList
            foreach ($sec in $nb.SelectNodes('descendant::*[local-name()="Section"]')) {
                $pagesList = New-Object System.Collections.ArrayList
                foreach ($pg in $sec.SelectNodes('*[local-name()="Page"]')) {
                    $lvl = if ($pg.pageLevel) { [int]$pg.pageLevel } else { 1 }
                    [void]$pagesList.Add(@{
                        id = [string]$pg.ID
                        name = [string]$pg.name
                        lastModified = [string]$pg.lastModifiedTime
                        level = $lvl
                    })
                }
                [void]$sectionsList.Add(@{
                    id = [string]$sec.ID
                    name = [string]$sec.name
                    pages = @($pagesList)
                })
            }
            [void]$notebooks.Add(@{
                id = [string]$nb.ID
                name = [string]$nb.name
                sections = @($sectionsList)
            })
        }
        Send-Json -Response $response -Status 200 -Body @{ ok = $true; notebooks = @($notebooks) }
    }
    catch {
        Send-Error -Response $response -Status 500 -Code 'onenote_error' -Detail $_.Exception.Message
    }
}

# クエリ:
#   ?ids=<pageId>;<pageId>;...        対象ページ ID (URL エンコード推奨)
#   ?since=YYYY-MM-DD                 lastModified がこれ以降のもののみ
#   ?max=200                          上限件数
# 戻り値: ページ配列 { pageId, notebook, section, title, lastModified, body }
function Invoke-OneNotePages {
    param([System.Net.HttpListenerContext]$Context)
    $response = $Context.Response
    $q = $Context.Request.QueryString
    $ids = @()
    if ($q['ids']) { $ids = ($q['ids'] -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ } }
    $max = 500
    if ($q['max']) { $tmp = 0; if ([int]::TryParse($q['max'], [ref]$tmp)) { $max = $tmp } }
    $since = $null
    if ($q['since']) { try { $since = [DateTime]::Parse($q['since']) } catch { } }

    $one = Get-OneNoteOrNull
    if (-not $one) {
        Send-Error -Response $response -Status 503 -Code 'no_onenote' -Detail 'OneNote を起動/接続できませんでした (Windows + OneNote が必要)'
        return
    }
    try {
        $hierXml = ''
        $one.GetHierarchy('', 4, [ref]$hierXml)
        [xml]$hier = $hierXml
        $allPages = $hier.SelectNodes('//*[local-name()="Page"]')

        # 対象ページを絞る (ids 指定があればそれだけ。無ければ全部)。
        $targets = New-Object System.Collections.ArrayList
        foreach ($pg in $allPages) {
            if ($ids.Count -gt 0 -and ($ids -notcontains [string]$pg.ID)) { continue }
            if ($since) {
                $lm = $null
                try { $lm = [DateTime]::Parse([string]$pg.lastModifiedTime) } catch { }
                if ($lm -and $lm -lt $since) { continue }
            }
            [void]$targets.Add($pg)
            if ($targets.Count -ge $max) { break }
        }

        $results = New-Object System.Collections.ArrayList
        foreach ($pg in $targets) {
            $section = $pg.ParentNode
            $notebook = $section.ParentNode
            while ($notebook -and ($notebook.LocalName -ne 'Notebook')) { $notebook = $notebook.ParentNode }
            $content = ''
            try { $one.GetPageContent([string]$pg.ID, [ref]$content) } catch { continue }
            $text = ''
            try {
                [xml]$cdoc = $content
                $tNodes = $cdoc.SelectNodes('//*[local-name()="T"]')
                $parts = @()
                foreach ($t in $tNodes) {
                    $raw = [string]$t.InnerText
                    if ($raw) { $parts += $raw }
                }
                # 文字参照や軽い HTML タグを落として可読テキストに。
                $text = ($parts -join "`n") -replace '&nbsp;', ' ' -replace '<[^>]+>', ''
            } catch { }
            [void]$results.Add(@{
                pageId       = [string]$pg.ID
                title        = [string]$pg.name
                lastModified = [string]$pg.lastModifiedTime
                notebook     = if ($notebook) { [string]$notebook.name } else { '' }
                section      = [string]$section.name
                body         = $text
            })
        }

        $sinceStr = if ($since) { $since.ToString('yyyy-MM-dd') } else { '' }
        Write-Host ("[onenote] returned {0} pages (ids:{1} since:{2})" -f $results.Count, $ids.Count, $sinceStr)
        Send-Json -Response $response -Status 200 -Body @{ ok = $true; count = $results.Count; pages = @($results) }
    }
    catch {
        Send-Error -Response $response -Status 500 -Code 'onenote_error' -Detail $_.Exception.Message
    }
}

# 指定 pageId のページを OneNote 上で表示。
function Invoke-OneNoteOpen {
    param([System.Net.HttpListenerContext]$Context)
    $response = $Context.Response
    $id = $Context.Request.QueryString['id']
    if (-not $id) {
        Send-Error -Response $response -Status 400 -Code 'bad_request' -Detail 'クエリ id (pageId) が必要です'
        return
    }
    $one = Get-OneNoteOrNull
    if (-not $one) {
        Send-Error -Response $response -Status 503 -Code 'no_onenote' -Detail 'OneNote を起動/接続できませんでした'
        return
    }
    try {
        $one.NavigateTo($id, $null, $false)
        Write-Host ("[onenote] navigated to page id={0}" -f $id)
        Send-Json -Response $response -Status 200 -Body @{ ok = $true }
    }
    catch {
        Send-Error -Response $response -Status 500 -Code 'onenote_error' -Detail $_.Exception.Message
    }
}

# ─── Request handler ────────────────────────────────────────────────────────

function Invoke-RelayRequest {
    param([System.Net.HttpListenerContext]$Context)

    $request  = $Context.Request
    $response = $Context.Response
    $method   = $request.HttpMethod.ToUpper()
    $ts       = (Get-Date).ToString('HH:mm:ss')
    Write-Host ("[{0}] {1} {2}" -f $ts, $method, $request.Url.PathAndQuery)

    # ── CORS preflight ──
    if ($method -eq 'OPTIONS') {
        $response.StatusCode = 204
        Add-CorsHeaders -Response $response
        $response.OutputStream.Close()
        return
    }

    # ── ローカル機能: 死活確認 ──
    $path = $request.Url.AbsolutePath
    if ($path -eq '/tadori/health') {
        Send-Json -Response $response -Status 200 -Body @{ ok = $true; relay = 'tadori-ai-relay'; version = 1 }
        return
    }

    # ── ローカル機能: Outlook からのメールインポート (読み取り専用) ──
    if ($path -eq '/tadori/outlook/import') {
        Invoke-OutlookImport -Context $Context
        return
    }

    # ── ローカル機能: Internet-Message-Id でメールを Outlook 上に表示 ──
    if ($path -eq '/tadori/outlook/open') {
        Invoke-OutlookOpen -Context $Context
        return
    }

    # ── ローカル機能: OneNote 取り込み (階層取得 / ページ抽出 / 開く) ──
    if ($path -eq '/tadori/onenote/hierarchy') { Invoke-OneNoteHierarchy -Context $Context; return }
    if ($path -eq '/tadori/onenote/pages')     { Invoke-OneNotePages     -Context $Context; return }
    if ($path -eq '/tadori/onenote/open')      { Invoke-OneNoteOpen      -Context $Context; return }

    # ── 開発者モード: ローカル dist のバンドル配信 (loader が読む) ──
    if ($path -eq '/tadori/tadori.bundle.js') {
        Send-StaticFile -Response $response -FilePath (Join-Path $script:BundleDir 'tadori.bundle.js') -ContentType 'application/javascript; charset=utf-8'
        return
    }
    if ($path -eq '/tadori/version.txt') {
        Send-StaticFile -Response $response -FilePath (Join-Path $script:BundleDir 'version.txt') -ContentType 'text/plain; charset=utf-8'
        return
    }
    if ($path -eq '/tadori/bundle-dir') {
        if ($method -eq 'GET') {
            $exists = Test-Path -LiteralPath $script:BundleDir
            $hasBundle = Test-Path -LiteralPath (Join-Path $script:BundleDir 'tadori.bundle.js')
            Send-Json -Response $response -Status 200 -Body @{ ok = $true; dir = "$script:BundleDir"; exists = [bool]$exists; hasBundle = [bool]$hasBundle }
            return
        }
        if ($method -eq 'POST') {
            try {
                $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
                $raw = $reader.ReadToEnd(); $reader.Dispose() | Out-Null
                $p = if ($raw) { $raw | ConvertFrom-Json } else { $null }
                $dir = "$($p.dir)".Trim()
                if ($dir) { $script:BundleDir = $dir; Write-Host "[bundle] dir set to: $dir" }
                $exists = Test-Path -LiteralPath $script:BundleDir
                $hasBundle = Test-Path -LiteralPath (Join-Path $script:BundleDir 'tadori.bundle.js')
                Send-Json -Response $response -Status 200 -Body @{ ok = $true; dir = "$script:BundleDir"; exists = [bool]$exists; hasBundle = [bool]$hasBundle }
            } catch {
                Send-Json -Response $response -Status 400 -Body @{ ok = $false; error = @{ code = 'bad_request'; detail = $_.Exception.Message } }
            }
            return
        }
    }

    # ── Compose upstream URL ──
    # bookmarklet 側で baseUrl に target のパスを含めても含めなくても OK に
    # するため、incoming path 先頭が targetPath と一致したら剥がす。
    $incoming = $request.Url.PathAndQuery
    $rel = $incoming
    if ($targetPath -and $rel.StartsWith($targetPath)) {
        $rel = $rel.Substring($targetPath.Length)
        if (-not $rel) { $rel = '/' }
    }
    $upstreamUrl = $Target + $rel

    # ── Build HttpRequestMessage ──
    $httpMethod = New-Object System.Net.Http.HttpMethod($method)
    $msg = New-Object System.Net.Http.HttpRequestMessage($httpMethod, $upstreamUrl)

    if ($request.HasEntityBody) {
        $ms = New-Object System.IO.MemoryStream
        $request.InputStream.CopyTo($ms)
        $bodyBytes = $ms.ToArray()
        $ms.Dispose()
        $content = New-Object System.Net.Http.ByteArrayContent($bodyBytes, 0, $bodyBytes.Length)
        if ($request.ContentType) {
            try {
                $content.Headers.ContentType =
                    [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse($request.ContentType)
            }
            catch { }
        }
        $msg.Content = $content
    }

    # header forwarding — 認証/コンテンツ系のみ転送 (Host/Connection 等は除外)。
    $forwardKeys = @(
        'api-key', 'x-api-key',
        'accept', 'authorization'
    )
    foreach ($name in $request.Headers.AllKeys) {
        if ($forwardKeys -contains $name.ToLower()) {
            $val = $request.Headers[$name]
            $msg.Headers.TryAddWithoutValidation($name, $val) | Out-Null
        }
    }

    # ── Send upstream (stream-aware) ──
    try {
        $task = $client.SendAsync(
            $msg,
            [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead
        )
        $upstream = $task.GetAwaiter().GetResult()

        $response.StatusCode = [int]$upstream.StatusCode
        Add-CorsHeaders -Response $response
        $ct = $null
        if ($upstream.Content -and $upstream.Content.Headers.ContentType) {
            $ct = $upstream.Content.Headers.ContentType.ToString()
        }
        if ($ct) { $response.ContentType = $ct }
        if ($upstream.Headers.CacheControl) {
            $response.Headers.Add('Cache-Control', $upstream.Headers.CacheControl.ToString())
        }
        $sse = ($ct -and ($ct -like 'text/event-stream*'))
        if ($sse) {
            $response.SendChunked = $true
        }
        elseif ($upstream.Content.Headers.ContentLength) {
            $response.ContentLength64 = $upstream.Content.Headers.ContentLength
        }
        else {
            $response.SendChunked = $true
        }

        $upstreamStream = $upstream.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        $buffer = New-Object byte[] 1024
        while ($true) {
            $n = $upstreamStream.Read($buffer, 0, $buffer.Length)
            if ($n -le 0) { break }
            try {
                $response.OutputStream.Write($buffer, 0, $n)
                $response.OutputStream.Flush()
            }
            catch { break }   # ブラウザがキャンセル → サイレントに終了
        }
        $upstreamStream.Dispose()
        $upstream.Dispose()
    }
    catch [System.Net.Http.HttpRequestException] {
        $detail = $_.Exception.Message
        $inner = $_.Exception.InnerException
        if ($inner) { $detail += " — $($inner.Message)" }
        Send-Error -Response $response -Status 502 -Code 'upstream_error' -Detail $detail
    }
    catch {
        Send-Error -Response $response -Status 500 -Code 'relay_failed' -Detail $_.Exception.Message
    }
    finally {
        try { $response.OutputStream.Close() } catch { }
        try { $msg.Dispose() } catch { }
    }
}

# ─── Main loop ──────────────────────────────────────────────────────────────

[Console]::TreatControlCAsInput = $false

try {
    while ($listener.IsListening) {
        $ctx = $null
        try {
            $ctx = $listener.GetContext()
        }
        catch [System.Net.HttpListenerException] {
            break
        }
        if ($ctx) {
            try { Invoke-RelayRequest -Context $ctx }
            catch { Write-Warning "request handler error: $($_.Exception.Message)" }
        }
    }
}
finally {
    Write-Host ''
    Write-Host '[shutdown] stopping listener...' -ForegroundColor DarkGray
    try { $listener.Stop() } catch { }
    try { $listener.Close() } catch { }
    try { $client.Dispose() } catch { }
}
