# Foundry Local 速度ベンチマーク (CPU 実行想定)。
#
# 目的: Tadori で chatProvider='foundry-local' を選択肢にする前に、業務 PC で
# 実際にどれくらいの速度が出るかを事前検証する。
#
# 計測内容:
#   - TTFT (Time To First Token, 最初の出力が返るまでのレイテンシ)
#   - 出力トークン数 (≒ 出力文字数 / 2.5 で近似 ※日本語多めの想定)
#   - tokens/sec (= 出力速度の体感)
#
# 3 シナリオを実行:
#   A. クエリルータ風 (JSON 短文出力 ~50 tok) — 頻発する軽量タスク
#   B. 短い要約 (~200 tok)                    — 普段のサマリ用途
#   C. 詳細回答 (~500 tok)                    — RAG の最終回答用途
#
# 使い方:
#   .\test-foundry-local.ps1                              # 既定モデルで 1 回ずつ
#   .\test-foundry-local.ps1 -Runs 3                      # 各シナリオ 3 回 (中央値報告)
#   .\test-foundry-local.ps1 -Model phi-4-mini-instruct-generic-cpu
#   .\test-foundry-local.ps1 -BaseUrl http://localhost:5273/v1
#
# 前提: Foundry Local が起動済 (例: `foundry start` で localhost:5273 が listen)
#       既定モデル URL は MS の Foundry Local 公開仕様準拠。実環境で差があれば
#       -BaseUrl で上書きしてください。

[CmdletBinding()]
param(
  [string]$BaseUrl = 'http://localhost:5273/v1',
  [string]$Model   = '',
  [int]   $Runs    = 1
)

$ErrorActionPreference = 'Stop'

# ─── ユーティリティ ────────────────────────────────────────────────────────

function Get-FoundryModels {
    try {
        $res = Invoke-RestMethod -Uri "$BaseUrl/models" -Method GET -TimeoutSec 8
        if ($null -eq $res.data) { return @() }
        return @($res.data | Select-Object -ExpandProperty id)
    } catch {
        Write-Host "❌ /models に接続できません: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host '   → Foundry Local が起動しているか確認 (foundry start)' -ForegroundColor Yellow
        Write-Host "   → 既定エンドポイントが違う場合は -BaseUrl http://localhost:<port>/v1 を指定" -ForegroundColor Yellow
        return $null
    }
}

# 1 回の chat completion を SSE ストリーミングで実行し、計測値を返す。
# HttpWebRequest を使うのは Invoke-RestMethod が SSE chunked を扱いにくいため。
function Invoke-ChatStream {
    param(
        [Parameter(Mandatory)] [string] $Prompt,
        [Parameter(Mandatory)] [string] $ModelName,
        [int] $MaxTokens = 500
    )

    $body = @{
        model = $ModelName
        messages = @(
            @{ role = 'system'; content = 'You are a concise assistant. 日本語で簡潔に答えてください。' },
            @{ role = 'user';   content = $Prompt }
        )
        stream = $true
        max_tokens = $MaxTokens
        # 一部の OpenAI 互換実装は include_usage に対応 (Foundry Local も対応の想定)。
        # 未対応なら無視されるだけなので付けておく。
        stream_options = @{ include_usage = $true }
    } | ConvertTo-Json -Depth 6

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $ttftMs = 0
    $output = New-Object System.Text.StringBuilder
    $chunkCount = 0
    $reportedTokens = 0 # usage から取れた場合の正確な値

    $req = [System.Net.HttpWebRequest]::Create("$BaseUrl/chat/completions")
    $req.Method = 'POST'
    $req.ContentType = 'application/json'
    $req.Accept = 'text/event-stream'
    $req.Timeout = 600000          # 10 分 (CPU だと長い回答で結構かかる)
    $req.ReadWriteTimeout = 600000

    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($body)
    $req.ContentLength = $bodyBytes.Length
    $rs = $req.GetRequestStream()
    $rs.Write($bodyBytes, 0, $bodyBytes.Length)
    $rs.Close()

    try {
        $res = $req.GetResponse()
        $stream = $res.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)

        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if (-not $line.StartsWith('data:')) { continue }
            $data = $line.Substring(5).Trim()
            if ($data -eq '[DONE]') { break }
            try {
                $chunk = $data | ConvertFrom-Json
            } catch { continue }

            # 出力 delta
            if ($chunk.choices -and $chunk.choices.Count -gt 0) {
                $delta = $chunk.choices[0].delta.content
                if ($delta) {
                    if ($ttftMs -eq 0) { $ttftMs = $sw.ElapsedMilliseconds }
                    [void]$output.Append($delta)
                    $chunkCount++
                }
            }
            # usage (最終チャンクで来る場合あり)
            if ($chunk.usage -and $chunk.usage.completion_tokens) {
                $reportedTokens = [int]$chunk.usage.completion_tokens
            }
        }
        $reader.Close()
        $res.Close()
    } catch {
        Write-Host "  ⚠ リクエスト失敗: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }

    $sw.Stop()
    $totalMs = $sw.ElapsedMilliseconds
    $outputText = $output.ToString()

    # トークン数: usage が取れていればそれを、無ければ文字数から近似
    # (日本語混じりだと「1 トークン ≒ 2.5 文字」が経験則的に妥当)
    $approxTokens = if ($reportedTokens -gt 0) { $reportedTokens } else { [Math]::Round($outputText.Length / 2.5) }
    $tokensPerSec = if ($totalMs -gt 0) { [Math]::Round($approxTokens / ($totalMs / 1000.0), 1) } else { 0 }

    return [pscustomobject]@{
        TtftMs        = $ttftMs
        TotalMs       = $totalMs
        OutputLen     = $outputText.Length
        ApproxTokens  = $approxTokens
        ExactTokens   = $reportedTokens
        TokensPerSec  = $tokensPerSec
        ChunkCount    = $chunkCount
        OutputPreview = $outputText.Substring(0, [Math]::Min(80, $outputText.Length))
    }
}

# 中央値 (PowerShell の Statistics 系が貧弱なので手書き)。
# @(...) で必ず配列化しないと要素 1 件の時にスカラーになって添字アクセスが死ぬ。
function Get-Median {
    param([Parameter(Mandatory)] [double[]] $Values)
    if ($Values.Count -eq 0) { return 0 }
    $sorted = @($Values | Sort-Object)
    $mid = [Math]::Floor($sorted.Count / 2)
    if ($sorted.Count % 2 -eq 1) { return $sorted[$mid] }
    return ($sorted[$mid - 1] + $sorted[$mid]) / 2
}

# ─── テストシナリオ ────────────────────────────────────────────────────────

$scenarios = @(
    [pscustomobject]@{
        Name      = 'A. クエリルータ風 (JSON 短文出力)'
        Detail    = '想定: ~50 tok。Tadori の queryRouter が毎質問で走る軽量タスク'
        Prompt    = 'ユーザの質問を分類して JSON 1 行だけで返してください。質問: 「Direct Connect の冗長化要件は?」 出力形式: {"mode":"keyword|semantic|mixed","keywords":["..."],"vectorQuery":"..."}'
        MaxTokens = 120
    },
    [pscustomobject]@{
        Name      = 'B. 短い要約 (3 行に要約)'
        Detail    = '想定: ~200 tok。短いメール要約や OneNote ページの 1 行要約'
        Prompt    = "次のメール本文を 3 行で要約してください。" + `
                    "「来週月曜 14:00 から会議室 A で打ち合わせを行います。" + `
                    "アジェンダは Phase 2 計画レビューです。出席者は田中、山田、佐藤、平田。" + `
                    "資料は事前に共有フォルダに置きます。質問は事前に Slack で募ります。」"
        MaxTokens = 350
    },
    [pscustomobject]@{
        Name      = 'C. 詳細回答 (500 tok 級)'
        Detail    = '想定: ~500 tok。RAG の最終回答想定 (体感の遅さが一番見える)'
        Prompt    = 'AWS Direct Connect で冗長化を行う際に考慮すべき設計ポイントを 5 つ、それぞれ 2-3 文で説明してください。'
        MaxTokens = 900
    }
)

# ─── Main ────────────────────────────────────────────────────────────────

Write-Host ''
Write-Host '═══ Foundry Local ベンチマーク ═══' -ForegroundColor Cyan
Write-Host ("Endpoint : {0}" -f $BaseUrl)
Write-Host ("Runs/scn : {0}" -f $Runs)

# モデル選択
if (-not $Model) {
    Write-Host '/v1/models からモデル一覧取得中... ' -NoNewline
    $models = Get-FoundryModels
    if ($null -eq $models) { exit 1 }
    Write-Host ("({0} 件)" -f $models.Count)
    if ($models.Count -eq 0) {
        Write-Host '❌ 利用可能なモデルが 1 つもありません。foundry pull <model> でモデルを取り込んでください。' -ForegroundColor Red
        exit 1
    }
    Write-Host '利用可能モデル:'
    $models | ForEach-Object { Write-Host "  - $_" }
    $Model = $models[0]
    Write-Host ("使用モデル: {0} (先頭を自動採用)" -f $Model) -ForegroundColor Yellow
} else {
    Write-Host ("使用モデル: {0}" -f $Model) -ForegroundColor Yellow
}

# シナリオごとに測定
$allResults = @{}
foreach ($s in $scenarios) {
    Write-Host ''
    Write-Host ('▶ ' + $s.Name) -ForegroundColor Cyan
    Write-Host ('  ' + $s.Detail) -ForegroundColor DarkGray
    Write-Host ('  prompt: ' + $s.Prompt.Substring(0, [Math]::Min(70, $s.Prompt.Length)) + '...')

    $rows = @()
    for ($i = 1; $i -le $Runs; $i++) {
        if ($Runs -gt 1) { Write-Host ("  Run {0}/{1}: " -f $i, $Runs) -NoNewline }
        $r = Invoke-ChatStream -Prompt $s.Prompt -ModelName $Model -MaxTokens $s.MaxTokens
        if ($null -eq $r) { continue }
        $rows += $r
        if ($Runs -gt 1) {
            Write-Host (" TTFT={0}ms total={1}ms ~{2}tok @ {3}tok/s" -f `
                $r.TtftMs, $r.TotalMs, $r.ApproxTokens, $r.TokensPerSec)
        }
    }
    if ($rows.Count -eq 0) { continue }

    $ttftMed   = Get-Median ($rows | ForEach-Object { [double]$_.TtftMs })
    $totalMed  = Get-Median ($rows | ForEach-Object { [double]$_.TotalMs })
    $tpsMed    = Get-Median ($rows | ForEach-Object { [double]$_.TokensPerSec })
    $tokensMed = Get-Median ($rows | ForEach-Object { [double]$_.ApproxTokens })

    Write-Host ('  ━ 中央値:') -ForegroundColor White
    Write-Host ("     TTFT       : {0,7:N0} ms" -f $ttftMed)
    Write-Host ("     Total      : {0,7:N0} ms" -f $totalMed)
    Write-Host ("     出力 tok   : {0,7:N0} 個 (近似)" -f $tokensMed)
    Write-Host ("     スループット: {0,7:N1} tok/s" -f $tpsMed)
    Write-Host ('     出力例     : ' + $rows[0].OutputPreview)

    $allResults[$s.Name] = [pscustomobject]@{
        TtftMs       = $ttftMed
        TotalMs      = $totalMed
        TokensMedian = $tokensMed
        TokensPerSec = $tpsMed
    }
}

# ─── 体感評価 ────────────────────────────────────────────────────────────

Write-Host ''
Write-Host '═══ 体感評価 ═══' -ForegroundColor Cyan
# シナリオ C (詳細回答) の tok/s を主指標に
if ($allResults.Count -gt 0) {
    $cKey = ($allResults.Keys | Where-Object { $_ -like 'C.*' } | Select-Object -First 1)
    if ($cKey) {
        $tps = $allResults[$cKey].TokensPerSec
        $verdict = if     ($tps -ge 15) { '🟢 快適 (Azure OpenAI 並み)' }
                   elseif ($tps -ge 10) { '🟡 まずまず (体感少し遅め、実用 OK)' }
                   elseif ($tps -ge 5)  { '🟠 厳しい (短いタスクなら OK、長文回答は待ちが目立つ)' }
                   else                  { '🔴 待ちが目立つ (Tadori 普段使いには厳しい、機密タスク用途のみ推奨)' }
        Write-Host ('  詳細回答スループット: {0} tok/s → {1}' -f $tps, $verdict)
    }
}

Write-Host ''
Write-Host '※ 体感の参考目安' -ForegroundColor DarkGray
Write-Host '   - Azure OpenAI (gpt-4o-mini) ≒ 50+ tok/s'
Write-Host '   - Foundry Local CPU 推論 (Phi-4-mini 3.8B) ≒ 5-15 tok/s'
Write-Host '   - 体感「快適」と感じるのは 15 tok/s 以上'
Write-Host '   - 5 tok/s 未満は「待ちが見える」ストレス領域'
Write-Host ''
