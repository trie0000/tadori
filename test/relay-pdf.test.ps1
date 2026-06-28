# relay の PDF 抽出 (PdfPig) を実機ロード込みで検証する。
# 目的: scripts/lib/pdfpig/ の DLL 一式 (本体 + 実行時依存) が揃っていて、
#       AssemblyResolve 経由で実際に PdfDocument.Open → テキスト抽出まで通ることを保証。
#       過去に「Register-ObjectEvent が使えない」「System.Memory 等の依存DLL欠落」で
#       PDF 抽出が全滅した回帰を二度と通さないためのガード。
#
# 実行: pwsh -NoProfile -File test/relay-pdf.test.ps1   (npm run test:relay-pdf)
# .NET Framework が無い Mac/pwsh でも DLL ロード可否と抽出は検証できる。

$ErrorActionPreference = 'Stop'
$root   = Split-Path -Parent $PSScriptRoot
$libDir = Join-Path $root 'scripts/lib/pdfpig'
$pdf    = Join-Path $PSScriptRoot 'fixtures/sample.pdf'
$fail   = 0

function Assert($cond, $msg) {
  if ($cond) { Write-Host "  PASS: $msg" -ForegroundColor Green }
  else { Write-Host "  FAIL: $msg" -ForegroundColor Red; $script:fail++ }
}

Write-Host '== PdfPig 実行時依存 DLL の存在 ==' -ForegroundColor Cyan
$required = @(
  'UglyToad.PdfPig','UglyToad.PdfPig.Core','UglyToad.PdfPig.Fonts',
  'UglyToad.PdfPig.Tokenization','UglyToad.PdfPig.Tokens',
  'System.Memory','System.Buffers','System.Numerics.Vectors',
  'System.Runtime.CompilerServices.Unsafe','Microsoft.Bcl.HashCode'
)
foreach ($n in $required) {
  Assert (Test-Path -LiteralPath (Join-Path $libDir "$n.dll")) "$n.dll が同梱されている"
}

Write-Host '== AssemblyResolve 登録 + ロード ==' -ForegroundColor Cyan
$script:PdfPigLibDir = $libDir
$resolver = [System.ResolveEventHandler] {
  param($s, $e)
  try {
    $simple = ($e.Name -split ',')[0].Trim()
    $cand = Join-Path $script:PdfPigLibDir ("$simple.dll")
    if (Test-Path -LiteralPath $cand) { return [Reflection.Assembly]::LoadFrom($cand) }
  } catch { }
  return $null
}
try {
  [AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
  Assert $true 'add_AssemblyResolve で登録できる (Register-ObjectEvent 回帰なし)'
} catch { Assert $false "AssemblyResolve 登録: $($_.Exception.Message)" }

foreach ($n in @('UglyToad.PdfPig.Core','UglyToad.PdfPig.Tokens','UglyToad.PdfPig.Tokenization','UglyToad.PdfPig.Fonts','UglyToad.PdfPig')) {
  try { [Reflection.Assembly]::LoadFrom((Join-Path $libDir "$n.dll")) | Out-Null }
  catch { Assert $false "ロード失敗 $n : $($_.Exception.Message)" }
}

Write-Host '== PdfDocument.Open + テキスト抽出 ==' -ForegroundColor Cyan
try {
  $bytes = [System.IO.File]::ReadAllBytes($pdf)
  $doc = [UglyToad.PdfPig.PdfDocument]::Open($bytes)
  $sb = New-Object System.Text.StringBuilder
  for ($i = 1; $i -le $doc.NumberOfPages; $i++) { [void]$sb.AppendLine([string]$doc.GetPage($i).Text) }
  $doc.Dispose()
  $text = $sb.ToString()
  Assert ($text -match 'Tadori PDF fixture text 12345') "抽出テキストにフィクスチャ文字列が含まれる ('$($text.Trim())')"
} catch {
  Assert $false "Open/抽出で例外: $($_.Exception.Message)"
}

Write-Host ''
if ($fail -eq 0) { Write-Host 'relay-pdf: 全テスト PASS' -ForegroundColor Green; exit 0 }
else { Write-Host "relay-pdf: $fail 件 FAIL" -ForegroundColor Red; exit 1 }
