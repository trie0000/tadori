# PPTX マニュアル RAG 拡張 — 設計ドキュメント

最終更新: 2026-05-25

## 1. 背景・目的

別ツールの展開に伴い、PowerPoint で作成された製品マニュアル (表 + 図形シェイプ
中心) に対する **「マニュアル解説問い合わせ用 AI チャット」** を提供したい。

Tadori は既にメール / OneNote の RAG 検索基盤として完成しており、`kind` フィー
ルドによる多種ソース対応 (mail / onenote / doc) のインフラを持つ。これを拡張し、
**`kind='pptx'` の新ソース** として PPTX マニュアルを取り込めるようにする。

委託先環境への展開 (機密性高) のため、**マニュアル PPTX サンプルを開発側で
入手することはできない**。仕様・設計をドキュメントで合意してから実装に入る。

## 2. スコープ

### 含むもの
- SharePoint ドキュメントライブラリ配下の **特定フォルダ URL** を入力として、そ
  の配下の全 `.pptx` を自動取り込み
- 各スライドを画像化 + 図形/表/テキスト抽出
- **GPT-5 Vision** に画像 + 補助テキストを渡し、図形配置・矢印・色分けまで意味を
  汲んだ markdown を生成
- markdown を chunk → 既存の embed → SP セグメント書き出しパイプラインへ流す
- 既存のハイブリッド検索・マルチターンチャット・引用カードを `kind='pptx'` でも
  動作させる
- 引用カードに **スライド PNG サムネ + PowerPoint で該当スライドを開く** ボタン
- 増分同期 (SP の TimeLastModified 比較で新規/更新ファイルのみ再取り込み)
- 削除検知 (SP に無いのに DB にあるファイル → セグメントから削除)

### 含まないもの (将来検討)
- アニメーション・動画埋め込みの扱い (PNG 化時点で 1 フレームに固定)
- パスワード付き PPTX (PowerPoint COM が unattended で開けない)
- スピーカーノート (本文ではないので別 chunk にするか後で判断)
- PPTX 上書き保存 / 編集機能 (Tadori は読取専用)
- 複数ライブラリの横断取り込み (1 設定 = 1 フォルダ URL。複数欲しければ複数登録)

## 3. ユーザーストーリー / UX

### 3.1. 初回セットアップ
1. ユーザが「設定 → 取り込み → PPTX マニュアル取り込み」を開く
2. 「フォルダ URL を追加」を押し、SP ドキュメントライブラリ配下のフォルダ URL を
   貼り付ける
   - 例: `https://contoso.sharepoint.com/sites/foo/Shared%20Documents/Manuals`
   - サブフォルダも再帰的に走査する (運用上、改訂版を年度別サブフォルダに置く
     ケースを想定)
3. 「同期」ボタンを押す
4. 進捗トースト: `manual-A.pptx (15/120 スライド) を処理中...`
5. 完了後、フォルダ内ファイル一覧 + 取り込み済みスライド数を表示

### 3.2. 通常利用 (検索 + ソースジャンプ)
- ユーザがチャット欄に「機微情報の申請って誰の承認が必要?」と入力
- 既存の検索フローで `kind='pptx'` のスライドもヒット
- 引用カードに:
  - 件名相当: 「manual-A.pptx — スライド 3: 申請フロー全体像」
  - サムネイル (スライド PNG)
  - markdown 化された本文の抜粋

#### ★ ソースジャンプ動作 (必須要件)

**回答本文中の `[3]` 引用番号クリック、または引用カード全体クリックで、参考の
PowerPoint ファイルを開き、ヒットしたスライドの該当ページにジャンプする。**

具体的フロー:
1. ユーザがチャット回答中の `[3]` をクリック (既存の mail / onenote と同じ
   インタラクション)
2. 該当 hit が `kind='pptx'` の場合、`relay /tadori/pptx-open` に
   `{ serverRelativeUrl, slideNo }` を POST
3. relay が PowerPoint COM で:
   - 既に該当 pptx が開かれていれば、その Presentation を再利用
   - 開かれていなければ `Presentations.Open(<SP の WebDAV 経由パス>)` で開く
   - `Application.Activate()` で PowerPoint を最前面化
   - `Presentation.Windows(1).View.GotoSlide(slideNo)` で指定スライドへ移動
4. ユーザ画面で PowerPoint がアクティブになり、該当スライドが表示される

代替フォールバック:
- relay が応答しない / PowerPoint が起動できない場合 → SP の Office Online
  ビューア (`<file-url>?action=embedview&wdSlideId=<slideNo>`) を新規タブで開く
- これによりリレーが落ちていても最低限のジャンプは保証される

UI 上の見え方:
- 引用カードに hover 時、カーソルが pointer に変わる
- カード全体クリックで上記ジャンプ (mail カードと同じ作法 — Outlook を開いて
  該当メールへ飛ぶのと完全に同じ UX)
- ボタン明示は不要 (カード全体がクリッカブル) だが、わかりやすさのため小さく
  「PowerPoint で開く」アイコン+ラベルを右上に出す

### 3.3. マニュアル改訂
1. 委託先 SP のフォルダに新版 pptx を上書きで配置
2. ユーザは「同期」ボタンを押すだけ
3. `TimeLastModified` が前回取り込み時刻より新しいファイルだけ再取り込み
4. 古いスライドの embedding は `removeSegments` で削除して新規追加 (CAS で安全に)

### 3.4. マニュアル削除
- SP からファイルを消した場合、次回同期時に「DB にあるが SP に無い」を検知し、
  該当ファイルの全 chunk を削除

## 4. データモデル

### 4.1. SavedHit (既存) の拡張
```ts
// src/search/vectorSearch.ts
export interface SavedHit {
  ...
  kind: 'mail' | 'onenote' | 'doc' | 'pptx';   // ← 'pptx' を追加
  // pptx 用フィールド (kind='pptx' のときのみ意味を持つ)
  pptxFile?: string;       // ファイル名 (例: "manual-A.pptx")
  pptxServerRelUrl?: string; // SP serverRelativeUrl (PowerPoint ジャンプ用)
  slideNo?: number;        // 1-origin スライド番号
  slideTitle?: string;     // スライドタイトル (Shape の Title placeholder)
  thumbServerRelUrl?: string; // PNG サムネの SP パス
}
```

### 4.2. segment レコード (既存形式踏襲)
```jsonl
// kind='pptx' のレコード例
{
  "id": "pptx-manual-A.pptx#3",
  "kind": "pptx",
  "subject": "申請フロー全体像",                    // = slideTitle
  "from": "manual-A.pptx",                          // = pptxFile
  "date": "2026-04-20T10:33:00Z",                   // = ファイル最終更新時刻
  "body": "# 申請フロー全体像\n\n左→右 4 段階の承認...", // = Vision LLM markdown
  "embedding": "<base64-float16>",
  "pptxFile": "manual-A.pptx",
  "pptxServerRelUrl": "/sites/foo/Shared Documents/Manuals/manual-A.pptx",
  "slideNo": 3,
  "slideTitle": "申請フロー全体像",
  "thumbServerRelUrl": "/sites/foo/Shared Documents/Tadori/thumbs/manual-A.pptx-3.png"
}
```

### 4.3. PPTX 取り込み設定 (localStorage)
```
tadori:pptx:folders   JSON配列 [{ url, label, lastSync, fileCount }, ...]
                      └ url:        フォルダの絶対 URL
                      └ label:      表示名 (省略時は URL 末尾)
                      └ lastSync:   最後に同期した UNIX ms (増分判定用)
                      └ fileCount:  前回同期時のファイル数 (UI 表示用)
```

### 4.4. 同期メタ (SP 上の Tadori Sync List)
- 既存の Tadori Sync リストに **`__pptx_sync__:<folderUrl-hash>__`** という meta 行を追加
  - 値: `{ lastSync, perFile: { "manual-A.pptx": { lastModified, slideCount } } }`
  - 用途: 増分同期判定 (ファイル別 lastModified を保持)
- ETag CAS で複数 writer の競合を防ぐ (manifest と同じ作法)

## 5. アーキテクチャ

### 5.1. 全体図

```
┌─ Browser (Tadori) ─────────────────────────────────────┐
│                                                         │
│  [UI] PPTX 取り込みペイン                               │
│   ├ フォルダ URL 入力                                   │
│   ├ 取り込み済みファイル一覧                            │
│   └ 同期/再取り込み ボタン                              │
│                ↓                                        │
│  [src/sync/pptxIngest.ts]                              │
│   ├ SP REST: フォルダ enumerate                         │
│   ├ SP REST: 各 PPTX バイナリ取得                       │
│   ├ relay にバイナリ POST                               │
│   ├ relay から slides[] 受領                            │
│   ├ 各 slide → Vision LLM → markdown                   │
│   ├ chunkText 生成                                      │
│   ├ embed (既存 client)                                 │
│   └ ingestToSegments (既存 writer.ts、kind='pptx')      │
│                                                         │
└────────────────────────────────────────────────────────┘
                ↓ HTTP                       ↑
┌─ relay (PowerShell) ───────────────────────────────────┐
│  POST /tadori/pptx-extract                              │
│   ├ multipart/form-data で PPTX バイナリ受領            │
│   ├ Temp に保存 (%TEMP%\tadori-pptx-<guid>.pptx)        │
│   ├ PowerPoint COM で開く                               │
│   │   $ppt = New-Object -ComObject PowerPoint.Application│
│   │   $pres = $ppt.Presentations.Open($tempPath, ReadOnly)│
│   ├ 各スライド走査:                                     │
│   │   - $slide.Export($pngPath, "PNG", 1920, 1080)     │
│   │   - $slide.Shapes 走査でテキスト/表抽出             │
│   ├ JSON 返却 { slides: [...] }                         │
│   └ Close + Quit + Temp 削除 (finally で必須)           │
│                                                         │
│  POST /tadori/pptx-open  (引用ジャンプ用)               │
│   ├ params: serverRelativeUrl, slideNo                  │
│   ├ PowerPoint でファイル開く + 該当スライドへジャンプ  │
│   │   $pres.SlideShowSettings.RangeType = ...           │
│   │   $pres.Slides($slideNo).Select()                   │
│   └ ウィンドウ最前面化                                  │
│                                                         │
└────────────────────────────────────────────────────────┘
                ↓
        [Azure OpenAI (GPT-5 Vision)]
         + Vision 入力 (PNG base64)
         + 補助テキスト (Shape 抽出原文)
         → markdown 出力
```

### 5.2. 新規ファイル

| パス | 役割 |
|---|---|
| `src/sync/pptxIngest.ts` | フォルダ列挙 → 各ファイル取り込み → 増分判定 → 削除検知 |
| `src/sync/pptxFolders.ts` | フォルダ設定の localStorage 読み書き |
| `src/embeddings/visionClient.ts` | GPT-5 vision に PNG + テキスト送って markdown 受領 |
| `scripts/relay/tadori-pptx-extract.ps1` | PowerPoint COM ラッパ (relay 本体から呼ぶ) |
| `docs/pptx-rag-design.md` | この文書 |

### 5.3. 改修ファイル

| パス | 改修内容 |
|---|---|
| `scripts/tadori-ai-relay.ps1` | `/tadori/pptx-extract` と `/tadori/pptx-open` エンドポイント追加 |
| `src/sharepoint/client.ts` | フォルダ enumerate (`listFolderItems(serverRelUrl, recursive)`) + ファイル binary 取得 (`fetchFileBytes`) を追加 |
| `src/search/vectorSearch.ts` | `SavedHit.kind` に `'pptx'` を追加、pptx 用フィールド追加 |
| `src/db/writer.ts` | `IngestRecord.kind` に `'pptx'` を追加 |
| `src/ui/chat.ts` | 引用カードの kind バッジに pptx 対応 (アイコン + ラベル + サムネ表示 + 「PowerPoint で開く」ボタン) |
| `src/ui/settingsHub.ts` | 「PPTX マニュアル取り込み」ペイン追加 |
| `src/utils/chunking.ts` | 既に汎用化済みのはずだが、pptx 用に念のため確認 |

## 6. relay 実装詳細

### 6.1. PPTX → スライド配列

PowerShell + PowerPoint COM。安全策として `try/finally` で必ず `Quit()` する。

```powershell
function Extract-Pptx($srcPath, $tmpDir) {
    $ppt = $null
    $pres = $null
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
        # PowerPoint は最小化不可。OS X11 みたいに HEADLESS は無いので画面は出る。
        # ただし WithWindow:$false で開けば不可視 (Office 2016 以降)。
        $pres = $ppt.Presentations.Open(
            $srcPath, [bool]$true, [bool]$false, [bool]$false  # ReadOnly, Untitled, WithWindow=false
        )
        $slides = @()
        foreach ($slide in $pres.Slides) {
            $no = $slide.SlideNumber
            $pngPath = Join-Path $tmpDir "slide-$no.png"
            # Export は同期。サイズは Vision LLM の上限 (2048px) を意識
            $slide.Export($pngPath, "PNG", 1920, 1080)

            # Shape 走査 — テキスト/表抽出
            $title = ""
            $textBlocks = @()
            $tables = @()
            foreach ($shape in $slide.Shapes) {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $t = $shape.TextFrame.TextRange.Text
                    if ($shape.PlaceholderFormat -and $shape.PlaceholderFormat.Type -eq 1) {
                        # Title placeholder
                        $title = $t
                    } else {
                        $textBlocks += $t
                    }
                }
                if ($shape.HasTable) {
                    $tableObj = Extract-Table $shape.Table
                    $tables += , $tableObj
                }
            }

            $pngBytes = [IO.File]::ReadAllBytes($pngPath)
            $slides += @{
                slideNo    = $no
                title      = $title
                pngBase64  = [Convert]::ToBase64String($pngBytes)
                rawText    = ($textBlocks -join "`n")
                tables     = $tables
            }
            Remove-Item $pngPath -ErrorAction SilentlyContinue
        }
        return @{ slides = $slides }
    } finally {
        if ($pres) { $pres.Close() }
        if ($ppt)  { $ppt.Quit() }
        # COM 解放 (PowerShell は GC で残ることがあるので明示)
        [Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
        [GC]::Collect()
    }
}
```

### 6.2. 並列性

- PowerPoint COM は **シングルスレッド前提**。relay の HTTP リスナを並列処理可
  にすると COM が壊れる。
- → `/tadori/pptx-extract` は **mutex で逐次化** する。一度に 1 ファイル。
- 同時に複数ユーザが同期しても relay 側でキューイング。

### 6.3. リソース管理

- `WithWindow=$false` で開いても PowerPoint プロセスは残る → `Quit()` は必須
- 例外時の `[Runtime.InteropServices.Marshal]::ReleaseComObject` 抜けは memory leak
- ノマド機の Office パスワード必須運用とぶつかる可能性 → 設定ハブで「Office のロックが解除されていること」を明示

### 6.4. PowerPoint 開く + スライドへジャンプ (引用クリック時)

```powershell
function Open-PptxAtSlide($spFileUrl, $slideNo) {
    # 既存の PowerPoint インスタンスを取得 (なければ新規)
    $ppt = $null
    try {
        $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
    } catch {
        $ppt = New-Object -ComObject PowerPoint.Application
    }

    # PowerPoint を可視化 (引用ジャンプ用なので画面表示が要件)
    $ppt.Visible = $true

    # 既に同じファイルが開いていれば再利用
    $target = $null
    foreach ($p in $ppt.Presentations) {
        if ($p.FullName -ieq $spFileUrl -or
            $p.FullName -ieq ($spFileUrl -replace '^https://', 'https:\\')) {
            $target = $p
            break
        }
    }
    if (-not $target) {
        # SP URL から直接開ける (Office は WebDAV 経由で SP を扱える)
        # ReadOnly:$false (編集可) で開く — ユーザが書き込む可能性も尊重
        $target = $ppt.Presentations.Open($spFileUrl, [bool]$false, [bool]$false, [bool]$true)
    }

    # スライドへジャンプ + アクティブ化
    $target.Windows.Item(1).Activate()
    $target.Windows.Item(1).View.GotoSlide([int]$slideNo)

    # PowerPoint ウィンドウを最前面化
    Add-Type @"
        using System.Runtime.InteropServices;
        public class Win {
            [DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr h);
        }
"@
    [Win]::SetForegroundWindow([System.IntPtr]::new($target.Windows.Item(1).HWND)) | Out-Null
}
```

**重要な点:**
- 既に開いているプレゼンを再利用 (二重起動防止 + ユーザの編集中状態を保持)
- SP の WebDAV パス (`https://contoso.sharepoint.com/.../file.pptx`) を直接 Open 可能。
  Cookie 認証は OS の Office 認証経由で透過的に通る
- `View.GotoSlide(slideNo)` が標準 (1-origin)。`SlideShowSettings.RangeType` は本番モード移行用なので使わない (編集ビューでジャンプしたい)
- ジャンプ後、Tadori UI には戻ってこない (= PowerPoint がアクティブのまま)。
  ユーザは Alt+Tab で戻る前提

### 6.5. フォールバック (relay 落ち / 非 Windows / Office 未インストール)

引用ジャンプは relay 必須機能だが、relay 落ちでも最低限の閲覧を保証する:

1. ブラウザ側で先に `Test-RelayUp` (既存) を呼んで relay の死活確認
2. relay 死亡 → SP の Office Online ビューアの URL に新規タブで遷移:
   ```
   https://contoso.sharepoint.com/sites/foo/_layouts/15/Doc.aspx?
     sourcedoc=<file-guid>&action=embedview&wdSlideId=<slideNo>
   ```
3. Office Online は SP の Cookie で認証通る + ブラウザ完結 + スライド指定可能
4. 編集はできないが閲覧+検索引用ジャンプとして十分

## 7. Vision LLM (GPT-5) 呼び出し

### 7.1. 入力

- PNG (base64) — 1 スライドの完全レンダリング
- 補助テキスト — `rawText` (Shape 抽出原文。OCR ノイズの補強)
- 表データ — `tables` (構造化済み行列)

### 7.2. プロンプト案

```
あなたは技術マニュアルの図解を読み取って Markdown に構造化するアシスタントです。
提供される画像は PowerPoint のスライド 1 枚です。下記の点に注意してください:

1. スライドのタイトルは `# <タイトル>` で書く
2. 概要 / 各セクションは `## <見出し>` で構造化
3. 図形による「流れ図」「構成図」は、矢印の向き・色分け・位置関係から論理を読み取り、
   箇条書きや表で再表現する
   例: A→B→C のフローなら "1. A: 〜  2. B: 〜  3. C: 〜" のように番号付き箇条書きに
4. 表 (table) はそのまま Markdown 表に変換 (構造化済みデータも提供される)
5. 色分けに意味がある場合 (例: 必須=赤、任意=青) は本文中に注記
6. 装飾的な要素 (背景画像、ロゴ等) は無視
7. このスライドが取り扱う「主題」を最後に 1 行サマリで書く: `> 要点: <50字以内>`

補助情報:
- Shape から抽出した生テキスト (順不同):
  <rawText>
- 検出された表構造:
  <tables JSON>

回答は Markdown のみ。前置きや「私は AI です」のような断り書きは不要。
```

### 7.3. コスト試算 (GPT-5 vision)

- 1 スライド画像 (1920x1080 PNG, 中精度): 入力 ~1500 tok + 出力 ~500 tok ≒ 2000 tok
- GPT-5 vision の料金が GPT-4o 相当 (input $2.5/M, output $10/M) と仮定すると 1 スライド ≒ 1 円
- **100 スライドのマニュアル ≒ 100 円**
- 取り込みは原則 1 回 + 改訂時のみ再実行 → 運用コストは低い

### 7.4. リトライ・エラー処理

- 429 / 5xx は既存の `sleepRespectingAbort + inferRetryDelayMs` パターン再利用
- 完全失敗時は当該スライドだけスキップ + ログに残す
- 「3 枚連続失敗したらマニュアル全体を中断して通知」のセーフティ

## 8. SharePoint REST API

### 8.1. フォルダ enumerate

```http
GET /_api/web/getFolderByServerRelativeUrl('/sites/foo/Shared Documents/Manuals')?$expand=Files,Folders&$select=Files/Name,Files/ServerRelativeUrl,Files/TimeLastModified,Files/Length,Folders/ServerRelativeUrl
```

- 再帰: サブフォルダ分は別途叩く (`Folders` で取得した path で再帰)
- フィルタ: `.pptx` (大文字小文字不問) で client 側絞り込み

### 8.2. ファイル binary 取得

```http
GET /_api/web/getFileByServerRelativeUrl('/sites/foo/Shared Documents/Manuals/manual-A.pptx')/$value
```

- レスポンスは ArrayBuffer で受ける
- 巨大ファイル (~50MB) も想定してストリーミングは不要だが、メモリ消費に注意

### 8.3. サムネ PNG のアップロード

既存の `uploadFileText` の bytes 版を新設 (`uploadFileBytes`)。

```ts
SharePointClient.uploadFileBytes(serverRelUrl: string, bytes: ArrayBuffer, opts: { overwrite?: boolean })
```

PNG サムネは `<segments-root>/thumbs/<file>-<slideNo>.png` として保存。
カーディナリティは「全マニュアルのスライド数」分。

## 9. UI 設計

### 9.1. 「PPTX マニュアル取り込み」ペイン

設定 → 取り込み → 既存「Outlook から取り込み」/「OneNote から取り込み」の下に追加。

```
┌── PPTX マニュアル取り込み ────────────────────────────┐
│                                                       │
│  SharePoint のドキュメントライブラリ配下のフォルダを │
│  指定して、その配下の .pptx を一括取り込みします。   │
│                                                       │
│  ┌─ 取り込み済みフォルダ ───────────────────────────┐│
│  │                                                  ││
│  │ [追加] [全フォルダを同期]                        ││
│  │                                                  ││
│  │ ┌─ Manuals (15 ファイル, 最終同期 1h 前) ──────┐ ││
│  │ │  └ manual-A.pptx (120 スライド)              │ ││
│  │ │  └ manual-B.pptx (45 スライド)               │ ││
│  │ │  └ manual-C.pptx (88 スライド)               │ ││
│  │ │  ...                                          │ ││
│  │ │ [同期] [削除]                                 │ ││
│  │ └───────────────────────────────────────────────┘ ││
│  └──────────────────────────────────────────────────┘│
│                                                       │
└───────────────────────────────────────────────────────┘
```

### 9.2. 引用カード (kind='pptx')

```
┌─[3] 📊 PPTX  manual-A.pptx — スライド 3: 申請フロー全体像  (0.873)
│ ┌─ サムネ ──────────────────┐
│ │  (スライドの 200x150 PNG) │
│ └─────────────────────────────┘
│ # 申請フロー全体像
│ ## フロー (左→右)
│ 1. **申請者**: 申請書を提出
│ ...
│ [PowerPoint で開く]
└────────────────────────────────────
```

## 10. セキュリティ / 委託先環境配慮

- **委託先 PPTX が開発側に届かない前提** → 委託先環境内で完結する設計
- relay は委託先 PC で動作。Tadori UI も委託先ブラウザで動作。Vision LLM は委託先
  契約の Azure OpenAI を使う
- 開発者 (Anthropic 開発側) はこの設計ドキュメントとコードだけ書く
- PNG サムネと markdown も委託先 SP に保存される → 委託先の閲覧権限の範囲を超えない
- relay が PPTX を Temp に書く → セッション終了時に削除 (`finally`)。プロセスクラッシュ時の Temp 残骸は OS の TEMP クリーンアップ任せ

## 11. テスト計画 (委託先環境)

開発側でサンプルが無いので、委託先での動作確認を以下のステップで:

1. **疎通テスト**: relay の `/tadori/pptx-extract` に小さい (1 スライド) ダミー pptx
   を投げ、JSON が返ることを確認
2. **Vision テスト**: 1 スライドだけ Vision LLM に通し、markdown が想定通りか確認
3. **チャンク末端**: chunkText 生成 + embed + segment 書き込みまで通すが、
   検索クエリで `kind='pptx'` がヒットするか
4. **小規模マニュアル**: 10 スライド程度の本物マニュアル 1 本で end-to-end
5. **大規模マニュアル**: 100 スライド超で時間・コスト・メモリ実測
6. **マルチファイル**: フォルダに 3 ファイル置いて一括同期
7. **増分**: 1 ファイル更新後の再同期で、更新分だけ再処理されるか
8. **削除**: 1 ファイル消した後の同期で、該当 chunk が削除されるか
9. **競合**: 2 ブラウザ同時同期で manifest CAS が機能するか

## 12. オープン項目 (実装前に決めたい)

下記の既定値で進める (実装中に問題が出たら都度見直し)。

| # | 項目 | 既定値 (採用) | 理由 |
|---|---|---|---|
| 1 | PowerPoint プロセスを毎回 Quit() するか | **取り込み時は毎回 Quit / 引用ジャンプ時は維持** | 取り込みは大量処理で COM 不安定化リスク。引用ジャンプはユーザの編集状態保持が嬉しい |
| 2 | スピーカーノートは取り込むか | **取り込む (本文に concat、`## ノート` 見出し付き)** | 補足情報として価値あり。Vision LLM への入力には含めず raw text として追加 |
| 3 | チャンク粒度 | **1 スライド = 1 chunk** | 引用単位とジャンプ先が一致して UX 直感的。長すぎたら例外的に分割 |
| 4 | フォルダ URL 入力の形式 | **絶対 URL 受付 → 内部で serverRelativeUrl に正規化** | ユーザはブラウザのアドレスバーからコピーが普通 |
| 5 | サムネサイズ | **320x180 サムネ + 1920x1080 フル両方保存** | カード用と引用時詳細用を分ける。フルは on-demand ロード |
| 6 | 同期トリガ | **手動のみ (将来 Sticky 統合の余地は残す)** | マニュアル改訂は頻度低 + 取り込みコスト大なので明示操作が望ましい |
| 7 | 取り込み中の UI | **進捗バー + 詳細ログトースト** | 取り込みは数十分〜のオーダー。ユーザが進行状況を把握できる必要あり |

## 13. 工数 (確定後の実装計画)

| Step | 内容 | 工数 |
|---|---|---|
| 1 | SP client 拡張 (folder enumerate + binary fetch) | 0.5d |
| 2 | relay PPTX 抽出エンドポイント | 0.5d |
| 3 | Vision LLM クライアント | 0.5d |
| 4 | pptxIngest パイプライン + 増分判定 | 1d |
| 5 | UI ペイン + 同期表示 | 0.5d |
| 6 | 引用カード `pptx` 種別 + PowerPoint ジャンプ | 0.5d |
| 7 | テスト + バグ修正 (委託先環境) | 0.5d |
| **計** | | **4d** (PoC 1〜3 で 1.5d、本体 4〜6 で 2d) |

## 14. 承認

- [ ] スコープに過不足ないか (§2)
- [ ] データモデルが既存 Tadori と整合か (§4)
- [ ] アーキテクチャ図が現実的か (§5)
- [ ] Vision LLM プロンプトの方針 OK か (§7.2)
- [ ] UI 設計の方向性 OK か (§9)
- [ ] オープン項目 §12 の決定

承認後、実装フェーズ (§13 の順) に入る。
