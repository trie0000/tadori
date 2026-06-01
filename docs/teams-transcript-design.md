# Teams 会議文字起こし RAG 取り込み — 設計ドキュメント

最終更新: 2026-06-01

## 1. 背景・目的

Teams 会議の文字起こし (transcript) を Tadori の検索対象に加え、「あの会議で
何が決まったか」「○○について誰が何を言ったか」を自然言語で検索できるようにする。

御社環境の制約 (Cookie 認証 SP REST のみ / Graph API のアプリ登録は不可 /
relay は COM 用) を踏まえ、**SharePoint・OneDrive 上に保存される `.vtt`
ファイルをファイルベースで取り込む** 方式を採用する。これは既存の PPTX
取り込み (SP フォルダ URL → ファイル走査) パターンをほぼそのまま流用でき、
relay も Vision も不要。

## 2. スコープ

### 含むもの
- SP / OneDrive の「Recordings」フォルダ等を **フォルダ URL 指定** で取り込み
- `.vtt` (WebVTT) 字幕ファイルのパース (話者 + 発言 + タイムスタンプ抽出)
- **生チャンク方式**: 発言を ~800 字単位にまとめてそのまま embed (LLM 要約なし、
  取り込みコスト 0 = 埋め込み API のみ)
- 既存の増分同期 / 削除検知 / サイト別分離をそのまま継承
- 引用カードから録画 (Stream) の該当時刻へジャンプ

### 含まないもの (将来検討)
- `.docx` 形式の議事録 (まず .vtt のみ。docx は zip+xml パースが必要)
- 会議要約チャンクの LLM 生成 (生チャンクのみ。後で追加可能)
- 録画されず文字起こしのみの会議 (ファイル化されず Graph/Stream 側のみ → 対象外)
- 話者の名寄せ / 発言者統計

## 3. アクセス経路: .vtt ファイルベース

Teams の文字起こしは録画と一緒に SP / OneDrive にファイルとして保存される。

| 会議の種類 | 保存先 | ファイル |
|---|---|---|
| チャネル会議 | チームの SP サイト `/<channel>/Recordings/` | `<会議名>.mp4` + `<会議名>.vtt` |
| 通常会議 (予定/個人) | 主催者の OneDrive `/Recordings/` (`-my.sharepoint.com/personal/<user>/`) | 同上 |

- `.vtt` が文字起こしの実体。既存の `listFolderItems` / `fetchFileBytes` で取得可能。
- **前提**: 録画 + 文字起こしが有効な会議のみファイル化される。文字起こしのみ
  (録画なし) の会議は対象外 (Graph/Stream 側にしか無い)。
- OneDrive 個人サイトは host が `-my.sharepoint.com`、パスが `/personal/<user>`。
  既存の `siteCollectionPath` 正規表現は `/personal/` を既にカバー済み。

### Graph API を使わない理由
`/me/onlineMeetings/{id}/transcripts` は綺麗だが OAuth アプリ登録 / トークンが
必要で、御社環境では現実的でない。Cookie 認証で完結するファイルベースを採用。

## 3.5. .vtt に含まれる情報 / 含まれない情報 (重要)

Teams の .vtt は「発言の字幕」であり、会議のメタ情報は持たない。

| 情報 | .vtt 内 | 取得方法 |
|---|---|---|
| 発言テキスト | ✅ | cue のテキスト行 |
| 発言者名 (表示名) | ✅ | `<v 山田太郎>` voice タグ |
| タイムスタンプ | ✅ | `00:00:01.230 --> ...` |
| **会議名** | ❌ | ファイル名から推定 |
| **会議日時** | ❌ | ファイル名 → 無ければ lastModified |
| 参加者一覧 (非発言者含む) | ❌ | 取得不可 (発言者しか出ない) |
| チャネル/チーム名 | ❌ | フォルダパスから推測可 (任意) |

注意点 (索引の信頼性に関わる):
- 発言者名は Teams **表示名** (メール/社員ID は無い)。同姓同名は区別不可。
- 文字起こしは発言者の取り違え・誤変換がありうる → 「誰が言ったか」検索は
  参考程度。100% ではない前提で扱う。
- 会議名・日時は .vtt 内に無いため **ファイル名が唯一の手がかり**。命名規則は
  テナント設定 / Teams バージョンで揺れる。

### 確定した方針
- **会議名・日時**: ファイル名をパースして抽出。取れなければファイル名そのまま
  (会議名) + lastModified (日時)。
  - 想定パターン例: `週次定例-20260528_100000-Meeting Recording.vtt`,
    `「プロジェクトX キックオフ」-20260528_140000.vtt`
  - `YYYYMMDD_HHMMSS` 形式の日時トークンを正規表現で拾い、それ以外を会議名に。
  - 末尾の `-Meeting Recording` 等の定型 suffix は除去。
- **発言者名**: チャンク本文に `[mm:ss 話者名]` の形で埋め込む。発言者も embed
  対象になり「○○が何を言ったか」検索に効く。

## 4. .vtt パース仕様

### WebVTT フォーマット例
```
WEBVTT

00:00:01.230 --> 00:00:05.670
<v 山田太郎>では本日の議題ですが、申請フローの見直しについて…

00:00:05.900 --> 00:00:09.100
<v 鈴木花子>はい、前回の宿題から確認させてください。
```

### パーサ (src/transcript/vtt.ts, 新規)
- `WEBVTT` ヘッダをスキップ
- 各キュー (cue) を分解: `<開始 --> 終了>` 行 + テキスト行
- テキストから話者タグ `<v 名前>` を抽出 (無ければ話者不明)
- HTML エンティティ / `<c>` 等の装飾タグを除去
- 出力: `{ startSec, endSec, speaker, text }[]`

```ts
export interface VttCue {
  startSec: number;   // 開始秒 (ジャンプ用)
  endSec: number;
  speaker: string;    // "<v 名前>" の名前。無ければ ''
  text: string;       // 装飾除去済みの発言
}
export function parseVtt(text: string): VttCue[];
```

### タイムスタンプ → 秒
`00:01:05.230` → `65.23` (時:分:秒.ミリ秒)。先頭が `分:秒` 2 要素の場合もある。

## 5. チャンク化 (生チャンク)

- 連続する cue を **~800 字 / overlap 80 字** でまとめる (OneNote と同じ
  `splitIntoChunks` を流用できるよう、まず話者付きテキストへ整形してから渡す)
- チャンク本文の形式:
  ```
  [00:01:05 山田太郎] では本日の議題ですが、申請フローの見直しについて…
  [00:01:30 鈴木花子] はい、前回の宿題から確認させてください。
  ...
  ```
- 各チャンクが保持するメタ:
  - `startSec`: チャンク先頭 cue の開始秒 (録画ジャンプ用)
  - 主な話者 (チャンク内最頻 or 先頭話者)

## 6. データモデル

### 6.1. kind = 'transcript' を新設
mail / onenote / doc / pptx に並ぶ第 5 の種別。下記すべてに追加:
- `src/sync/segments.ts` SegmentRecord.kind
- `src/db/store.ts` MailRecord.kind
- `src/db/writer.ts` IngestMail.kind
- `src/search/vectorSearch.ts` MailHit.kind
- `src/ui/sessions.ts` SavedHit.kind
- `src/search/searchKinds.ts` SearchKind (UI のソース選択チップ)

### 6.2. レコード (kind='transcript')
```jsonl
{
  "id": "transcript://<serverRelUrl>#3",
  "kind": "transcript",
  "messageId": "transcript://<serverRelUrl>#3",
  "conversationId": "<serverRelUrl>",        // 同一会議の全チャンクを束ねる
  "subject": "週次定例 2026-05-28",           // 会議名 (ファイル名から)
  "from": "山田太郎 ほか",                     // 主な話者 / チャネル名
  "date": "2026-05-28T10:00:00Z",            // 会議日時 (ファイル lastModified or 名前から)
  "body": "[00:01:05 山田] …\n[00:01:30 鈴木] …",
  "chunkIdx": 3,
  "chunkCount": 20,
  "emb": "<base64-float16>",
  // transcript 固有メタ:
  "transcriptFile": "週次定例 2026-05-28.vtt",
  "vttServerRelUrl": "/sites/.../Recordings/週次定例 2026-05-28.vtt",
  "recordingServerRelUrl": "/sites/.../Recordings/週次定例 2026-05-28.mp4",  // 同名 .mp4 があれば
  "startSec": 65,                             // チャンク先頭の開始秒 (録画ジャンプ用)
  "srcHash": "<内容ハッシュ>"                  // 差分判定 (PPTX と同じ仕組み)
}
```

新フィールド (segments/store/writer/vectorSearch/sessions に追加):
`transcriptFile?`, `vttServerRelUrl?`, `recordingServerRelUrl?`, `startSec?`

### 6.3. フォルダ設定 (localStorage, サイト別)
PPTX の `pptxFolders.ts` と同型を `transcriptFolders.ts` として複製。
キー: `tadori:transcript:folders:<siteHash>`

## 7. パイプライン (src/sync/transcriptIngest.ts, 新規)

PPTX の `pptxIngest.ts` を強く参考にする。relay 呼び出しが無い分むしろ単純。

```
syncTranscriptFolder(folder, s, fallbackSiteUrl, onProgress, signal, opts):
  1. resolveSpFolder(folder.url) → siteUrl + folderServerRel
  2. listFolderItems(folderServerRel, {recursive}) → .vtt でフィルタ
  3. 増分判定 (lastModified) + 削除検知    ← pickTargets を流用/複製
  4. 各 .vtt:
       a. fetchFileBytes → TextDecoder('utf-8') で文字列化
          (vtt は UTF-8。BOM があれば除去)
       b. parseVtt → cue[]
       c. チャンク化 (話者付き整形 → splitIntoChunks)
       d. srcHash 計算 (PPTX と同じく差分スキップに使う)
       e. IngestMail[] 組み立て (kind='transcript')
       f. ingestToSegments
  5. perFile 更新
```

- **Vision も relay も無い** → ブラウザ内で完結 (fetch + parse + embed)
- 差分・削除検知・サムネ無し以外は PPTX と同じ構造

## 8. UI

### 8.1. 取り込みペイン (settingsHub)
PPTX の `buildPptxImport` と同型の `buildTranscriptImport` を追加。
- フォルダ URL 入力 (Recordings フォルダ等) + 再帰チェック
- 登録済みフォルダ一覧 + 同期 / 個別再取込 / 削除
- 進捗バー
- 「サムネ再生成」は不要 (画像が無いため)

### 8.2. 検索対象ソース選択
`searchKinds.ts` に `'transcript'` を追加:
- ラベル: 「会議」
- アイコン: 新規 (吹き出し+人 or マイク風)

### 8.3. 引用カード
- バッジ: 「Teams」
- 件名: 「週次定例 2026-05-28 — 00:01:05 山田太郎」
- 「録画を開く」ボタン:
  - `recordingServerRelUrl` があれば Stream のディープリンク
    `<...>.mp4?t=<startSec>` を新規タブで開く → **該当時刻から再生**
  - 録画が無ければ `.vtt` をブラウザ表示 (fallback)
  - relay 不要 (すべてブラウザの window.open)
- 本文: チャンク (話者付きテキスト) をそのまま表示

## 9. 新規 / 改修ファイル

| パス | 区分 | 内容 |
|---|---|---|
| `src/transcript/vtt.ts` | 新規 | WebVTT パーサ |
| `src/sync/transcriptFolders.ts` | 新規 | フォルダ設定 (サイト別 localStorage) |
| `src/sync/transcriptIngest.ts` | 新規 | 取り込みパイプライン |
| `src/search/searchKinds.ts` | 改修 | 'transcript' 追加 |
| `src/sync/segments.ts` | 改修 | kind + transcript メタ追加 |
| `src/db/store.ts` | 改修 | 同上 + applyRecord 伝播 |
| `src/db/writer.ts` | 改修 | IngestMail に追加 + records 構築 |
| `src/search/vectorSearch.ts` | 改修 | MailHit に追加 + toHit 伝播 |
| `src/ui/sessions.ts` | 改修 | SavedHit に追加 |
| `src/ui/settingsHub.ts` | 改修 | buildTranscriptImport ペイン |
| `src/ui/chat.ts` | 改修 | 引用カード transcript 種別 + 録画ジャンプ |
| `src/ui/icons.ts` | 改修 | 会議アイコン追加 |

## 10. 工数

| Step | 内容 | 目安 |
|---|---|---|
| 1 | vtt.ts パーサ + 単体確認 | 0.5d |
| 2 | kind='transcript' をデータモデル全層に追加 | 0.5d |
| 3 | transcriptFolders + transcriptIngest | 0.5d |
| 4 | UI ペイン + ソース選択チップ | 0.5d |
| 5 | 引用カード + 録画ジャンプ | 0.5d |
| **計** | | **約 2.5d** |

PPTX より relay/Vision/サムネが無い分軽い。

## 11. テスト計画 (委託先/実環境)

1. .vtt 1 本をブラウザで parse → cue 数・話者・タイムスタンプ確認
2. チャンク化 → embed → segment 書込 → 検索ヒット確認
3. 「会議」ソースだけに絞って検索 → transcript だけ返る
4. 引用カードの「録画を開く」→ Stream が該当時刻から再生
5. 増分: 同じフォルダ再同期で未変更ファイルがスキップされる
6. 削除: SP から .vtt を消して再同期 → chunk が消える
7. サイト切替で別サイトの会議録が混ざらない

## 12. 決定済み事項

- ✅ 取込スコープ: **生チャンク** (LLM 要約なし)
- ✅ 対応形式: **.vtt のみ** (.docx は将来)
- ✅ アクセス: ファイルベース (Graph 不使用)
- ✅ 会議名・日時: **ファイル名から推定**、無ければ lastModified
- ✅ 発言者名: **チャンク本文に埋め込む** (`[mm:ss 話者名]`)
- ✅ **運用モデル: B = 専用フォルダに手動アップ**。SP に会議録置き場を 1 つ
  作り、対象の .vtt を手動でコピー/アップロード → その 1 フォルダを Tadori に
  登録して同期。PPTX と同一の「フォルダ URL 登録 + 増分同期」機構で実現でき、
  コード差分なし。
  - 1 フォルダ集約が前提なので recursive は基本 OFF (任意で ON 可)
  - 「同期」= 新規/更新 .vtt だけ取り込み (全件やり直しではない)
  - 不要な会議は手動で対象 .vtt をアップしない / フォルダから消せばよい
- ✅ 設計ドキュメント先行 → 本書

## 13. オープン項目 (実装中に確認)

| # | 項目 | 暫定 |
|---|---|---|
| 1 | 録画ジャンプの URL 形式 | `<mp4>?t=<sec>` (Stream)。動かなければ .vtt 表示に fallback |
| 2 | 話者名のチャンク内集約 (from 列) | 先頭話者 or 最頻。まず先頭話者で |
| 3 | 1 会議の最大チャンク数上限 | 長時間会議の暴発防止に上限検討 (まず無制限で実測) |
