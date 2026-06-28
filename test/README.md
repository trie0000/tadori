# ローカル実機検証 (push 前にこれを通す)

机上の理屈でなく、**本番コードをこの Mac で実際に動かして**回帰を検出するための環境。
新しい依存は増やしていない (esbuild + Node 標準 `node --test` + pwsh のみ)。

## 実行

```bash
npm test              # 検索/取り込みロジック (本番 searchVectors を実行)
npm run test:relay-pdf # relay の PDF 抽出 (PdfPig の DLL ロード + 抽出)
npm run type-check    # tsc 型チェック
```

`npm test` は失敗時に非0で終了する。**push 前に最低 `npm test` と `npm run type-check` を通すこと。**

## 仕組み

- `test/run.mjs` が `test/*.test.ts` を esbuild で Node 向けにバンドルして実行する。
- 本番コードのうち外部 I/O を持つ部分だけスタブに差し替える (それ以外は実コード):
  - `db/engine` → `test/_stubs/engine.ts` … SP/IndexedDB に触れず、注入した VectorDb を返す
  - `embeddings/router` → `test/_stubs/router.ts` … Azure/Voyage に触れず、注入したクエリ埋め込みを返す
  - `mailhtml`/`mailtext` … DOMPurify を import 時実行するため no-op 化 (Node に DOM が無い)
- 埋め込みは `test/_fixtures.ts` の `vec(seed)` で決定論生成。検索ランキングが再現可能。

これにより `VectorDb.search` / doc スコープ / kind フィルタ / 次元不一致処理など、
**ブラウザで動くのと同じ JS ロジック**を Node 上で検証できる。

## カバー済みの回帰

- doc フォルダスコープが空配列 `[]` で doc 全件が消える (`search.test.ts`)
- kind フィルタ / doc フォルダ絞り込み (一致・不一致)
- クエリと次元不一致のレコードがスコア0で沈む
- PdfPig 本体 + 実行時依存 DLL (System.Memory 等) の欠落、AssemblyResolve 登録の回帰 (`relay-pdf.test.ps1`)

## できないこと (正直な限界)

- SharePoint 上での Cookie 認証・実セグメント配布、Azure OpenAI 実エンドポイント、
  Windows COM (Outlook/OneNote/PowerPoint) は Mac で再現できない。
  これらに依存する経路は最終的に実機ブラウザ/Windows relay での確認が必要。
- ここで担保できるのは「純ロジック + PDF 抽出」まで。新しい不具合を見つけたら、
  まずここに**失敗するテストを足してから**直すこと。
