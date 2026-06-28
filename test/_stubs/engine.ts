// テスト用 engine スタブ。本番 searchVectors はこの getEngine を呼ぶが、SP/IndexedDB
// には触れず、テストが注入した VectorDb をそのまま返す。これにより検索ロジック本体
// (フィルタ連鎖・dedup・pull 拡張・doc スコープ) を実コードで検証できる。

import { VectorDb } from '../../src/db/store';

let _db = new VectorDb();
let _stats = { manifestSealed: 1, cached: 0, downloaded: 0, dbSize: 0 };

export function __setDb(db: VectorDb): void { _db = db; _stats.dbSize = db.size; }

export async function getEngine(_siteUrl: string): Promise<any> {
  return {
    db: _db,
    sync: { lastStats: _stats, sync: async () => {} },
    store: {}, cache: {},
  };
}
