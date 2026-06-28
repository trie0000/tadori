// 同期オーケストレーション。
// 起動時: IndexedDB キャッシュからセグメントを適用 → manifest を見て差分だけ
// SharePoint からDL → 適用 & キャッシュ。以降は定期ポーリングで差分追従。

import { SpVectorStore } from './spStore';
import { SegmentCache } from '../db/cache';
import { VectorDb } from '../db/store';
import { missingSealed, type Manifest } from './segments';
import { relayLog } from '../lib/relayLog';

/** doc/transcript を含むセグメントだけ詳細を relay へ。なぜ DB に入らないかの切り分け用。 */
function logSeg(where: string, id: string, db: VectorDb): void {
  const a = db.lastApply;
  const interesting = (a.byKind['doc'] ?? 0) + (a.byKind['transcript'] ?? 0) > 0;
  if (!interesting) return;
  relayLog(`[${where}] seg=${id} records=${a.records} 種別=${JSON.stringify(a.byKind)} ` +
    `→ 適用=${a.applied} 削除=${a.deletes} 棄却(emb無)=${a.droppedNoEmb} 棄却(seq古)=${a.droppedOldSeq}`);
}

export interface SyncResult {
  loadedFromCache: number;
  downloaded: number;
  total: number;
}

export class VectorSync {
  /** 直近の sync 統計 (診断用)。manifest が示すセグメント数 / DL 数 / DB 件数。 */
  lastStats: { manifestSealed: number; cached: number; downloaded: number; dbSize: number } = {
    manifestSealed: -1, cached: 0, downloaded: 0, dbSize: 0,
  };

  constructor(
    private readonly store: SpVectorStore,
    private readonly cache: SegmentCache,
    private readonly db: VectorDb,
  ) {}

  /** キャッシュ適用 → manifest 差分DL → 適用。検索可能になるまで。 */
  async sync(onProgress?: (phase: 'cache' | 'manifest' | 'download', done: number, total: number) => void): Promise<SyncResult> {
    // 1. キャッシュ済みセグメントを適用
    const cachedIds = await this.cache.allIds();
    let i = 0;
    for (const id of cachedIds) {
      const seg = await this.cache.get(id);
      if (seg) { this.db.applySegment(seg); logSeg('cache', id, this.db); }
      onProgress?.('cache', ++i, cachedIds.length);
    }
    const have = new Set(cachedIds);
    relayLog(`同期: キャッシュ ${cachedIds.length}seg 適用後 DB=${this.db.size} 種別=${JSON.stringify(this.db.kindCounts())}`);

    // 2. manifest 取得 (無ければ origin 未作成 → キャッシュ分のみで終了)
    onProgress?.('manifest', 0, 0);
    const manifest = await this.store.readManifest();
    if (!manifest) {
      return { loadedFromCache: cachedIds.length, downloaded: 0, total: this.db.size };
    }

    // 3. 差分判定: 未取得の封印セグメント + 変化した open
    const toFetch = missingSealed(manifest, have);
    const prevManifest = await this.cache.getManifest();
    if (manifest.open && this.openChanged(manifest, prevManifest, have)) {
      toFetch.push(manifest.open.id);
    }
    relayLog(`同期: manifest sealed=${manifest.sealed.length}${manifest.open ? '+open' : ''} ` +
      `キャッシュ済=${have.size} 要DL=${toFetch.length}`);

    // 4. DL → 適用 → キャッシュ
    let dl = 0;
    let fetchFail = 0;
    for (const id of toFetch) {
      const seg = await this.store.readSegment(id);
      if (seg) {
        this.db.applySegment(seg);
        logSeg('DL', id, this.db);
        await this.cache.put(id, seg);
        dl++;
      } else {
        fetchFail++;
        relayLog(`同期: seg=${id} の取得に失敗 (readSegment が null)`);
      }
      onProgress?.('download', dl, toFetch.length);
    }
    if (fetchFail > 0) relayLog(`同期: DL 失敗 ${fetchFail}件`);

    // 5. コンパクション世代変化なら、manifest に無いキャッシュを掃除
    await this.pruneOrphans(manifest);
    await this.cache.setManifest(manifest);

    this.lastStats = {
      manifestSealed: manifest.sealed.length + (manifest.open ? 1 : 0),
      cached: cachedIds.length, downloaded: dl, dbSize: this.db.size,
    };
    relayLog(`同期 完了: DB=${this.db.size} 種別=${JSON.stringify(this.db.kindCounts())}`);
    return { loadedFromCache: cachedIds.length, downloaded: dl, total: this.db.size };
  }

  private openChanged(manifest: Manifest, prev: Manifest | null, have: ReadonlySet<string>): boolean {
    if (!manifest.open) return false;
    if (!have.has(manifest.open.id)) return true;             // open 未取得
    if (!prev?.open || prev.open.id !== manifest.open.id) return true;
    return prev.open.hash !== manifest.open.hash;             // 中身が変わった
  }

  /** コンパクション後など、現 manifest に存在しないキャッシュを削除。 */
  private async pruneOrphans(manifest: Manifest): Promise<void> {
    const valid = new Set<string>(manifest.sealed);
    if (manifest.open) valid.add(manifest.open.id);
    const cachedIds = await this.cache.allIds();
    for (const id of cachedIds) {
      if (!valid.has(id)) await this.cache.delete(id);
    }
  }
}
