// 同期オーケストレーション。
// 起動時: IndexedDB キャッシュからセグメントを適用 → manifest を見て差分だけ
// SharePoint からDL → 適用 & キャッシュ。以降は定期ポーリングで差分追従。

import { SpVectorStore } from './spStore';
import { SegmentCache } from '../db/cache';
import { VectorDb } from '../db/store';
import { missingSealed, type Manifest } from './segments';

export interface SyncResult {
  loadedFromCache: number;
  downloaded: number;
  total: number;
}

export class VectorSync {
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
      if (seg) this.db.applySegment(seg);
      onProgress?.('cache', ++i, cachedIds.length);
    }
    const have = new Set(cachedIds);

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

    // 4. DL → 適用 → キャッシュ
    let dl = 0;
    for (const id of toFetch) {
      const seg = await this.store.readSegment(id);
      if (seg) {
        this.db.applySegment(seg);
        await this.cache.put(id, seg);
        dl++;
      }
      onProgress?.('download', dl, toFetch.length);
    }

    // 5. コンパクション世代変化なら、manifest に無いキャッシュを掃除
    await this.pruneOrphans(manifest);
    await this.cache.setManifest(manifest);

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
