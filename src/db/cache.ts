// IndexedDB によるセグメントキャッシュ。
// 起動のたびに SharePoint から全DLしないよう、DL済みセグメント(JSON)と最後に
// 見た manifest をブラウザにローカル保存する。次回は差分だけ取得すればよい。

import { type Segment, type Manifest, parseSegment, serializeSegment, parseManifest, serializeManifest } from '../sync/segments';

const DB_NAME = 'tadori';
const DB_VERSION = 1;
const STORE_SEG = 'segments';   // key = segment id, value = JSON text
const STORE_META = 'meta';      // key = 'manifest', value = JSON text

function open(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_SEG)) db.createObjectStore(STORE_SEG);
      if (!db.objectStoreNames.contains(STORE_META)) db.createObjectStore(STORE_META);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function tx<T>(db: IDBDatabase, store: string, mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    const t = db.transaction(store, mode);
    const req = fn(t.objectStore(store));
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export class SegmentCache {
  private dbp: Promise<IDBDatabase> | null = null;
  private db(): Promise<IDBDatabase> { return (this.dbp ??= open()); }

  async allIds(): Promise<string[]> {
    const db = await this.db();
    const keys = await tx<IDBValidKey[]>(db, STORE_SEG, 'readonly', s => s.getAllKeys());
    return keys.map(String);
  }

  async get(id: string): Promise<Segment | null> {
    const db = await this.db();
    const text = await tx<string | undefined>(db, STORE_SEG, 'readonly', s => s.get(id));
    return text ? parseSegment(text) : null;
  }

  async put(id: string, seg: Segment): Promise<void> {
    const db = await this.db();
    await tx(db, STORE_SEG, 'readwrite', s => s.put(serializeSegment(seg), id));
  }

  async delete(id: string): Promise<void> {
    const db = await this.db();
    await tx(db, STORE_SEG, 'readwrite', s => s.delete(id));
  }

  async getManifest(): Promise<Manifest | null> {
    const db = await this.db();
    const text = await tx<string | undefined>(db, STORE_META, 'readonly', s => s.get('manifest'));
    return text ? parseManifest(text) : null;
  }

  async setManifest(m: Manifest): Promise<void> {
    const db = await this.db();
    await tx(db, STORE_META, 'readwrite', s => s.put(serializeManifest(m), 'manifest'));
  }

  /** すべてのセグメントと manifest を消す (取り込み済みメールの全削除用)。 */
  async clearAll(): Promise<void> {
    const db = await this.db();
    await tx(db, STORE_SEG, 'readwrite', s => s.clear());
    await tx(db, STORE_META, 'readwrite', s => s.clear());
  }
}
