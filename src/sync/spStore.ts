// ベクトルDB セグメント群の SharePoint 配置 (ファイルサーバ不要・SharePoint だけ)。
// /<site>/<library>/Tadori/ 配下に manifest.json と seg-*.json を置く。
// 読み書きはブラウザの Cookie 認証 REST (relay は SPO に認証できないため)。

import { SharePointClient } from '../sharepoint/client';
import {
  type Manifest, type Segment,
  emptyManifest, serializeManifest, parseManifest, serializeSegment, parseSegment, segmentId,
} from './segments';

const MANIFEST_NAME = 'manifest.json';
const DEFAULT_LIBRARY = 'Shared Documents'; // SPO 既定ライブラリの server-relative 名

export class SpVectorStore {
  private readonly folder: string; // site 込み server-relative (例: /sites/n365/Shared Documents/Tadori)

  constructor(private readonly sp: SharePointClient, library = DEFAULT_LIBRARY, sub = 'Tadori') {
    this.folder = `${sp.serverRelativeSite()}/${library}/${sub}`;
  }

  /** 配置フォルダを作成 (冪等)。 */
  async ensure(): Promise<void> {
    await this.sp.ensureFolder(this.folder);
  }

  async readManifest(): Promise<Manifest | null> {
    const text = await this.sp.readFileText(`${this.folder}/${MANIFEST_NAME}`);
    if (text == null) return null;
    return parseManifest(text);
  }

  /** ETag 付きで manifest を読む (CAS 用)。manifest が無ければ null。 */
  async readManifestWithEtag(): Promise<{ manifest: Manifest; etag: string } | null> {
    const r = await this.sp.readFileTextWithEtag(`${this.folder}/${MANIFEST_NAME}`);
    if (!r) return null;
    return { manifest: parseManifest(r.text), etag: r.etag };
  }

  async writeManifest(m: Manifest): Promise<void> {
    await this.sp.uploadFileText(this.folder, MANIFEST_NAME, serializeManifest(m));
  }

  /** manifest を CAS で上書き (PUT + If-Match)。etag が空文字なら通常の overwrite で書く (初回)。 */
  async writeManifestCas(m: Manifest, etag: string): Promise<void> {
    if (!etag) {
      await this.sp.uploadFileText(this.folder, MANIFEST_NAME, serializeManifest(m));
      return;
    }
    await this.sp.uploadFileTextCas(`${this.folder}/${MANIFEST_NAME}`, serializeManifest(m), etag);
  }

  async readSegment(id: string): Promise<Segment | null> {
    const text = await this.sp.readFileText(`${this.folder}/${id}.json`);
    if (text == null) return null;
    return parseSegment(text);
  }

  async writeSegment(seg: Segment): Promise<void> {
    await this.sp.uploadFileText(this.folder, `${seg.id}.json`, serializeSegment(seg));
  }

  /** 既存ファイルがあれば idx を bump して衝突回避しながら書く。確定した segment id を返す。
   *  他 writer が同じ idx で書こうとした場合の名前衝突 (両方が同じ seg-NNNNN.json を書こうとする)
   *  を防ぐ。Phase 1 で必須。 */
  async writeSegmentNoOverwrite(seg: Segment, startIdx: number, maxAttempts = 50): Promise<{ id: string; idx: number }> {
    let idx = startIdx;
    for (let i = 0; i < maxAttempts; i++) {
      const id = segmentId(idx);
      const ok = await this.sp.uploadFileTextNoOverwrite(this.folder, `${id}.json`, serializeSegment({ ...seg, id }));
      if (ok) return { id, idx };
      idx++; // 衝突 → 次番号で再試行
    }
    throw new Error(`segment id 衝突が ${maxAttempts} 回連続: 試行回数超過`);
  }

  /** 既存セグメント id 一覧 (manifest 不整合の復旧用)。 */
  async listSegmentIds(): Promise<string[]> {
    const names = await this.sp.listFolderFileNames(this.folder);
    return names.filter(n => n.startsWith('seg-') && n.endsWith('.json')).map(n => n.slice(0, -5));
  }

  /** SharePoint 上の Tadori フォルダ配下の manifest と全セグメントを削除。
   *  「取り込みメールの全削除」用。フォルダ自体は残す (次回取り込みで再利用)。 */
  async deleteAll(): Promise<void> {
    const names = await this.sp.listFolderFileNames(this.folder);
    for (const name of names) {
      if (name === MANIFEST_NAME || (name.startsWith('seg-') && name.endsWith('.json'))) {
        await this.sp.deleteFile(`${this.folder}/${name}`).catch(() => { /* best-effort */ });
      }
    }
  }

  /** 初回・manifest 不在時は空 manifest を作って配置する。 */
  async ensureManifest(): Promise<Manifest> {
    const m = await this.readManifest();
    if (m) return m;
    const fresh = emptyManifest();
    await this.ensure();
    await this.writeManifest(fresh);
    return fresh;
  }
}

/** manifest 更新の差分を表す patcher 関数。最新の manifest を受け取って、
 *  自分の変更 (seg 追加 / maxSeq 更新等) を適用した新 manifest を返す。
 *  CAS 失敗時は再読込された最新 manifest で再呼出されるため、毎回 idempotent に動作すること
 *  (例: sealed に id を「無ければ追加」、maxSeq は max を取る、version は +1 など)。 */
export type ManifestPatcher = (current: Manifest) => Manifest;

/** Manifest を CAS で更新。412 競合時は最新を再読込して patcher を再適用してリトライ。
 *  Tadori の分散書き込みの正確性を担保する最終防衛線。lease が機能していれば
 *  通常はリトライ無しで即成功する。 */
export async function updateManifestWithCas(
  store: SpVectorStore,
  patcher: ManifestPatcher,
  maxRetries = 5,
): Promise<Manifest> {
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const cur = await store.readManifestWithEtag();
    const base = cur?.manifest ?? emptyManifest();
    const etag = cur?.etag ?? '';
    const next = patcher(base);
    next.updatedAt = new Date().toISOString();
    try {
      await store.writeManifestCas(next, etag);
      return next;
    } catch (e) {
      const err = e as Error & { code?: string };
      if (err.code !== 'PRECONDITION_FAILED' || attempt === maxRetries) throw e;
      // jittered backoff (50〜250ms)
      await new Promise(r => setTimeout(r, 50 + Math.random() * 200));
    }
  }
  throw new Error('manifest CAS: max retries exceeded');
}
