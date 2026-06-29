// OneNote のラベル付き取り込みバッチ (localStorage、サイト別)。
// 「ラベル＝複数ページの束」を複数管理し、チャットの「＋」ピッカーで
// ラベル単位に検索対象を絞れるようにする。キー: tadori:onenote:batches:<siteHash>

import { siteHash } from '../sharepoint/spSites';

function keyFor(siteUrl: string): string {
  return `tadori:onenote:batches:${siteHash(siteUrl)}`;
}

export interface OneNoteBatch {
  /** ユーザーが付けるラベル (検索のサブ項目名)。レコードの label と一致させる。 */
  label: string;
  /** このラベルに取り込んだ OneNote ページ ID 群。 */
  pageIds: string[];
  lastSyncAt: number;
}

function load(siteUrl: string): OneNoteBatch[] {
  try {
    const raw = localStorage.getItem(keyFor(siteUrl));
    if (!raw) return [];
    const arr = JSON.parse(raw) as OneNoteBatch[];
    return Array.isArray(arr) ? arr : [];
  } catch { return []; }
}

function save(siteUrl: string, list: OneNoteBatch[]): void {
  try { localStorage.setItem(keyFor(siteUrl), JSON.stringify(list)); } catch { /* quota */ }
}

export function listOneNoteBatches(siteUrl: string): OneNoteBatch[] {
  return load(siteUrl);
}

/** 検索ピッカー用のラベル一覧 (重複なし、登録順)。 */
export function oneNoteLabels(siteUrl: string): string[] {
  return load(siteUrl).map(b => b.label);
}

/** ラベルに pageIds を追記 (和集合)。既存ラベルならマージ、無ければ新規作成。 */
export function recordOneNoteBatch(siteUrl: string, label: string, pageIds: string[]): void {
  const lbl = label.trim();
  if (!lbl) return;
  const list = load(siteUrl);
  const idx = list.findIndex(b => b.label === lbl);
  if (idx >= 0) {
    const merged = Array.from(new Set([...list[idx].pageIds, ...pageIds]));
    list[idx] = { ...list[idx], pageIds: merged, lastSyncAt: Date.now() };
  } else {
    list.push({ label: lbl, pageIds: [...new Set(pageIds)], lastSyncAt: Date.now() });
  }
  save(siteUrl, list);
}

export function removeOneNoteBatch(siteUrl: string, label: string): void {
  save(siteUrl, load(siteUrl).filter(b => b.label !== label));
}

/** あるページ ID が属するラベルを返す (無ければ undefined)。取り込み時の label 解決に使う。 */
export function labelForPageId(siteUrl: string, pageId: string): string | undefined {
  for (const b of load(siteUrl)) if (b.pageIds.includes(pageId)) return b.label;
  return undefined;
}
