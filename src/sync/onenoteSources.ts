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
  /** 取り込み対象のノートブック ID (配下の全セクション/ページが対象)。 */
  notebookIds: string[];
  /** 取り込み対象のセクション ID (配下の全ページが対象)。 */
  sectionIds: string[];
  /** 解決済みページ ID のスナップショット (検索の照合キー = conversationId)。
   *  ノート/セクション配下のページ + 個別選択ページの和集合。再同期で更新。 */
  pageIds: string[];
  lastSyncAt: number;
}

/** バッチへの追記内容 (いずれも和集合でマージ)。 */
export interface BatchSelection {
  pageIds?: string[];
  notebookIds?: string[];
  sectionIds?: string[];
}

function uniq(a: string[]): string[] { return [...new Set(a.filter(Boolean))]; }

function load(siteUrl: string): OneNoteBatch[] {
  try {
    const raw = localStorage.getItem(keyFor(siteUrl));
    if (!raw) return [];
    const arr = JSON.parse(raw) as Partial<OneNoteBatch>[];
    if (!Array.isArray(arr)) return [];
    // 旧形式 (pageIds のみ) も読めるよう欠損フィールドを補完。
    return arr.filter(b => b && typeof b.label === 'string').map(b => ({
      label: b.label as string,
      notebookIds: Array.isArray(b.notebookIds) ? b.notebookIds : [],
      sectionIds: Array.isArray(b.sectionIds) ? b.sectionIds : [],
      pageIds: Array.isArray(b.pageIds) ? b.pageIds : [],
      lastSyncAt: typeof b.lastSyncAt === 'number' ? b.lastSyncAt : 0,
    }));
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

/** ラベルにノート/セクション/ページを追記 (すべて和集合でマージ)。
 *  既存ラベルならマージ (＝1ラベルに複数ノート/セクションを束ねる・重複も可)、無ければ新規作成。 */
export function recordOneNoteBatch(siteUrl: string, label: string, sel: BatchSelection): void {
  const lbl = label.trim();
  if (!lbl) return;
  const list = load(siteUrl);
  const idx = list.findIndex(b => b.label === lbl);
  if (idx >= 0) {
    const cur = list[idx];
    list[idx] = {
      ...cur,
      notebookIds: uniq([...cur.notebookIds, ...(sel.notebookIds ?? [])]),
      sectionIds: uniq([...cur.sectionIds, ...(sel.sectionIds ?? [])]),
      pageIds: uniq([...cur.pageIds, ...(sel.pageIds ?? [])]),
      lastSyncAt: Date.now(),
    };
  } else {
    list.push({
      label: lbl,
      notebookIds: uniq(sel.notebookIds ?? []),
      sectionIds: uniq(sel.sectionIds ?? []),
      pageIds: uniq(sel.pageIds ?? []),
      lastSyncAt: Date.now(),
    });
  }
  save(siteUrl, list);
}

/** バッチの解決済みページ ID を丸ごと差し替え (再同期でコンテナ→現行ページを反映)。 */
export function setOneNoteBatchPageIds(siteUrl: string, label: string, pageIds: string[]): void {
  const list = load(siteUrl);
  const idx = list.findIndex(b => b.label === label);
  if (idx < 0) return;
  list[idx] = { ...list[idx], pageIds: uniq(pageIds), lastSyncAt: Date.now() };
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
