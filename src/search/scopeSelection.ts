// チャット「＋」ピッカーの検索対象サブ項目選択 (サイト別 localStorage)。
//   mail    … 選択した to/cc アドレス
//   onenote … 選択したラベル (→ pageIds に解決)
//   folders … 選択したフォルダ URL (1フォルダに pdf/pptx/docx 等が混在しても種別横断で1選択)
// 各項目とも「空 = 絞り込みなし (全件対象)」。kind 自体の ON/OFF は searchKinds 側で管理。

import { siteHash } from '../sharepoint/spSites';
import { toServerRelativeUrl } from '../sharepoint/client';
import { listOneNoteBatches } from '../sync/onenoteSources';
import type { SourceScope } from './sourceScope';
import type { SearchKind } from './searchKinds';

export interface SubSelection {
  mail: string[];
  onenote: string[];
  folders: string[];
}

const EMPTY: SubSelection = { mail: [], onenote: [], folders: [] };

function keyFor(siteUrl: string): string {
  return `tadori:source-subsel:${siteHash(siteUrl)}`;
}

export function loadSubSel(siteUrl: string): SubSelection {
  try {
    const raw = localStorage.getItem(keyFor(siteUrl));
    if (!raw) return { ...EMPTY };
    const o = JSON.parse(raw) as Partial<SubSelection> & { doc?: string[]; pptx?: string[]; transcript?: string[] };
    // 旧形式 (doc/pptx/transcript 別) は folders に統合して読み込む。
    const folders = o.folders ?? [...(o.doc ?? []), ...(o.pptx ?? []), ...(o.transcript ?? [])];
    return { mail: o.mail ?? [], onenote: o.onenote ?? [], folders: [...new Set(folders)] };
  } catch { return { ...EMPTY }; }
}

export function saveSubSel(siteUrl: string, sel: SubSelection): void {
  try { localStorage.setItem(keyFor(siteUrl), JSON.stringify(sel)); } catch { /* quota */ }
}

/** 選択ラベル群 → OneNote ページ ID 集合 (バッチ設定から解決)。 */
function labelsToPageIds(siteUrl: string, labels: string[]): string[] {
  if (labels.length === 0) return [];
  const set = new Set(labels);
  const ids = new Set<string>();
  for (const b of listOneNoteBatches(siteUrl)) if (set.has(b.label)) for (const id of b.pageIds) ids.add(id);
  return [...ids];
}

const FOLDER_KINDS: SearchKind[] = ['doc', 'pptx', 'transcript'];

/** active な kind と サブ選択から、searchVectors に渡す { kinds, scope } を組み立てる。
 *  各サブ選択が空ならその軸は絞り込みなし (全件)。
 *  フォルダを選んだ場合は doc/pptx/transcript を kinds に補う (フォルダ配下を種別横断で拾うため)。 */
export function buildScope(siteUrl: string, activeKinds: SearchKind[], sel: SubSelection): { kinds: SearchKind[]; scope: SourceScope } {
  const scope: SourceScope = {};
  if (activeKinds.includes('mail') && sel.mail.length) scope.mailAddresses = sel.mail;
  if (activeKinds.includes('onenote') && sel.onenote.length) {
    const ids = labelsToPageIds(siteUrl, sel.onenote);
    if (ids.length) scope.onenotePageIds = ids;
  }
  let kinds = activeKinds;
  if (sel.folders.length) {
    scope.folders = sel.folders.map(toServerRelativeUrl);
    // フォルダ配下の doc/pptx/transcript を kindFilter で落とさないよう補完。
    kinds = [...new Set([...activeKinds, ...FOLDER_KINDS])];
  }
  return { kinds, scope };
}
