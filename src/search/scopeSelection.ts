// チャット「＋」ピッカーの検索対象サブ項目選択 (サイト別 localStorage)。
//   mail       … 選択した to/cc アドレス
//   onenote    … 選択したラベル (→ pageIds に解決)
//   doc/pptx/transcript … 選択したフォルダ URL
// 各種別とも「空 = その種別は全件対象」。kind 自体の ON/OFF は searchKinds 側で管理。

import { siteHash } from '../sharepoint/spSites';
import { toServerRelativeUrl } from '../sharepoint/client';
import { listOneNoteBatches } from '../sync/onenoteSources';
import type { SourceScope } from './sourceScope';
import type { SearchKind } from './searchKinds';

export interface SubSelection {
  mail: string[];
  onenote: string[];
  doc: string[];
  pptx: string[];
  transcript: string[];
}

const EMPTY: SubSelection = { mail: [], onenote: [], doc: [], pptx: [], transcript: [] };

function keyFor(siteUrl: string): string {
  return `tadori:source-subsel:${siteHash(siteUrl)}`;
}

export function loadSubSel(siteUrl: string): SubSelection {
  try {
    const raw = localStorage.getItem(keyFor(siteUrl));
    if (!raw) return { ...EMPTY };
    const o = JSON.parse(raw) as Partial<SubSelection>;
    return {
      mail: o.mail ?? [], onenote: o.onenote ?? [], doc: o.doc ?? [],
      pptx: o.pptx ?? [], transcript: o.transcript ?? [],
    };
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

/** active な kind と サブ選択から、searchVectors に渡す { kinds, scope } を組み立てる。
 *  各種別のサブ選択が空なら、その種別は絞り込みなし (全件)。 */
export function buildScope(siteUrl: string, activeKinds: SearchKind[], sel: SubSelection): { kinds: SearchKind[]; scope: SourceScope } {
  const scope: SourceScope = {};
  if (activeKinds.includes('mail') && sel.mail.length) scope.mailAddresses = sel.mail;
  if (activeKinds.includes('onenote') && sel.onenote.length) {
    const ids = labelsToPageIds(siteUrl, sel.onenote);
    if (ids.length) scope.onenotePageIds = ids;
  }
  if (activeKinds.includes('doc') && sel.doc.length) scope.docFolders = sel.doc.map(toServerRelativeUrl);
  if (activeKinds.includes('pptx') && sel.pptx.length) scope.pptxFolders = sel.pptx.map(toServerRelativeUrl);
  if (activeKinds.includes('transcript') && sel.transcript.length) scope.transcriptFolders = sel.transcript.map(toServerRelativeUrl);
  return { kinds: activeKinds, scope };
}
