// 検索の「種別ごとのサブ項目」絞り込み。チャットの「＋」ピッカーで選んだ小項目を
// レコードに照合する述語を作る。
//
// 不変条件: ある種別のサブ選択が「空/未指定」なら、その種別は全件対象 (絞り込みなし)。
//   空配列を「どれにも一致しない」と解釈して全消しする不具合 (doc スコープで発生) を
//   各種別で繰り返さないため、length===0 は「絞り込みなし」とする。
//
// 照合キー:
//   mail       … to/cc に含まれるアドレス (mailAddresses)
//   onenote    … 取り込みバッチの label (onenoteLabels)
//   doc        … docServerRelUrl の接頭辞 (docFolders)
//   pptx       … pptxServerRelUrl の接頭辞 (pptxFolders)
//   transcript … vttServerRelUrl の接頭辞 (transcriptFolders)

export interface SourceScope {
  mailAddresses?: string[];
  onenoteLabels?: string[];
  docFolders?: string[];
  pptxFolders?: string[];
  transcriptFolders?: string[];
}

export interface ScopeRecord {
  kind?: string;
  label?: string;
  to?: string[];
  cc?: string[];
  docServerRelUrl?: string;
  pptxServerRelUrl?: string;
  vttServerRelUrl?: string;
  conversationId?: string;
}

/** URL エンコード差 (%20 vs スペース)・末尾スラッシュ・大文字小文字を吸収。 */
export function normUrl(s: string): string {
  let v = s.trim().replace(/\/+$/, '');
  try { v = decodeURIComponent(v); } catch { /* keep */ }
  return v.toLowerCase();
}

function nonEmpty(a?: string[]): string[] | null {
  return a && a.length > 0 ? a : null;
}

function prefixMatch(raw: string | undefined, fallback: string | undefined, prefixes: string[]): boolean {
  const src = raw || fallback || '';
  if (!src) return true; // 接頭辞を持たない古いレコードは安全側で通す
  const u = normUrl(src);
  return prefixes.some(p => { const np = normUrl(p); return u === np || u.startsWith(np + '/'); });
}

/** scope に基づき「このレコードが検索対象か」を返す述語。種別未選択は全通し。 */
export function makeInScope(scope?: SourceScope): (r: ScopeRecord) => boolean {
  const mail = nonEmpty(scope?.mailAddresses)?.map(a => a.trim().toLowerCase());
  const note = nonEmpty(scope?.onenoteLabels);
  const docs = nonEmpty(scope?.docFolders);
  const ppts = nonEmpty(scope?.pptxFolders);
  const trns = nonEmpty(scope?.transcriptFolders);

  return (r: ScopeRecord): boolean => {
    switch (r.kind) {
      case 'mail': {
        if (!mail) return true;
        const addrs = [...(r.to ?? []), ...(r.cc ?? [])].map(a => a.toLowerCase());
        return addrs.some(a => mail.some(sel => a.includes(sel)));
      }
      case 'onenote':
        if (!note) return true;
        return r.label != null && note.includes(r.label);
      case 'doc':
        if (!docs) return true;
        return prefixMatch(r.docServerRelUrl, r.conversationId, docs);
      case 'pptx':
        if (!ppts) return true;
        return prefixMatch(r.pptxServerRelUrl, r.conversationId, ppts);
      case 'transcript':
        if (!trns) return true;
        return prefixMatch(r.vttServerRelUrl, r.conversationId, trns);
      default:
        return true;
    }
  };
}
