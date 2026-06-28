// doc(文書) のフォルダスコープ判定。検索時に kind='doc' レコードを、選択された
// フォルダ(serverRelativeUrl 接頭辞)配下のものだけに絞る。
//
// ★重要な不変条件: docFolderPrefixes が undefined または空配列のときは「絞り込み
//   なし = 全フォルダ対象」。空配列を「どのフォルダにも一致しない」と解釈すると
//   [].some(...) が常に false になり doc 全件が消える (実機で発生したバグ)。

/** URL エンコード差 (%20 vs スペース) と末尾スラッシュ/大文字小文字を吸収。 */
export function normUrl(s: string): string {
  let v = s.trim().replace(/\/+$/, '');
  try { v = decodeURIComponent(v); } catch { /* keep */ }
  return v.toLowerCase();
}

export interface DocScopeRecord {
  kind?: string;
  docServerRelUrl?: string;
  conversationId?: string;
}

/** docFolderPrefixes で絞り込む述語を返す。空/未指定なら常に true (全通し)。 */
export function makeDocInScope(docFolderPrefixes?: string[]): (r: DocScopeRecord) => boolean {
  const docPrefixes = (docFolderPrefixes && docFolderPrefixes.length > 0)
    ? docFolderPrefixes.map(normUrl) : null;
  return (record: DocScopeRecord): boolean => {
    if (record.kind !== 'doc' || docPrefixes == null) return true; // doc 以外 / 絞り込み無しは通す
    // docServerRelUrl が無い古いレコードは conversationId を代用。両方無ければ通す (安全側)。
    const raw = record.docServerRelUrl || record.conversationId || '';
    if (!raw) return true;
    const u = normUrl(raw);
    return docPrefixes.some(p => u === p || u.startsWith(p + '/'));
  };
}
