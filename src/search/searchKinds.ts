// RAG 検索対象のソース種別を管理 (localStorage 永続化)。
//
// ユーザがチャットボックス近くのチップで「メール / OneNote / PPTX」のうち
// どれを検索対象にするか選択する。デフォルトは 3 種全部。
//
// 注: 'doc' (汎用文書) は内部利用なので UI には出さない (将来必要なら追加)。

export type SearchKind = 'mail' | 'onenote' | 'pptx' | 'transcript';

export const ALL_SEARCH_KINDS: readonly SearchKind[] = ['mail', 'onenote', 'pptx', 'transcript'];

/** UI 表示用ラベル。 */
export const SEARCH_KIND_LABELS: Record<SearchKind, string> = {
  mail:    'メール',
  onenote: 'OneNote',
  pptx:    'PPTX',
  transcript: '会議',
};

const STORAGE_KEY = 'tadori:search-kinds';

/** 現在の検索対象。空または不正値ならデフォルト (全種) を返す。 */
export function getSelectedKinds(): SearchKind[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [...ALL_SEARCH_KINDS];
    const arr = JSON.parse(raw) as unknown;
    if (!Array.isArray(arr)) return [...ALL_SEARCH_KINDS];
    const valid = arr.filter((k): k is SearchKind => ALL_SEARCH_KINDS.includes(k as SearchKind));
    // 空配列は「ユーザが意図的に全部外した状態」として尊重する (検索 0 件になる)
    return valid;
  } catch { return [...ALL_SEARCH_KINDS]; }
}

/** 検索対象を保存。空配列もそのまま保存可。 */
export function setSelectedKinds(kinds: SearchKind[]): void {
  try {
    // 順序固定 + 重複排除して保存
    const norm = ALL_SEARCH_KINDS.filter(k => kinds.includes(k));
    localStorage.setItem(STORAGE_KEY, JSON.stringify(norm));
  } catch { /* quota */ }
}

/** チップ表示用のアイコン種別名 (icons モジュールのキー)。 */
export const SEARCH_KIND_ICON: Record<SearchKind, 'message' | 'notebook' | 'presentation' | 'mic'> = {
  mail:    'message',
  onenote: 'notebook',
  pptx:    'presentation',
  transcript: 'mic',
};
