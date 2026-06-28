// RAG 検索対象のソース種別を管理 (localStorage 永続化)。
//
// ユーザがチャットボックス近くのチップで「メール / OneNote / PPTX」のうち
// どれを検索対象にするか選択する。デフォルトは 3 種全部。
//
// 注: 'doc' (汎用文書) は内部利用なので UI には出さない (将来必要なら追加)。

export type SearchKind = 'mail' | 'onenote' | 'pptx' | 'transcript' | 'doc';

export const ALL_SEARCH_KINDS: readonly SearchKind[] = ['mail', 'onenote', 'pptx', 'transcript', 'doc'];

/** UI 表示用ラベル。 */
export const SEARCH_KIND_LABELS: Record<SearchKind, string> = {
  mail:    'メール',
  onenote: 'OneNote',
  pptx:    'PPTX',
  transcript: '会議',
  doc:     '文書',
};

const STORAGE_KEY = 'tadori:search-kinds';

// 「過去に存在を認識した種別」一覧。新種別 (後から追加した doc 等) を既定で
// 検索対象に含めるための判定に使う。これが無いと、種別追加前に保存された
// 選択リストに新種別が入っておらず、ユーザが明示的に外した訳でもないのに
// 検索対象から漏れてしまう。
const KNOWN_KEY = 'tadori:search-kinds-known';

/** 現在の検索対象。空または不正値ならデフォルト (全種) を返す。
 *  ★ 新しく増えた種別 (保存後に ALL_SEARCH_KINDS に追加されたもの) は、
 *    ユーザが明示的に外した訳ではないので既定で「含める」。 */
export function getSelectedKinds(): SearchKind[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [...ALL_SEARCH_KINDS];
    const arr = JSON.parse(raw) as unknown;
    if (!Array.isArray(arr)) return [...ALL_SEARCH_KINDS];
    const saved = arr.filter((k): k is SearchKind => ALL_SEARCH_KINDS.includes(k as SearchKind));

    // 既知種別リストを読む。未知 (= 後から追加された) 種別を saved に補う。
    let known: string[] = [];
    try { const k = localStorage.getItem(KNOWN_KEY); if (k) known = JSON.parse(k) as string[]; } catch { /* none */ }
    const newKinds = ALL_SEARCH_KINDS.filter(k => !known.includes(k));
    const merged = ALL_SEARCH_KINDS.filter(k => saved.includes(k) || newKinds.includes(k));

    // known を最新化して、次回以降は新種別扱いしない。
    if (newKinds.length > 0) {
      try {
        localStorage.setItem(KNOWN_KEY, JSON.stringify([...ALL_SEARCH_KINDS]));
        localStorage.setItem(STORAGE_KEY, JSON.stringify(merged));
      } catch { /* quota */ }
    }
    return merged;
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
export const SEARCH_KIND_ICON: Record<SearchKind, 'message' | 'notebook' | 'presentation' | 'mic' | 'fileText'> = {
  mail:    'message',
  onenote: 'notebook',
  pptx:    'presentation',
  transcript: 'mic',
  doc:     'fileText',
};
