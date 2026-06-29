// 社内用語・略語辞書 (同義語グループ)。検索の「クエリ展開」に使う。
// 実体は SP の Tadori フォルダに glossary.json (チーム共有) を置き、localStorage に
// キャッシュ (検索時は同期で引けるように)。意味(def)は任意 — 検索ヒット率を上げるだけなら
// 同義語の集合だけで十分。def は「用語の意味を問う質問」用の擬似ソースに使える(任意)。

import { SharePointClient } from '../sharepoint/client';
import { siteHash } from '../sharepoint/spSites';

export interface GlossaryEntry {
  /** 代表名 (表示・正規化のアンカー)。 */
  canonical: string;
  /** 別名・略語・表記ゆれ。canonical と合わせて「等価な表記の集合」。 */
  aliases: string[];
  /** 任意の定義 (空でよい)。 */
  def?: string;
}

const LIB = 'Shared Documents';
const SUB = 'Tadori';
const FILE = 'glossary.json';

function cacheKey(siteUrl: string): string { return `tadori:glossary:${siteHash(siteUrl)}`; }
function folderOf(sp: SharePointClient): string { return `${sp.serverRelativeSite()}/${LIB}/${SUB}`; }

function sanitize(arr: unknown): GlossaryEntry[] {
  if (!Array.isArray(arr)) return [];
  const out: GlossaryEntry[] = [];
  for (const e of arr) {
    if (!e || typeof e !== 'object') continue;
    const o = e as Record<string, unknown>;
    const canonical = typeof o.canonical === 'string' ? o.canonical.trim() : '';
    const aliases = Array.isArray(o.aliases) ? o.aliases.filter((x): x is string => typeof x === 'string').map(s => s.trim()).filter(Boolean) : [];
    const def = typeof o.def === 'string' ? o.def : undefined;
    if (!canonical && aliases.length === 0) continue;
    out.push({ canonical, aliases, def: def || undefined });
  }
  return out;
}

/** localStorage キャッシュから同期で取得 (検索時用)。 */
export function loadGlossary(siteUrl: string): GlossaryEntry[] {
  try {
    const raw = localStorage.getItem(cacheKey(siteUrl));
    return raw ? sanitize(JSON.parse(raw)) : [];
  } catch { return []; }
}

function saveCache(siteUrl: string, entries: GlossaryEntry[]): void {
  try { localStorage.setItem(cacheKey(siteUrl), JSON.stringify(entries)); } catch { /* quota */ }
}

/** SP から glossary.json を読み込み、キャッシュ更新して返す (起動時 / 設定画面表示時)。 */
export async function fetchGlossary(siteUrl: string): Promise<GlossaryEntry[]> {
  const sp = new SharePointClient(siteUrl);
  try {
    const text = await sp.readFileText(`${folderOf(sp)}/${FILE}`);
    const entries = text ? sanitize(JSON.parse(text)) : [];
    saveCache(siteUrl, entries);
    return entries;
  } catch {
    return loadGlossary(siteUrl); // SP 不達ならキャッシュで代替
  }
}

/** SP へ保存 + キャッシュ更新。 */
export async function persistGlossary(siteUrl: string, entries: GlossaryEntry[]): Promise<void> {
  const clean = sanitize(entries);
  saveCache(siteUrl, clean);
  const sp = new SharePointClient(siteUrl);
  await sp.ensureFolder(folderOf(sp));
  await sp.uploadFileText(folderOf(sp), FILE, JSON.stringify(clean, null, 2));
}

/** Excel からの貼り付け (TSV/CSV) をエントリ配列に。1列目=代表名, 2列目=別名(カンマ/セミコロン区切り), 3列目=意味(任意)。 */
export function parseGlossaryTable(text: string): GlossaryEntry[] {
  const out: GlossaryEntry[] = [];
  for (const lineRaw of text.replace(/\r\n/g, '\n').split('\n')) {
    const line = lineRaw.trim();
    if (!line) continue;
    const cols = line.split(line.includes('\t') ? '\t' : ',').map(c => c.trim());
    const canonical = cols[0] || '';
    const aliases = (cols[1] || '').split(/[,;、；]/).map(s => s.trim()).filter(Boolean);
    const def = cols[2] || undefined;
    // 見出し行 ("正式名"等) はスキップ
    if (/^(正式名|canonical|用語|term)$/i.test(canonical)) continue;
    if (!canonical && aliases.length === 0) continue;
    out.push({ canonical, aliases, def });
  }
  return out;
}

function norm(s: string): string { return s.trim().toLowerCase(); }

/** クエリ展開: query に含まれる用語の「同義の他表記」を返す (検索語に足す用)。
 *  - 各エントリは [canonical, ...aliases] を等価グループとみなす。
 *  - グループ内のいずれかの表記が query に出現したら、グループの他表記を候補に追加。
 *  - 既に query に含まれる語は除外。maxAdds で上限 (展開しすぎによる精度低下を防ぐ)。 */
export function expandQueryTerms(query: string, entries: GlossaryEntry[], maxAdds = 8): string[] {
  const q = norm(query);
  if (!q) return [];
  const added = new Set<string>();
  for (const e of entries) {
    const group = [e.canonical, ...e.aliases].filter(Boolean);
    const hit = group.some(form => form.length >= 2 && q.includes(norm(form)));
    if (!hit) continue;
    for (const form of group) {
      if (form.length < 2) continue;
      if (q.includes(norm(form))) continue;     // 既出は足さない
      if (added.has(form)) continue;
      added.add(form);
      if (added.size >= maxAdds) return [...added];
    }
  }
  return [...added];
}
