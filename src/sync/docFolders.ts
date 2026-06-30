// ドキュメント取り込みフォルダの設定 (localStorage、サイト別)。
// docx / doc / pdf / md / txt を SP フォルダから取り込む。
// pptx/transcript の Folders と同型。キー: tadori:doc:folders:<siteHash>

import { siteHash } from '../sharepoint/spSites';

function keyFor(siteUrl: string): string {
  return `tadori:doc:folders:${siteHash(siteUrl)}`;
}

export interface DocFolderConfig {
  url: string;
  label?: string;
  recursive: boolean;
  /** このフォルダの .pptx を Vision 解析するか (既定 false = テキストのみ)。 */
  visionForPptx?: boolean;
  lastSyncAt: number;
  perFile: Record<string, string>;
}

function load(siteUrl: string): DocFolderConfig[] {
  try {
    const raw = localStorage.getItem(keyFor(siteUrl));
    if (!raw) return [];
    const arr = JSON.parse(raw) as DocFolderConfig[];
    return Array.isArray(arr) ? arr : [];
  } catch { return []; }
}

function save(siteUrl: string, list: DocFolderConfig[]): void {
  try { localStorage.setItem(keyFor(siteUrl), JSON.stringify(list)); } catch { /* quota */ }
}

export function listDocFolders(siteUrl: string): DocFolderConfig[] {
  return load(siteUrl);
}

export function addDocFolder(siteUrl: string, cfg: Omit<DocFolderConfig, 'lastSyncAt' | 'perFile'>): void {
  const list = load(siteUrl);
  const idx = list.findIndex(f => normalizeKey(f.url) === normalizeKey(cfg.url));
  if (idx >= 0) list[idx] = { ...list[idx], label: cfg.label, recursive: cfg.recursive, visionForPptx: cfg.visionForPptx };
  else list.push({ ...cfg, lastSyncAt: 0, perFile: {} });
  save(siteUrl, list);
}

export function removeDocFolder(siteUrl: string, url: string): void {
  save(siteUrl, load(siteUrl).filter(f => normalizeKey(f.url) !== normalizeKey(url)));
}

export function updateDocFolderSync(siteUrl: string, url: string, perFile: Record<string, string>): void {
  const list = load(siteUrl);
  const idx = list.findIndex(f => normalizeKey(f.url) === normalizeKey(url));
  if (idx < 0) return;
  list[idx] = { ...list[idx], lastSyncAt: Date.now(), perFile };
  save(siteUrl, list);
}

export function normalizeKey(url: string): string {
  let s = url.trim().replace(/\/+$/, '').toLowerCase();
  try { s = decodeURIComponent(s); } catch { /* keep raw */ }
  return s;
}

// ─── 検索スコープ (チャット側: どのフォルダを検索対象に含めるか) ───────────────
// 設定で取り込んだフォルダのうち、検索に含めるものの URL 集合を保持。
// 未設定 (キー無し) は「全フォルダ対象」を意味する (デフォルト挙動)。

function scopeKeyFor(siteUrl: string): string {
  return `tadori:doc:search-scope:${siteHash(siteUrl)}`;
}

/** 検索対象に含めるフォルダ URL 一覧。未設定なら null (= 全部対象)。 */
export function getDocSearchScope(siteUrl: string): string[] | null {
  try {
    const raw = localStorage.getItem(scopeKeyFor(siteUrl));
    if (raw == null) return null;
    const arr = JSON.parse(raw) as unknown;
    return Array.isArray(arr) ? arr.filter((x): x is string => typeof x === 'string') : null;
  } catch { return null; }
}

export function setDocSearchScope(siteUrl: string, urls: string[]): void {
  try { localStorage.setItem(scopeKeyFor(siteUrl), JSON.stringify(urls)); } catch { /* quota */ }
}

/** 実効スコープ: スコープ未設定なら全登録フォルダ、設定済みなら登録済みと交差。 */
export function effectiveDocScope(siteUrl: string): string[] {
  const all = listDocFolders(siteUrl).map(f => f.url);
  const scope = getDocSearchScope(siteUrl);
  if (scope == null) return all;                       // 未設定 = 全部
  const set = new Set(scope.map(normalizeKey));
  return all.filter(u => set.has(normalizeKey(u)));    // 登録済み ∩ スコープ
}

export function deriveLabel(url: string): string {
  try {
    const u = url.replace(/\/+$/, '');
    const i = u.lastIndexOf('/');
    return i < 0 ? u : decodeURIComponent(u.slice(i + 1));
  } catch { return url; }
}
