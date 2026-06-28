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
  if (idx >= 0) list[idx] = { ...list[idx], label: cfg.label, recursive: cfg.recursive };
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

export function deriveLabel(url: string): string {
  try {
    const u = url.replace(/\/+$/, '');
    const i = u.lastIndexOf('/');
    return i < 0 ? u : decodeURIComponent(u.slice(i + 1));
  } catch { return url; }
}
