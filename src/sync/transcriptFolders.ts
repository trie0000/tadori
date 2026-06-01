// Teams 会議文字起こし取り込みフォルダの設定 (localStorage、サイト別)。
// 運用モデル B: SP に「会議録置き場」を 1 つ作り、.vtt を手動アップ → 同期。
//
// PPTX の pptxFolders.ts と同型。キーは tadori:transcript:folders:<siteHash>。
//
// 設計参照: docs/teams-transcript-design.md §6.3

import { siteHash } from '../sharepoint/spSites';

function keyFor(siteUrl: string): string {
  return `tadori:transcript:folders:${siteHash(siteUrl)}`;
}

export interface TranscriptFolderConfig {
  /** 表示用 URL (絶対 URL or serverRelativeUrl。入力されたまま保持)。 */
  url: string;
  /** ユーザ任意のラベル (省略時は URL 末尾セグメント)。 */
  label?: string;
  /** サブフォルダも再帰的に走査するか (会議録置き場は基本 1 階層なので既定 OFF)。 */
  recursive: boolean;
  /** 最後に同期した UNIX ms (UI 表示用。0 = 未同期)。 */
  lastSyncAt: number;
  /** 前回同期時のファイル別最終更新時刻 (filename → ISO8601)。増分判定の高速ヒント。 */
  perFile: Record<string, string>;
}

function load(siteUrl: string): TranscriptFolderConfig[] {
  try {
    const raw = localStorage.getItem(keyFor(siteUrl));
    if (!raw) return [];
    const arr = JSON.parse(raw) as TranscriptFolderConfig[];
    return Array.isArray(arr) ? arr : [];
  } catch { return []; }
}

function save(siteUrl: string, list: TranscriptFolderConfig[]): void {
  try { localStorage.setItem(keyFor(siteUrl), JSON.stringify(list)); } catch { /* quota */ }
}

export function listTranscriptFolders(siteUrl: string): TranscriptFolderConfig[] {
  return load(siteUrl);
}

export function addTranscriptFolder(siteUrl: string, cfg: Omit<TranscriptFolderConfig, 'lastSyncAt' | 'perFile'>): void {
  const list = load(siteUrl);
  const idx = list.findIndex(f => normalizeKey(f.url) === normalizeKey(cfg.url));
  if (idx >= 0) {
    list[idx] = { ...list[idx], label: cfg.label, recursive: cfg.recursive };
  } else {
    list.push({ ...cfg, lastSyncAt: 0, perFile: {} });
  }
  save(siteUrl, list);
}

export function removeTranscriptFolder(siteUrl: string, url: string): void {
  const list = load(siteUrl).filter(f => normalizeKey(f.url) !== normalizeKey(url));
  save(siteUrl, list);
}

export function updateTranscriptFolderSync(siteUrl: string, url: string, perFile: Record<string, string>): void {
  const list = load(siteUrl);
  const idx = list.findIndex(f => normalizeKey(f.url) === normalizeKey(url));
  if (idx < 0) return;
  list[idx] = { ...list[idx], lastSyncAt: Date.now(), perFile };
  save(siteUrl, list);
}

/** URL 比較用キー (末尾スラッシュ / URL エンコード差を吸収)。 */
export function normalizeKey(url: string): string {
  let s = url.trim().replace(/\/+$/, '').toLowerCase();
  try { s = decodeURIComponent(s); } catch { /* keep raw */ }
  return s;
}

/** URL から表示用ラベル (末尾セグメント) を推測。 */
export function deriveLabel(url: string): string {
  try {
    const u = url.replace(/\/+$/, '');
    const i = u.lastIndexOf('/');
    if (i < 0) return u;
    return decodeURIComponent(u.slice(i + 1));
  } catch { return url; }
}
