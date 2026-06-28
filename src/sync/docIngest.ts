// ドキュメント (docx / doc / pdf / md / txt) 取り込みパイプライン。
//
// フロー (1 フォルダ):
//   1. SP REST でフォルダ配下のファイル一覧 (対応拡張子でフィルタ)
//   2. 増分判定 (lastModified) + 削除検知
//   3. 各ファイル:
//        a. fetchFileBytes
//        b. md/txt → ブラウザで UTF-8 デコード
//           docx/doc/pdf/rtf → relay /tadori/doc-extract で Word COM 抽出
//        c. splitIntoChunks
//        d. IngestMail[] (kind='doc') → ingestToSegments
//   4. perFile 更新
//
// 設計参照: ユーザ要望 (docx/doc/pdf/md/txt をフォルダ指定で検索対象に)

import { SharePointClient, toServerRelativeUrl, type FileInfo } from '../sharepoint/client';
import { splitIntoChunks } from '../lib/chunk';
import { ingestToSegments, deleteFromSegments, type IngestMail } from '../db/writer';
import { getEngine } from '../db/engine';
import type { RuntimeSettings } from '../api/aiSettings';
import { updateDocFolderSync, type DocFolderConfig } from './docFolders';
import { docxToText, xlsxToText } from '../docs/docxText';
import { pdfToText } from '../docs/pdfText';

// すべてブラウザ内でパースする (Word/Excel/relay 不要)。
//   .md/.txt → TextDecoder
//   .docx/.xlsx → ネイティブ ZIP 解凍 + XML 抽出 (docxText.ts)
//   .pdf → pdf.js (pdfText.ts)
// .doc (旧バイナリ) / .rtf は専用フォーマットなので現状未対応 (UNSUPPORTED で通知)。
const TEXT_EXT = new Set(['.md', '.markdown', '.txt']);
const DOCX_EXT = new Set(['.docx']);
const XLSX_EXT = new Set(['.xlsx']);
const PDF_EXT = new Set(['.pdf']);
const ALL_EXT = new Set([...TEXT_EXT, ...DOCX_EXT, ...XLSX_EXT, ...PDF_EXT]);

export interface DocIngestProgress {
  file: string;
  fileIdx: number;
  fileTotal: number;
  chunkIdx: number;
  chunkTotal: number;
  phase: 'fetch' | 'parse' | 'embed' | 'done' | 'skip' | 'error' | 'delete';
  message?: string;
}

export interface DocIngestResult {
  ingestedFiles: number;
  ingestedChunks: number;
  skippedFiles: number;
  deletedFiles: number;
  failedFiles: number;
}

function siteCollectionPath(serverRel: string): string {
  const m = serverRel.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
  return m ? m[1] : '';
}

function resolveSpFolder(folderUrl: string, fallbackSiteUrl: string): { siteUrl: string; folderServerRel: string } {
  const trimmed = folderUrl.trim();
  if (!trimmed) throw new Error('フォルダ URL が空です');
  const folderServerRel = toServerRelativeUrl(trimmed);
  let origin = '';
  try { const u = new URL(trimmed); origin = `${u.protocol}//${u.host}`; }
  catch {
    try { const fb = new URL(fallbackSiteUrl); origin = `${fb.protocol}//${fb.host}`; }
    catch { origin = fallbackSiteUrl.replace(/\/+$/, '').replace(/\/_api\/.*$/, ''); }
  }
  const siteUrl = origin + siteCollectionPath(folderServerRel);
  if (!folderServerRel) throw new Error(`フォルダ URL の serverRelativeUrl を解釈できませんでした: ${folderUrl}`);
  return { siteUrl, folderServerRel };
}

function extOf(name: string): string {
  const i = name.lastIndexOf('.');
  return i < 0 ? '' : name.slice(i).toLowerCase();
}

function filterDocFiles(items: FileInfo[]): FileInfo[] {
  return items.filter(f => {
    const n = f.name.toLowerCase();
    if (n.startsWith('~$') || n.startsWith('.')) return false;
    return ALL_EXT.has(extOf(n));
  });
}

/** ファイルバイト列をテキスト化。docx/xlsx/md/txt はブラウザ、pdf は relay (PdfPig)。 */
async function extractText(ext: string, bytes: ArrayBuffer, relayBaseUrl: string, fileName: string, signal?: AbortSignal): Promise<string> {
  if (TEXT_EXT.has(ext)) return new TextDecoder('utf-8').decode(bytes).replace(/^﻿/, '');
  if (DOCX_EXT.has(ext)) return docxToText(bytes);
  if (XLSX_EXT.has(ext)) return xlsxToText(bytes);
  if (PDF_EXT.has(ext)) return pdfToText(bytes, relayBaseUrl, fileName, signal);
  throw new Error(`未対応の形式: ${ext}`);
}

function pickTargets(now: FileInfo[], prev: Record<string, string>, force = false, targetFiles?: ReadonlySet<string>): {
  toIngest: FileInfo[]; skipped: FileInfo[]; deleted: string[];
} {
  const toIngest: FileInfo[] = [];
  const skipped: FileInfo[] = [];
  const nowNames = new Set<string>();
  for (const f of now) {
    nowNames.add(f.name);
    if (targetFiles && !targetFiles.has(f.name)) { skipped.push(f); continue; }
    if (force) { toIngest.push(f); continue; }
    const prevTs = prev[f.name];
    if (!prevTs || prevTs !== f.timeLastModified) toIngest.push(f);
    else skipped.push(f);
  }
  const deleted: string[] = [];
  if (!targetFiles) for (const name of Object.keys(prev)) if (!nowNames.has(name)) deleted.push(name);
  return { toIngest, skipped, deleted };
}

function hashStr(s: string): string {
  let h = 5381;
  for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}

/** SP のファイルをブラウザで開く (Office Online / 既定ビューア)。 */
export function openDocSource(hit: { docServerRelUrl?: string }): void {
  if (!hit.docServerRelUrl) return;
  const abs = hit.docServerRelUrl.startsWith('http') ? hit.docServerRelUrl : `${location.origin}${hit.docServerRelUrl}`;
  window.open(abs, '_blank', 'noopener');
}

export async function syncDocFolder(
  folder: DocFolderConfig,
  s: RuntimeSettings,
  fallbackSiteUrl: string,
  onProgress?: (p: DocIngestProgress) => void,
  signal?: AbortSignal,
  opts: { force?: boolean; targetFiles?: ReadonlySet<string> } = {},
): Promise<DocIngestResult> {
  const { siteUrl, folderServerRel } = resolveSpFolder(folder.url, fallbackSiteUrl);
  const sp = new SharePointClient(siteUrl);
  console.log('[tadori] doc sync start', { inputUrl: folder.url, siteUrl, folderServerRel });

  onProgress?.({ file: '', fileIdx: 0, fileTotal: 0, chunkIdx: 0, chunkTotal: 0, phase: 'fetch', message: `フォルダ一覧を取得中… (${folderServerRel})` });
  const items = await sp.listFolderItems(folderServerRel, { recursive: folder.recursive });
  const docFiles = filterDocFiles(items);
  console.log(`[tadori] doc sync: ${items.length} items / ${docFiles.length} doc files`);
  if (docFiles.length === 0) {
    onProgress?.({
      file: '', fileIdx: 0, fileTotal: 0, chunkIdx: 0, chunkTotal: 0, phase: 'skip',
      message: items.length === 0
        ? `フォルダにファイルが見つかりません (${folderServerRel})。URL/権限を確認してください。`
        : `${items.length} 個のファイルがありますが対応文書 (docx/doc/pdf/md/txt) は 0 件でした。`,
    });
  }

  const { toIngest, skipped, deleted } = pickTargets(docFiles, folder.perFile, opts.force === true, opts.targetFiles);

  let deletedFiles = 0;
  if (!opts.targetFiles && deleted.length > 0) {
    const eng = await getEngine(siteUrl);
    for (const fname of deleted) {
      if (signal?.aborted) break;
      onProgress?.({ file: fname, fileIdx: 0, fileTotal: 0, chunkIdx: 0, chunkTotal: 0, phase: 'delete', message: `${fname} は SP から消えていた — chunk を削除` });
      const stale = `${folderServerRel}/${fname}`;
      for (const mid of eng.db.messageIdsForConversation(stale)) {
        try { await deleteFromSegments(mid, siteUrl); } catch (e) { console.warn('[doc] delete chunk failed:', mid, (e as Error).message); }
      }
      deletedFiles++;
    }
  }

  let ingestedFiles = 0, ingestedChunks = 0, failedFiles = 0;
  const newPerFile: Record<string, string> = {};
  if (opts.targetFiles) for (const [name, ts] of Object.entries(folder.perFile)) newPerFile[name] = ts;
  else for (const f of skipped) newPerFile[f.name] = f.timeLastModified;

  for (let i = 0; i < toIngest.length; i++) {
    if (signal?.aborted) break;
    const f = toIngest[i];
    const fileIdx = i + 1, fileTotal = toIngest.length;
    const e = extOf(f.name);
    try {
      onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: 0, chunkTotal: 0, phase: 'fetch', message: `${f.name} をダウンロード中…` });
      const bytes = await sp.fetchFileBytes(f.serverRelativeUrl);
      if (!bytes) {
        onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: 0, chunkTotal: 0, phase: 'skip', message: `${f.name} が取得できませんでした` });
        continue;
      }

      onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: 0, chunkTotal: 0, phase: 'parse', message: `${f.name} を解析中…` });
      // すべてブラウザ内でパース (Word/relay 不要)。
      console.log(`[tadori] doc parse: ${f.name} (${(bytes.byteLength / 1024).toFixed(0)} KB, ${e})`);
      let text = (await extractText(e, bytes, s.relayBaseUrl, f.name, signal)).trim();
      console.log(`[tadori] doc parsed: ${f.name} → ${text.length} chars`);
      if (!text) {
        onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: 0, chunkTotal: 0, phase: 'skip', message: `${f.name} からテキストを抽出できませんでした` });
        newPerFile[f.name] = f.timeLastModified;
        continue;
      }

      const chunks = splitIntoChunks(text, { maxChars: 800, overlap: 80 });
      if (chunks.length === 0) { newPerFile[f.name] = f.timeLastModified; continue; }

      const baseName = f.name.replace(/\.[^.]+$/, '');
      const mails: IngestMail[] = chunks.map((c, idx) => {
        const docPath = `doc://${f.serverRelativeUrl}#${idx}`;
        return {
          messageId: docPath,
          internetMessageId: '',
          conversationId: f.serverRelativeUrl,    // 同一ファイルの全チャンクを束ねる
          kind: 'doc',
          chunkIdx: idx,
          chunkCount: chunks.length,
          docPath,
          subject: c.heading ? `${baseName} — ${c.heading}` : baseName,
          from: f.name,
          to: [],
          cc: [],
          date: f.timeLastModified || new Date().toISOString(),
          body: c.text,
          isHtml: false,
          docFile: f.name,
          docServerRelUrl: f.serverRelativeUrl,
          srcHash: hashStr(c.text),
        };
      });

      // 既存 chunk のうち新版に無い index を削除 (ファイルが短くなった場合)
      try {
        const eng = await getEngine(siteUrl);
        const willHave = new Set(mails.map(m => m.messageId));
        for (const mid of eng.db.messageIdsForConversation(f.serverRelativeUrl)) {
          if (!willHave.has(mid)) { try { await deleteFromSegments(mid, siteUrl); } catch { /* best-effort */ } }
        }
      } catch { /* 初回は engine 空 */ }

      onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: chunks.length, chunkTotal: chunks.length, phase: 'embed', message: `${f.name} を埋め込み中… (${chunks.length} chunk)` });
      await ingestToSegments(mails, s, siteUrl, undefined, signal);
      ingestedChunks += mails.length;
      ingestedFiles++;
      newPerFile[f.name] = f.timeLastModified;
      onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: chunks.length, chunkTotal: chunks.length, phase: 'done', message: `${f.name} 完了 (${chunks.length} chunk)` });
    } catch (err) {
      if ((err as Error).name === 'AbortError') throw err;
      console.error('[doc] file failed:', f.name, err);
      failedFiles++;
      onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: 0, chunkTotal: 0, phase: 'error', message: `${f.name} 失敗: ${(err as Error).message}` });
    }
  }

  updateDocFolderSync(fallbackSiteUrl, folder.url, newPerFile);
  return { ingestedFiles, ingestedChunks, skippedFiles: skipped.length, deletedFiles, failedFiles };
}
