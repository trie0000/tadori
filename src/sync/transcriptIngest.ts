// Teams 会議文字起こし (.vtt) 取り込みパイプライン。
//
// フロー (1 フォルダ):
//   1. SP REST でフォルダ配下のファイル一覧 (.vtt でフィルタ)
//   2. 増分判定 (lastModified) + 削除検知
//   3. 各 .vtt:
//        a. fetchFileBytes → UTF-8 デコード
//        b. parseVtt → cue[]
//        c. cuesToText → splitIntoChunks (生チャンク)
//        d. IngestMail[] (kind='transcript') → ingestToSegments
//   4. perFile 更新
//
// PPTX と違い relay も Vision もサムネも無く、ブラウザ内で完結。
// 設計参照: docs/teams-transcript-design.md §7

import { SharePointClient, toServerRelativeUrl, type FileInfo } from '../sharepoint/client';
import { parseVtt, cuesToText, type VttCue } from '../transcript/vtt';
import { splitIntoChunks } from '../lib/chunk';
import { ingestToSegments, deleteFromSegments, type IngestMail } from '../db/writer';
import { getEngine } from '../db/engine';
import type { RuntimeSettings } from '../api/aiSettings';
import { updateTranscriptFolderSync, type TranscriptFolderConfig } from './transcriptFolders';

export interface TranscriptIngestProgress {
  file: string;
  fileIdx: number;
  fileTotal: number;
  /** 現在ファイルのチャンク処理位置 (1-origin)。0 = 未着手。 */
  chunkIdx: number;
  chunkTotal: number;
  phase: 'fetch' | 'parse' | 'embed' | 'done' | 'skip' | 'error' | 'delete';
  message?: string;
}

export interface TranscriptIngestResult {
  ingestedFiles: number;
  ingestedChunks: number;
  skippedFiles: number;
  deletedFiles: number;
  failedFiles: number;
}

/** serverRelativeUrl からサイトコレクションのパス部分を取り出す (pptxIngest と同じ)。 */
function siteCollectionPath(serverRel: string): string {
  const m = serverRel.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
  return m ? m[1] : '';
}

function resolveSpFolder(folderUrl: string, fallbackSiteUrl: string): { siteUrl: string; folderServerRel: string } {
  const trimmed = folderUrl.trim();
  if (!trimmed) throw new Error('フォルダ URL が空です');
  const folderServerRel = toServerRelativeUrl(trimmed);
  let origin = '';
  try {
    const u = new URL(trimmed);
    origin = `${u.protocol}//${u.host}`;
  } catch {
    try {
      const fb = new URL(fallbackSiteUrl);
      origin = `${fb.protocol}//${fb.host}`;
    } catch {
      origin = fallbackSiteUrl.replace(/\/+$/, '').replace(/\/_api\/.*$/, '');
    }
  }
  const siteUrl = origin + siteCollectionPath(folderServerRel);
  if (!folderServerRel) throw new Error(`フォルダ URL の serverRelativeUrl を解釈できませんでした: ${folderUrl}`);
  return { siteUrl, folderServerRel };
}

/** .vtt だけに絞る (大文字小文字不問、隠し/ロックファイル除外)。 */
function filterVttFiles(items: FileInfo[]): FileInfo[] {
  return items.filter(f => {
    const n = f.name.toLowerCase();
    if (!n.endsWith('.vtt')) return false;
    if (n.startsWith('~$') || n.startsWith('.')) return false;
    return true;
  });
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
  if (!targetFiles) {
    for (const name of Object.keys(prev)) if (!nowNames.has(name)) deleted.push(name);
  }
  return { toIngest, skipped, deleted };
}

/** djb2/base36 ハッシュ (差分判定用)。 */
function hashStr(s: string): string {
  let h = 5381;
  for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}

/** ファイル名から会議名 + 日時を推定。
 *  例: "週次定例-20260528_100000-Meeting Recording.vtt"
 *      → { meetingName: "週次定例", dateIso: "2026-05-28T10:00:00Z" }
 *  日時トークンが取れなければ dateIso=null (呼び出し側で lastModified にフォールバック)。 */
function parseMeetingNameDate(fileName: string): { meetingName: string; dateIso: string | null } {
  let base = fileName.replace(/\.vtt$/i, '');
  // YYYYMMDD_HHMMSS / YYYY-MM-DD HH-MM-SS 等の日時トークンを拾う
  let dateIso: string | null = null;
  const m1 = base.match(/(\d{4})(\d{2})(\d{2})[_-]?(\d{2})(\d{2})(\d{2})/);
  const m2 = base.match(/(\d{4})[-/](\d{2})[-/](\d{2})[ T_-]?(\d{2})[:-]?(\d{2})(?:[:-]?(\d{2}))?/);
  const m = m1 || m2;
  if (m) {
    const [, y, mo, d, h, mi, se] = m;
    const yyyy = Number(y), MM = Number(mo), dd = Number(d);
    const hh = Number(h), mm = Number(mi), ss = se ? Number(se) : 0;
    if (yyyy > 1900 && MM >= 1 && MM <= 12 && dd >= 1 && dd <= 31) {
      // ローカル時刻として ISO 化 (タイムゾーン情報は無いので素直に組み立て)
      const pad = (n: number): string => String(n).padStart(2, '0');
      dateIso = `${yyyy}-${pad(MM)}-${pad(dd)}T${pad(hh)}:${pad(mm)}:${pad(ss)}`;
      // 会議名は日時トークンを除いた残り
      base = base.replace(m[0], ' ');
    }
  }
  // 定型 suffix / 余分な区切りを掃除
  const meetingName = base
    .replace(/-?Meeting Recording/ig, '')
    .replace(/[-_]+/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .replace(/^[\s「『]+|[\s」』]+$/g, '')
    .trim() || fileName.replace(/\.vtt$/i, '');
  return { meetingName, dateIso };
}

/** チャンク内の主な話者 (先頭発言の話者) を返す。 */
function leadSpeaker(chunkText: string): string {
  const m = chunkText.match(/^\[[^\]]*?\s([^\]]+)\]/);
  return m ? m[1].trim() : '';
}

/** 同名 (拡張子違い) の録画 .mp4 があれば serverRelativeUrl を返す。 */
function findRecording(vtt: FileInfo, all: FileInfo[]): string {
  const stem = vtt.name.replace(/\.vtt$/i, '').toLowerCase();
  const mp4 = all.find(f => {
    const n = f.name.toLowerCase();
    return (n.endsWith('.mp4') || n.endsWith('.m4v')) && n.replace(/\.(mp4|m4v)$/i, '') === stem;
  });
  return mp4 ? mp4.serverRelativeUrl : '';
}

/** 1 フォルダ分の取り込み。 */
export async function syncTranscriptFolder(
  folder: TranscriptFolderConfig,
  s: RuntimeSettings,
  fallbackSiteUrl: string,
  onProgress?: (p: TranscriptIngestProgress) => void,
  signal?: AbortSignal,
  opts: { force?: boolean; targetFiles?: ReadonlySet<string> } = {},
): Promise<TranscriptIngestResult> {
  const { siteUrl, folderServerRel } = resolveSpFolder(folder.url, fallbackSiteUrl);
  const sp = new SharePointClient(siteUrl);
  console.log('[tadori] transcript sync start', { inputUrl: folder.url, siteUrl, folderServerRel });

  onProgress?.({ file: '', fileIdx: 0, fileTotal: 0, chunkIdx: 0, chunkTotal: 0, phase: 'fetch', message: `フォルダ一覧を取得中… (${folderServerRel})` });
  const items = await sp.listFolderItems(folderServerRel, { recursive: folder.recursive });
  const vttFiles = filterVttFiles(items);
  console.log(`[tadori] transcript sync: ${items.length} items / ${vttFiles.length} vtt files`);
  if (vttFiles.length === 0) {
    onProgress?.({
      file: '', fileIdx: 0, fileTotal: 0, chunkIdx: 0, chunkTotal: 0, phase: 'skip',
      message: items.length === 0
        ? `フォルダにファイルが見つかりません (${folderServerRel})。URL/権限を確認してください。`
        : `${items.length} 個のファイルがありますが .vtt は 0 件でした。`,
    });
  }

  const { toIngest, skipped, deleted } = pickTargets(vttFiles, folder.perFile, opts.force === true, opts.targetFiles);

  // 削除検知: SP から消えた .vtt の chunk を除去
  let deletedFiles = 0;
  if (!opts.targetFiles && deleted.length > 0) {
    const eng = await getEngine(fallbackSiteUrl);
    for (const fname of deleted) {
      if (signal?.aborted) break;
      onProgress?.({ file: fname, fileIdx: 0, fileTotal: 0, chunkIdx: 0, chunkTotal: 0, phase: 'delete', message: `${fname} は SP から消えていた — chunk を削除` });
      const stale = `${folderServerRel}/${fname}`;
      for (const mid of eng.db.messageIdsForConversation(stale)) {
        try { await deleteFromSegments(mid, fallbackSiteUrl); } catch (e) { console.warn('[transcript] delete chunk failed:', mid, (e as Error).message); }
      }
      deletedFiles++;
    }
  }

  let ingestedFiles = 0;
  let ingestedChunks = 0;
  let failedFiles = 0;
  const newPerFile: Record<string, string> = {};
  if (opts.targetFiles) {
    for (const [name, ts] of Object.entries(folder.perFile)) newPerFile[name] = ts;
  } else {
    for (const f of skipped) newPerFile[f.name] = f.timeLastModified;
  }

  for (let i = 0; i < toIngest.length; i++) {
    if (signal?.aborted) break;
    const f = toIngest[i];
    const fileIdx = i + 1;
    const fileTotal = toIngest.length;
    try {
      onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: 0, chunkTotal: 0, phase: 'fetch', message: `${f.name} をダウンロード中…` });
      const bytes = await sp.fetchFileBytes(f.serverRelativeUrl);
      if (!bytes) {
        onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: 0, chunkTotal: 0, phase: 'skip', message: `${f.name} が取得できませんでした` });
        continue;
      }

      onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: 0, chunkTotal: 0, phase: 'parse', message: `${f.name} を解析中…` });
      const vttText = new TextDecoder('utf-8').decode(bytes);
      const cues: VttCue[] = parseVtt(vttText);
      if (cues.length === 0) {
        onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: 0, chunkTotal: 0, phase: 'skip', message: `${f.name} に発言が見つかりませんでした` });
        newPerFile[f.name] = f.timeLastModified;
        continue;
      }

      const { meetingName, dateIso } = parseMeetingNameDate(f.name);
      const dateStr = dateIso || f.timeLastModified || new Date().toISOString();
      const recordingUrl = findRecording(f, items);

      // cue → テキスト → チャンク。各チャンクの先頭 cue 開始秒を startSec に。
      const fullText = cuesToText(cues);
      const chunks = splitIntoChunks(fullText, { maxChars: 800, overlap: 80 });
      if (chunks.length === 0) { newPerFile[f.name] = f.timeLastModified; continue; }

      // チャンク先頭の "[h:mm:ss 話者]" から開始秒を逆算するため、cue の時刻表をたどる。
      // 簡易に: チャンク内の最初のタイムスタンプ文字列を拾って秒へ。
      const mails: IngestMail[] = chunks.map((c, idx) => {
        const startSec = firstStartSec(c.text, cues);
        const docPath = `transcript://${f.serverRelativeUrl}#${idx}`;
        return {
          messageId: docPath,
          internetMessageId: '',
          conversationId: f.serverRelativeUrl,   // 同一会議の全チャンクを束ねる
          kind: 'transcript',
          chunkIdx: idx,
          chunkCount: chunks.length,
          docPath,
          subject: meetingName,
          from: leadSpeaker(c.text) || '(会議)',
          to: [],
          cc: [],
          date: dateStr,
          body: c.text,
          isHtml: false,
          transcriptFile: f.name,
          vttServerRelUrl: f.serverRelativeUrl,
          recordingServerRelUrl: recordingUrl,
          startSec,
          srcHash: hashStr(c.text),
        };
      });

      // 既存 chunk のうち新版に無い index を削除 (会議が短くなった場合)
      try {
        const eng = await getEngine(fallbackSiteUrl);
        const willHave = new Set(mails.map(m => m.messageId));
        for (const mid of eng.db.messageIdsForConversation(f.serverRelativeUrl)) {
          if (!willHave.has(mid)) {
            try { await deleteFromSegments(mid, fallbackSiteUrl); } catch { /* best-effort */ }
          }
        }
      } catch { /* 初回は engine 空 */ }

      onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: chunks.length, chunkTotal: chunks.length, phase: 'embed', message: `${f.name} を埋め込み中… (${chunks.length} chunk)` });
      await ingestToSegments(mails, s, fallbackSiteUrl, undefined, signal);
      ingestedChunks += mails.length;
      ingestedFiles++;
      newPerFile[f.name] = f.timeLastModified;
      onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: chunks.length, chunkTotal: chunks.length, phase: 'done', message: `${f.name} 完了 (${chunks.length} chunk)` });
    } catch (e) {
      if ((e as Error).name === 'AbortError') throw e;
      console.error('[transcript] file failed:', f.name, e);
      failedFiles++;
      onProgress?.({ file: f.name, fileIdx, fileTotal, chunkIdx: 0, chunkTotal: 0, phase: 'error', message: `${f.name} 失敗: ${(e as Error).message}` });
    }
  }

  updateTranscriptFolderSync(fallbackSiteUrl, folder.url, newPerFile);
  return { ingestedFiles, ingestedChunks, skippedFiles: skipped.length, deletedFiles, failedFiles };
}

/** チャンクテキスト先頭の "[h:mm:ss 話者]" の時刻を秒に変換して返す。
 *  cues から逆引きはせず、テキストの先頭タイムスタンプを直接パース (cuesToText と同じ形式)。 */
function firstStartSec(chunkText: string, cues: VttCue[]): number {
  const m = chunkText.match(/^\[(\d+(?::\d{2}){1,2})\s/);
  if (!m) return cues[0]?.startSec ?? 0;
  const parts = m[1].split(':').map(Number);
  if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + parts[2];
  if (parts.length === 2) return parts[0] * 60 + parts[1];
  return 0;
}

/** 録画 (Stream) を該当時刻から開く URL を組み立てて新規タブで開く。
 *  録画が無ければ .vtt をブラウザ表示。relay 不要・window.open のみ。 */
export function openTranscriptSource(hit: { recordingServerRelUrl?: string; vttServerRelUrl?: string; startSec?: number }): void {
  const origin = location.origin;
  if (hit.recordingServerRelUrl) {
    const abs = hit.recordingServerRelUrl.startsWith('http') ? hit.recordingServerRelUrl : `${origin}${hit.recordingServerRelUrl}`;
    const t = Math.max(0, Math.floor(hit.startSec ?? 0));
    // Stream/SharePoint の動画は ?t=<秒> または #t= で頭出しできる
    window.open(`${abs}${abs.includes('?') ? '&' : '?'}t=${t}`, '_blank', 'noopener');
    return;
  }
  if (hit.vttServerRelUrl) {
    const abs = hit.vttServerRelUrl.startsWith('http') ? hit.vttServerRelUrl : `${origin}${hit.vttServerRelUrl}`;
    window.open(abs, '_blank', 'noopener');
  }
}
