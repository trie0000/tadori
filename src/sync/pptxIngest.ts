// PPTX マニュアル取り込みパイプライン。
//
// フロー (1 フォルダ単位):
//   1. SP REST でフォルダ配下のファイル一覧取得 (.pptx でフィルタ)
//   2. 増分判定: lastModified が前回同期以降に変わったファイル + 未同期ファイル
//   3. 削除検知: 前回 perFile にあったが今回 SP に無いファイル → 関連 chunk を削除
//   4. 各対象ファイル:
//        a. SP REST で binary を ArrayBuffer 取得
//        b. relay /tadori/pptx-extract に POST → slides[] 受領
//        c. 各 slide を Vision LLM (GPT-5) で markdown 化
//        d. サムネ PNG (320x180 縮小版) を SP の thumbs フォルダにアップロード
//        e. IngestMail[] を組み立て (messageId = "pptx://<serverRelUrl>#<slideNo>")
//        f. ingestToSegments で SP セグメント書込 (既存 messageId は last-writer-wins で更新)
//   5. perFile を更新して localStorage に保存
//
// 設計参照: docs/pptx-rag-design.md §5, §10

import { SharePointClient, toServerRelativeUrl, type FileInfo } from '../sharepoint/client';
import { describeSlide, type VisionSlideInput } from '../embeddings/visionClient';
import { ingestToSegments, deleteFromSegments, type IngestMail } from '../db/writer';
import { getEngine } from '../db/engine';
import type { RuntimeSettings } from '../api/aiSettings';
import { updatePptxFolderSync, type PptxFolderConfig } from './pptxFolders';

/** relay の /tadori/pptx-extract レスポンス内の 1 スライド分。 */
interface PptxSlide {
  slideNo: number;
  title: string;
  pngBase64: string;
  rawText: string;
  tables: string[][][];
  notes: string;
}

interface PptxExtractResult {
  ok: boolean;
  count: number;
  slides: PptxSlide[];
}

export interface PptxIngestProgress {
  /** 処理中のファイル名 ("manual-A.pptx" など)。 */
  file: string;
  /** ファイル全体の中の処理位置 (1-origin)。 */
  fileIdx: number;
  fileTotal: number;
  /** 現在ファイルのスライド処理位置 (1-origin)。0 = 未着手。 */
  slideIdx: number;
  slideTotal: number;
  /** 状態。 */
  phase: 'fetch' | 'extract' | 'vision' | 'embed' | 'done' | 'skip' | 'error' | 'delete';
  /** メッセージ (UI トースト用)。 */
  message?: string;
}

export interface PptxIngestResult {
  /** 取り込みが走った (新規 or 更新) ファイル数。 */
  ingestedFiles: number;
  /** 取り込みが走った総スライド数 (= 新規 chunk 数)。 */
  ingestedSlides: number;
  /** lastModified 一致でスキップしたファイル数。 */
  skippedFiles: number;
  /** 削除されたファイル (SP から消えていた)。 */
  deletedFiles: number;
  /** Vision 等で失敗したスライド数。 */
  failedSlides: number;
}

/** serverRelativeUrl からサイトコレクションのパス部分を取り出す。
 *  /sites/<name>/... → /sites/<name>
 *  /teams/<name>/... → /teams/<name>
 *  /personal/<user>/... → /personal/<user>  (OneDrive 用)
 *  その他 (ルートサイト) → '' (空文字)
 *  SP REST は web スコープで呼ぶ必要があり、サイトコレクションを跨ぐと取得不可。
 *  そのため siteUrl は origin + これを組合せた値にする。 */
function siteCollectionPath(serverRel: string): string {
  const m = serverRel.match(/^(\/(?:sites|teams|personal)\/[^/]+)/i);
  return m ? m[1] : '';
}

/** 受け取った url (絶対 URL or serverRelativeUrl) を解決して、SP クライアントが
 *  期待する serverRelativeUrl と siteUrl (サイトコレクションスコープ) を組み立てる。 */
function resolveSpFolder(folderUrl: string, fallbackSiteUrl: string): { siteUrl: string; folderServerRel: string } {
  const trimmed = folderUrl.trim();
  if (!trimmed) throw new Error('フォルダ URL が空です');
  const folderServerRel = toServerRelativeUrl(trimmed);
  // origin (https://host) を決める: 絶対 URL なら入力から、そうでなければ fallback から。
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
  // サイトコレクションを serverRelative から抽出し、origin と合体させる。
  // これで /sites/foo/... へのアクセスは https://host/sites/foo/_api/web に行く。
  const scPath = siteCollectionPath(folderServerRel);
  const siteUrl = origin + scPath;
  if (!folderServerRel) throw new Error(`フォルダ URL の serverRelativeUrl を解釈できませんでした: ${folderUrl}`);
  return { siteUrl, folderServerRel };
}

/** PPTX ファイルだけに絞る (大文字小文字不問、隠しファイルとロックファイルを除外)。 */
function filterPptxFiles(items: FileInfo[]): FileInfo[] {
  return items.filter(f => {
    const n = f.name.toLowerCase();
    if (!n.endsWith('.pptx')) return false;
    if (n.startsWith('~$')) return false; // PowerPoint のロックファイル
    if (n.startsWith('.')) return false;
    return true;
  });
}

/** 増分判定: 前回 perFile と今回 SP の lastModified を比較。 */
function pickTargets(now: FileInfo[], prev: Record<string, string>): {
  toIngest: FileInfo[];
  skipped: FileInfo[];
  deleted: string[];   // ファイル名 (今回 SP に無いもの)
} {
  const toIngest: FileInfo[] = [];
  const skipped: FileInfo[] = [];
  const nowNames = new Set<string>();
  for (const f of now) {
    nowNames.add(f.name);
    const prevTs = prev[f.name];
    if (!prevTs || prevTs !== f.timeLastModified) toIngest.push(f);
    else skipped.push(f);
  }
  const deleted: string[] = [];
  for (const name of Object.keys(prev)) if (!nowNames.has(name)) deleted.push(name);
  return { toIngest, skipped, deleted };
}

/** relay の /tadori/pptx-extract を叩いてスライド配列を取得。 */
async function callPptxExtract(
  relayBaseUrl: string,
  pptxBytes: ArrayBuffer,
  fileName: string,
  signal?: AbortSignal,
): Promise<PptxSlide[]> {
  if (!relayBaseUrl) throw new Error('中継サーバ URL が未設定です (AI 接続で設定)');
  const url = `${relayBaseUrl.replace(/\/+$/, '')}/tadori/pptx-extract`;
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/octet-stream',
      'X-Tadori-Filename': encodeURIComponent(fileName),
    },
    body: pptxBytes,
    signal,
  });
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw new Error(`relay /pptx-extract HTTP ${res.status} ${t.slice(0, 300)}`);
  }
  const json = await res.json() as PptxExtractResult;
  if (!json.ok || !Array.isArray(json.slides)) throw new Error('relay /pptx-extract: 不正なレスポンス');
  return json.slides;
}

/** PowerPoint でファイル + 該当スライドを開く (引用ジャンプ)。 */
export async function openPptxAtSlide(
  relayBaseUrl: string,
  fileUrl: string,
  slideNo: number,
  signal?: AbortSignal,
): Promise<void> {
  if (!relayBaseUrl) throw new Error('中継サーバ URL が未設定です (AI 接続で設定)');
  const url = `${relayBaseUrl.replace(/\/+$/, '')}/tadori/pptx-open`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fileUrl, slideNo }),
    signal,
  });
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw new Error(`relay /pptx-open HTTP ${res.status} ${t.slice(0, 300)}`);
  }
}

/** 1 スライド → IngestMail (markdown 化 + メタ付与)。 */
function slideToIngestMail(
  fileInfo: FileInfo,
  slide: PptxSlide,
  markdown: string,
  thumbServerRelUrl: string,
): IngestMail {
  const subject = slide.title?.trim() || `スライド ${slide.slideNo}`;
  const docPath = `pptx://${fileInfo.serverRelativeUrl}#${slide.slideNo}`;
  const notesAppended = slide.notes ? `${markdown}\n\n## ノート\n${slide.notes}` : markdown;
  return {
    messageId: docPath,                       // ファイル + スライドで一意
    internetMessageId: '',
    conversationId: fileInfo.serverRelativeUrl, // 同一 pptx 内をスレッド扱い
    kind: 'pptx',
    chunkIdx: slide.slideNo - 1,
    chunkCount: 0,                            // 後で再設定はしない (slide 数は可変)
    docPath,
    subject,
    from: fileInfo.name,
    to: [],
    cc: [],
    date: fileInfo.timeLastModified || new Date().toISOString(),
    body: notesAppended,
    isHtml: false,
    pptxFile: fileInfo.name,
    pptxServerRelUrl: fileInfo.serverRelativeUrl,
    slideNo: slide.slideNo,
    slideTitle: subject,
    thumbServerRelUrl,
  };
}

/** サムネ PNG を SP にアップロード。失敗しても致命的ではないので警告ログのみ。
 *  ファイル名は ASCII 化 (SP の URL 安全性) して uniqueness は serverRelUrl+slideNo から hash で確保。 */
async function uploadThumb(
  sp: SharePointClient,
  thumbFolderServerRel: string,
  fileInfo: FileInfo,
  slideNo: number,
  pngBytes: ArrayBuffer,
): Promise<string> {
  const safeName = `thumb-${fileInfo.uniqueId || 'unknown'}-${slideNo}.png`;
  try {
    await sp.ensureFolder(thumbFolderServerRel);
    await sp.uploadFileBytes(thumbFolderServerRel, safeName, pngBytes);
    return `${thumbFolderServerRel}/${safeName}`;
  } catch (e) {
    console.warn('[pptx] thumb upload failed:', (e as Error).message);
    return '';
  }
}

/** ファイル名から既存の DB レコード (kind='pptx', conversationId=serverRelUrl) の messageId 群を抽出。
 *  ファイル削除時の chunk 削除や、スライド数縮小時の余剰削除に使う。 */
function findExistingMessageIds(eng: Awaited<ReturnType<typeof getEngine>>, serverRelUrl: string): string[] {
  const out: string[] = [];
  // VectorDb の内部走査用 API は提供されてないので、records は engine.db を直接覗く設計。
  // ここでは安全に、kind='pptx' && conversationId === serverRelUrl のレコードを探す。
  const db = eng.db as unknown as { records: Map<string, { messageId: string; kind?: string; conversationId?: string }> };
  if (!db.records || typeof db.records.forEach !== 'function') return out;
  db.records.forEach((r) => {
    if (r.kind === 'pptx' && r.conversationId === serverRelUrl) out.push(r.messageId);
  });
  return out;
}

/** 1 フォルダ分の取り込み (新規 / 更新ファイル + 削除検知)。 */
export async function syncPptxFolder(
  folder: PptxFolderConfig,
  s: RuntimeSettings,
  fallbackSiteUrl: string,
  onProgress?: (p: PptxIngestProgress) => void,
  signal?: AbortSignal,
): Promise<PptxIngestResult> {
  const { siteUrl, folderServerRel } = resolveSpFolder(folder.url, fallbackSiteUrl);
  const sp = new SharePointClient(siteUrl);
  console.log('[tadori] pptx sync start',
    { inputUrl: folder.url, siteUrl, folderServerRel, recursive: folder.recursive });

  // 1. enumerate
  onProgress?.({ file: '', fileIdx: 0, fileTotal: 0, slideIdx: 0, slideTotal: 0, phase: 'fetch', message: `フォルダ一覧を取得中… (${folderServerRel})` });
  const items = await sp.listFolderItems(folderServerRel, { recursive: folder.recursive });
  const pptxFiles = filterPptxFiles(items);
  console.log(`[tadori] pptx sync: ${items.length} items found / ${pptxFiles.length} pptx files`);
  if (pptxFiles.length === 0 && items.length === 0) {
    onProgress?.({
      file: '', fileIdx: 0, fileTotal: 0, slideIdx: 0, slideTotal: 0,
      phase: 'skip',
      message: `フォルダにファイルが見つかりません (${folderServerRel})。URL/権限を確認してください。`,
    });
  } else if (pptxFiles.length === 0) {
    onProgress?.({
      file: '', fileIdx: 0, fileTotal: 0, slideIdx: 0, slideTotal: 0,
      phase: 'skip',
      message: `${items.length} 個のファイルがありますが .pptx は 0 件でした。`,
    });
  }

  // 2. 増分判定 & 削除検知
  const { toIngest, skipped, deleted } = pickTargets(pptxFiles, folder.perFile);

  // サムネは <pptx-folder>/.tadori-thumbs に集約 (1 フォルダ運用想定)
  const thumbFolderServerRel = `${folderServerRel}/.tadori-thumbs`;

  // 3. 削除されたファイルの chunk を消す
  let deletedFiles = 0;
  if (deleted.length > 0) {
    const eng = await getEngine(siteUrl);
    for (const fname of deleted) {
      if (signal?.aborted) break;
      onProgress?.({ file: fname, fileIdx: 0, fileTotal: 0, slideIdx: 0, slideTotal: 0, phase: 'delete', message: `${fname} は SP から消えていた — chunk を削除` });
      // serverRelativeUrl がフォルダ内なので組み立て可
      const stale = `${folderServerRel}/${fname}`;
      const messageIds = findExistingMessageIds(eng, stale);
      for (const mid of messageIds) {
        try { await deleteFromSegments(mid, siteUrl); } catch (e) {
          console.warn('[pptx] delete chunk failed:', mid, (e as Error).message);
        }
      }
      deletedFiles++;
    }
  }

  // 4. 各ファイル取り込み
  let ingestedFiles = 0;
  let ingestedSlides = 0;
  let failedSlides = 0;
  const newPerFile: Record<string, string> = {};
  // skipped はそのまま perFile に残す (今回触らない)
  for (const f of skipped) newPerFile[f.name] = f.timeLastModified;

  for (let i = 0; i < toIngest.length; i++) {
    if (signal?.aborted) break;
    const f = toIngest[i];
    const fileIdx = i + 1;
    const fileTotal = toIngest.length;

    try {
      // fetch binary
      onProgress?.({ file: f.name, fileIdx, fileTotal, slideIdx: 0, slideTotal: 0, phase: 'fetch', message: `${f.name} をダウンロード中…` });
      const bytes = await sp.fetchFileBytes(f.serverRelativeUrl);
      if (!bytes) {
        onProgress?.({ file: f.name, fileIdx, fileTotal, slideIdx: 0, slideTotal: 0, phase: 'skip', message: `${f.name} が SP から取得できませんでした (削除済み?)` });
        continue;
      }

      // relay extract
      onProgress?.({ file: f.name, fileIdx, fileTotal, slideIdx: 0, slideTotal: 0, phase: 'extract', message: `${f.name} を PowerPoint で解析中…` });
      const slides = await callPptxExtract(s.relayBaseUrl, bytes, f.name, signal);

      // 既存 chunk のうち、新版で消えるスライド (例: 元 10 枚 → 8 枚) を先に削除
      try {
        const eng = await getEngine(siteUrl);
        const existing = findExistingMessageIds(eng, f.serverRelativeUrl);
        const willHave = new Set(slides.map(sl => `pptx://${f.serverRelativeUrl}#${sl.slideNo}`));
        for (const mid of existing) {
          if (!willHave.has(mid)) {
            try { await deleteFromSegments(mid, siteUrl); } catch (e) { console.warn('[pptx] prune stale slide:', mid, (e as Error).message); }
          }
        }
      } catch { /* engine 未初期化等は無視 (初回) */ }

      // 各スライドを Vision LLM で markdown 化 + サムネアップロード
      const mails: IngestMail[] = [];
      for (let j = 0; j < slides.length; j++) {
        if (signal?.aborted) break;
        const slide = slides[j];
        onProgress?.({
          file: f.name, fileIdx, fileTotal,
          slideIdx: j + 1, slideTotal: slides.length,
          phase: 'vision',
          message: `${f.name} スライド ${j + 1}/${slides.length} を Vision LLM で解析中…`,
        });

        try {
          const v = await describeSlide(slide as VisionSlideInput, s, signal);
          // サムネ (Vision 用と同じ PNG をそのまま保存。縮小は将来課題)
          const pngBytes = Uint8Array.from(atob(slide.pngBase64), c => c.charCodeAt(0)).buffer;
          const thumbUrl = await uploadThumb(sp, thumbFolderServerRel, f, slide.slideNo, pngBytes);
          mails.push(slideToIngestMail(f, slide, v.markdown, thumbUrl));
        } catch (e) {
          if ((e as Error).name === 'AbortError') throw e;
          console.warn('[pptx] vision/thumb failed:', f.name, slide.slideNo, (e as Error).message);
          failedSlides++;
        }
      }

      if (mails.length > 0) {
        onProgress?.({
          file: f.name, fileIdx, fileTotal,
          slideIdx: mails.length, slideTotal: slides.length,
          phase: 'embed',
          message: `${f.name} を埋め込み中… (${mails.length} chunk)`,
        });
        await ingestToSegments(mails, s, siteUrl, undefined, signal);
        ingestedSlides += mails.length;
        ingestedFiles++;
      }
      newPerFile[f.name] = f.timeLastModified;
      onProgress?.({
        file: f.name, fileIdx, fileTotal,
        slideIdx: slides.length, slideTotal: slides.length,
        phase: 'done',
        message: `${f.name} 完了 (${mails.length}/${slides.length} スライド)`,
      });
    } catch (e) {
      if ((e as Error).name === 'AbortError') throw e;
      console.error('[pptx] file failed:', f.name, e);
      onProgress?.({
        file: f.name, fileIdx, fileTotal,
        slideIdx: 0, slideTotal: 0,
        phase: 'error',
        message: `${f.name} 失敗: ${(e as Error).message}`,
      });
    }
  }

  // 5. perFile を保存
  updatePptxFolderSync(folder.url, newPerFile);

  return {
    ingestedFiles,
    ingestedSlides,
    skippedFiles: skipped.length,
    deletedFiles,
    failedSlides,
  };
}
