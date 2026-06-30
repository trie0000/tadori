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
  /** 取り込みが走った総スライド数 (= Vision 実行した新規 chunk 数)。 */
  ingestedSlides: number;
  /** 内容ハッシュ一致で Vision/embed をスキップしたスライド数 (差分取込の節約分)。 */
  skippedSlides: number;
  /** lastModified 一致でスキップしたファイル数。 */
  skippedFiles: number;
  /** 削除されたファイル (SP から消えていた)。 */
  deletedFiles: number;
  /** ファイル内で削除されたスライド数 (新版で消えたページ → 検索からも削除)。 */
  deletedSlides: number;
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

/** 増分判定: 前回 perFile と今回 SP の lastModified を比較。
 *  force=true なら全ファイルを toIngest にまとめる (Vision モデル変更時の再解析用)。
 *  targetFiles を指定すると、その名前に一致するファイルだけを toIngest に入れる
 *  (個別ファイル再取り込み用。force と併用すれば lastModified 無視で再処理)。 */
function pickTargets(
  now: FileInfo[],
  prev: Record<string, string>,
  force = false,
  targetFiles?: ReadonlySet<string>,
): {
  toIngest: FileInfo[];
  skipped: FileInfo[];
  deleted: string[];   // ファイル名 (今回 SP に無いもの)
} {
  const toIngest: FileInfo[] = [];
  const skipped: FileInfo[] = [];
  const nowNames = new Set<string>();
  for (const f of now) {
    nowNames.add(f.name);
    // 指定ファイルのみモード: target に無いファイルは触らない (skipped 扱いに)
    if (targetFiles && !targetFiles.has(f.name)) { skipped.push(f); continue; }
    if (force) { toIngest.push(f); continue; }
    const prevTs = prev[f.name];
    if (!prevTs || prevTs !== f.timeLastModified) toIngest.push(f);
    else skipped.push(f);
  }
  // 削除検知は「フォルダ全体同期」のときだけ行う。個別ファイル再取込時は他ファイルを巻き込まない。
  const deleted: string[] = [];
  if (!targetFiles) {
    for (const name of Object.keys(prev)) if (!nowNames.has(name)) deleted.push(name);
  }
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

/** スライドのソース内容ハッシュ (title + rawText + tables + notes)。
 *  PNG は対象外 (Shape テキストが同じなら図も同じとみなす — Export 結果はレイアウト
 *  非変化なら同一になるが、毎回バイト一致は保証されないので軽量側のテキストで判定)。
 *  djb2 + base36。Vision モデル ID は含めない (force 再取込で別管理するため)。 */
function slideSrcHash(slide: PptxSlide): string {
  const tablesStr = (slide.tables || []).map(t => t.map(row => row.join('')).join('')).join('');
  const s = `${slide.title || ''} ${slide.rawText || ''} ${tablesStr} ${slide.notes || ''}`;
  let h = 5381;
  for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}

/** Vision OFF 時のスライド markdown: title / 本文テキスト / 表 / ノートだけで組む。
 *  relay が抽出済みの rawText・tables・notes を使うので LLM 不要。 */
function slideTextMarkdown(slide: PptxSlide): string {
  const parts: string[] = [];
  if (slide.title?.trim()) parts.push(`# ${slide.title.trim()}`);
  if (slide.rawText?.trim()) parts.push(slide.rawText.trim());
  for (const tbl of slide.tables || []) {
    if (!tbl.length) continue;
    parts.push(tbl.map(row => '| ' + row.map(c => (c ?? '').replace(/\n/g, ' ')).join(' | ') + ' |').join('\n'));
  }
  if (slide.notes?.trim()) parts.push(`(ノート) ${slide.notes.trim()}`);
  return parts.join('\n\n').trim() || `(スライド ${slide.slideNo})`;
}

/** 1 スライド → IngestMail (markdown 化 + メタ付与)。 */
function slideToIngestMail(
  fileInfo: FileInfo,
  slide: PptxSlide,
  markdown: string,
  thumbServerRelUrl: string,
  srcHash: string,
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
    srcHash,
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

/** 1 フォルダ分の取り込み (新規 / 更新ファイル + 削除検知)。
 *  force=true なら lastModified によるスキップを無効化し、全 pptx を再解析する
 *  (Vision モデル変更で再分析したい場合等)。
 *  targetFiles を指定すると、その名前 (Set) に一致するファイルだけ再処理する
 *  (個別ファイル再取込)。force=true と併用が前提。削除検知はスキップ。
 *  Vision LLM のコストが再発生する点に注意。 */
export async function syncPptxFolder(
  folder: PptxFolderConfig,
  s: RuntimeSettings,
  fallbackSiteUrl: string,
  onProgress?: (p: PptxIngestProgress) => void,
  signal?: AbortSignal,
  opts: { force?: boolean; targetFiles?: ReadonlySet<string>; thumbsOnly?: boolean; vision?: boolean; persist?: (perFile: Record<string, string>) => void } = {},
): Promise<PptxIngestResult> {
  // vision 既定は true (従来通り)。false ならスライドの title/rawText/表/ノートだけで
  // markdown を組み、Vision LLM を呼ばない (ラベル/フォルダ単位の Vision OFF 用)。
  const useVision = opts.vision !== false;
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
  //   thumbsOnly: サムネ再生成のみ。全ファイル対象 + 削除検知なし + Vision/embed しない。
  //   force=true: lastModified 無視で全 pptx を再解析。targetFiles: 個別指定。
  const { toIngest, skipped, deleted } = opts.thumbsOnly
    ? pickTargets(pptxFiles, folder.perFile, true, opts.targetFiles)  // 全件 or 指定件を回す (削除検知は targetFiles 同様スキップさせるため下で deleted を無視)
    : pickTargets(pptxFiles, folder.perFile, opts.force === true, opts.targetFiles);
  if (opts.thumbsOnly) console.log('[tadori] pptx sync: サムネ再生成モード (Vision なし)');
  else if (opts.targetFiles) console.log(`[tadori] pptx sync: 個別ファイル再取込モード — 対象 ${opts.targetFiles.size} ファイル`);
  else if (opts.force) console.log('[tadori] pptx sync: 強制再取り込みモード — 全 pptx を再解析');

  // サムネは Tadori 管理フォルダ配下 (<site>/Shared Documents/Tadori/pptx-thumbs) に集約。
  // ★ オリジナルの PPTX フォルダには一切ファイルを作らない (ユーザの資料置き場を汚さない)。
  //   セグメント (manifest/seg-*.json) と同じ Tadori フォルダ配下にまとめる。
  const thumbFolderServerRel = (await getEngine(siteUrl)).store.pptxThumbFolder;

  // 3. 削除されたファイルの chunk を消す (サムネ再生成モードでは検索データを触らない)
  let deletedFiles = 0;
  if (!opts.thumbsOnly && deleted.length > 0) {
    const eng = await getEngine(siteUrl);
    for (const fname of deleted) {
      if (signal?.aborted) break;
      onProgress?.({ file: fname, fileIdx: 0, fileTotal: 0, slideIdx: 0, slideTotal: 0, phase: 'delete', message: `${fname} は SP から消えていた — chunk を削除` });
      // serverRelativeUrl がフォルダ内なので組み立て可
      const stale = `${folderServerRel}/${fname}`;
      const messageIds = eng.db.messageIdsForConversation(stale);
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
  let skippedSlides = 0;
  let deletedSlides = 0;
  let failedSlides = 0;
  // perFile の更新方針:
  //   - 通常同期 (targetFiles 無し): SP の現在の lastModified をそのまま採用
  //     (toIngest = 取込済み更新、skipped = 取込不要)
  //   - 個別再取込 (targetFiles 有り): target 外のファイルは **旧 perFile の値を保持**
  //     (実際に再処理していないのに「処理済み」と誤認しないため)
  const newPerFile: Record<string, string> = {};
  if (opts.targetFiles || opts.thumbsOnly) {
    // 個別再取込 / サムネ再生成: 既存 perFile を全部保持 (取込状態は変えない)。
    // ※ thumbsOnly は下のループでも newPerFile を書き換えない (サムネは内容変更ではない)。
    for (const [name, ts] of Object.entries(folder.perFile)) newPerFile[name] = ts;
  } else {
    // 通常同期: skipped (= 今回触らないが SP には居る) は SP の現在値で更新
    for (const f of skipped) newPerFile[f.name] = f.timeLastModified;
  }

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

      // ── サムネ再生成モード: Vision も embed も検索 DB も一切触らず、PNG だけ
      //    Tadori/pptx-thumbs に再アップロードする。誤って消したサムネの復旧用。
      if (opts.thumbsOnly) {
        let regenerated = 0;
        for (let j = 0; j < slides.length; j++) {
          if (signal?.aborted) break;
          const slide = slides[j];
          onProgress?.({
            file: f.name, fileIdx, fileTotal,
            slideIdx: j + 1, slideTotal: slides.length,
            phase: 'embed',
            message: `${f.name} サムネ再生成 ${j + 1}/${slides.length}…`,
          });
          try {
            const pngBytes = Uint8Array.from(atob(slide.pngBase64), c => c.charCodeAt(0)).buffer;
            await uploadThumb(sp, thumbFolderServerRel, f, slide.slideNo, pngBytes);
            regenerated++;
          } catch (e) {
            if ((e as Error).name === 'AbortError') throw e;
            console.warn('[pptx] thumb regen failed:', f.name, slide.slideNo, (e as Error).message);
            failedSlides++;
          }
        }
        if (regenerated > 0) ingestedFiles++;
        ingestedSlides += regenerated;
        onProgress?.({
          file: f.name, fileIdx, fileTotal,
          slideIdx: slides.length, slideTotal: slides.length,
          phase: 'done',
          message: `${f.name} サムネ再生成 完了 (${regenerated}/${slides.length})`,
        });
        continue; // Vision/embed/削除検知は完全スキップ
      }

      // 差分判定用に既存レコードを引く (messageId → srcHash)。
      // force=true (強制 / 個別再取込) のときは差分スキップせず全スライド Vision。
      const eng = await getEngine(siteUrl);
      const willHave = new Set(slides.map(sl => `pptx://${f.serverRelativeUrl}#${sl.slideNo}`));

      // 削除されたスライド (新版に無い既存 messageId) を検索 DB から除去
      try {
        const existing = eng.db.messageIdsForConversation(f.serverRelativeUrl);
        for (const mid of existing) {
          if (!willHave.has(mid)) {
            try { await deleteFromSegments(mid, siteUrl); deletedSlides++; }
            catch (e) { console.warn('[pptx] prune stale slide:', mid, (e as Error).message); }
          }
        }
      } catch { /* engine 未初期化等は無視 (初回) */ }

      // 各スライドは互いに独立なので、差分判定で残ったスライドの Vision/テキスト解析 +
      // サムネ生成を並列実行する (visionConcurrency 本まで)。mails は push 順不同で問題ない。
      const mails: IngestMail[] = [];
      const toProcess: { slide: PptxSlide; idx: number; hash: string }[] = [];
      for (let j = 0; j < slides.length; j++) {
        const slide = slides[j];
        const mid = `pptx://${f.serverRelativeUrl}#${slide.slideNo}`;
        const hash = slideSrcHash(slide);
        // 差分判定: force でなく、既存レコードの srcHash と一致 → スキップ
        if (!opts.force) {
          const existingRec = eng.db.get(mid);
          if (existingRec && existingRec.srcHash && existingRec.srcHash === hash) {
            skippedSlides++;
            onProgress?.({ file: f.name, fileIdx, fileTotal, slideIdx: j + 1, slideTotal: slides.length, phase: 'skip', message: `${f.name} スライド ${j + 1}/${slides.length} は変更なし — スキップ` });
            continue;
          }
        }
        toProcess.push({ slide, idx: j, hash });
      }

      const concurrency = Math.min(16, Math.max(1, s.visionConcurrency || 3));
      const processOne = async (item: { slide: PptxSlide; idx: number; hash: string }): Promise<void> => {
        if (signal?.aborted) return;
        const { slide, idx, hash } = item;
        onProgress?.({
          file: f.name, fileIdx, fileTotal, slideIdx: idx + 1, slideTotal: slides.length, phase: 'vision',
          message: useVision
            ? `${f.name} スライド ${idx + 1}/${slides.length} を Vision 解析中… (並列${concurrency})`
            : `${f.name} スライド ${idx + 1}/${slides.length} をテキスト抽出中… (並列${concurrency})`,
        });
        try {
          // Vision ON: 画像を LLM で markdown 化 / OFF: title+rawText+表+ノートだけで組む。
          const md = useVision
            ? (await describeSlide(slide as VisionSlideInput, s, signal)).markdown
            : slideTextMarkdown(slide);
          const pngBytes = Uint8Array.from(atob(slide.pngBase64), c => c.charCodeAt(0)).buffer;
          const thumbUrl = await uploadThumb(sp, thumbFolderServerRel, f, slide.slideNo, pngBytes);
          mails.push(slideToIngestMail(f, slide, md, thumbUrl, hash));
        } catch (e) {
          if ((e as Error).name === 'AbortError') return; // 中断は静かに (上位の signal で停止)
          console.warn('[pptx] vision/thumb failed:', f.name, slide.slideNo, (e as Error).message);
          failedSlides++;
        }
      };

      // 並列プール: 同時に concurrency 本まで。cursor を共有して各ワーカが次を取る。
      let cursor = 0;
      await Promise.all(Array.from({ length: Math.min(concurrency, toProcess.length) }, async () => {
        while (cursor < toProcess.length && !signal?.aborted) {
          await processOne(toProcess[cursor++]);
        }
      }));

      if (mails.length > 0) {
        onProgress?.({
          file: f.name, fileIdx, fileTotal,
          slideIdx: mails.length, slideTotal: slides.length,
          phase: 'embed',
          message: `${f.name} を埋め込み中… (${mails.length} chunk)`,
        });
        await ingestToSegments(mails, s, siteUrl, undefined, signal);
        ingestedSlides += mails.length;
      }
      // 何か処理した (新規/更新 or 削除) ファイルをカウント
      if (mails.length > 0 || deletedSlides > 0) ingestedFiles++;
      newPerFile[f.name] = f.timeLastModified;
      onProgress?.({
        file: f.name, fileIdx, fileTotal,
        slideIdx: slides.length, slideTotal: slides.length,
        phase: 'done',
        message: `${f.name} 完了 (更新 ${mails.length} / スキップ ${slides.length - mails.length - failedSlides} / 全 ${slides.length} スライド)`,
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

  // 5. perFile を保存。opts.persist があれば pptxFolders ではなくそちらへ
  //    (統合フォルダ取り込み = docFolder 側に pptx perFile を持たせる場合に使う)。
  if (opts.persist) opts.persist(newPerFile);
  else updatePptxFolderSync(fallbackSiteUrl, folder.url, newPerFile);

  return {
    ingestedFiles,
    ingestedSlides,
    skippedSlides,
    skippedFiles: skipped.length,
    deletedFiles,
    deletedSlides,
    failedSlides,
  };
}
