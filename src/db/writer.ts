// ベクトルDB への書き込み: メール → 埋め込み → セグメント生成 → SharePoint UL。
// 各書き込みバッチ = 新しい不変セグメント (≤SEGMENT_CAP 件)。既存 message-id は
// 重複排除でスキップ。削除は tombstone セグメントを追記。
//
// ※ 複数人運用では「書き込み担当(リース保有者)」だけがこれを呼ぶ。リース選出は
//   src/sync/lease.ts (別タスク) で gating する。

import { getEngine } from './engine';
import { getLease } from '../sync/lease';
import { embedDocsFor } from '../embeddings/router';
import { encodeEmbedding } from '../lib/float16';
import { normalize } from '../search/cosine';
import { cleanBody } from '../lib/mailtext';
import { htmlToText } from '../lib/mailhtml';
import {
  nextSegmentIndex, segmentId,
  type Segment, type SegmentRecord,
} from '../sync/segments';
import { updateManifestWithCas } from '../sync/spStore';
import type { RuntimeSettings } from '../api/aiSettings';

export interface IngestMail {
  messageId: string;
  /** RFC2822 Internet-Message-Id (Outlook での再検索キー)。 */
  internetMessageId?: string;
  /** スレッド識別子 (Outlook ConversationID) または OneNote ページID 等の親ドキュメントID。 */
  conversationId?: string;
  /** ソース種別。省略時は 'mail'。 */
  kind?: 'mail' | 'onenote' | 'doc' | 'pptx' | 'transcript';
  /** 取り込みバッチのラベル (OneNote のラベル付きバッチ用)。 */
  label?: string;
  /** PPTX 関連メタ (kind='pptx' のときのみ意味を持つ)。検索結果カード/ジャンプで使用。 */
  pptxFile?: string;
  pptxServerRelUrl?: string;
  slideNo?: number;
  slideTitle?: string;
  thumbServerRelUrl?: string;
  /** Teams 文字起こしメタ (kind='transcript' のときのみ意味を持つ)。 */
  transcriptFile?: string;
  vttServerRelUrl?: string;
  recordingServerRelUrl?: string;
  startSec?: number;
  /** ドキュメント (kind='doc') メタ。 */
  docFile?: string;
  docServerRelUrl?: string;
  /** ソース内容ハッシュ (差分判定用)。 */
  srcHash?: string;
  chunkIdx?: number;
  chunkCount?: number;
  docPath?: string;
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  date: string;
  body: string;
  /** body が HTML 形式か。埋め込みは本文テキストを抽出して行う。 */
  isHtml?: boolean;
}

export interface WriteResult {
  added: number;
  skipped: number;
  segments: number;
  /** 途中停止したか (停止までに保存した分は確定済み)。 */
  cancelled: boolean;
}

export type WritePhase = 'sync' | 'embed' | 'upload';

/** 1 バッチ = 1 セグメントで即コミットする件数。停止時の最小ロス単位。 */
const BATCH = 100;

/** 新規メールを埋め込んでセグメント化し SharePoint へ書き込む。
 *  signal で途中停止可能。バッチごとに確定するので、停止しても保存済みは残り、
 *  再実行すると message-id 重複排除で続きから取り込める。 */
export async function ingestToSegments(
  mails: IngestMail[],
  s: RuntimeSettings,
  siteUrl: string,
  onProgress?: (phase: WritePhase, done: number, total: number) => void,
  signal?: AbortSignal,
): Promise<WriteResult> {
  // 書き込み権 (リース) を取得。取れなければ他の人が取り込み中。
  if (!await getLease(siteUrl).ensureWriter()) {
    throw new Error('書き込み権限がありません (他のメンバーが取り込み中)。しばらく待って再実行してください。');
  }

  const eng = await getEngine(siteUrl);
  onProgress?.('sync', 0, 0);
  await eng.sync.sync(); // 書き込み前に最新へ追いつく

  // 重複排除: 既にDBにある message-id は除外
  const seen = new Set<string>();
  const fresh = mails.filter(m => {
    if (!m.messageId || eng.db.has(m.messageId) || seen.has(m.messageId)) return false;
    seen.add(m.messageId);
    return true;
  });
  const skipped = mails.length - fresh.length;
  if (fresh.length === 0) return { added: 0, skipped, segments: 0, cancelled: false };

  // CAS で安全に更新するため、manifest は毎回 store から最新を読み直す。
  // 初回は ensureManifest で空 manifest を作っておく。
  const initial = await eng.store.ensureManifest();
  let seq = initial.maxSeq;
  let idx = nextSegmentIndex(initial.sealed);
  const manifestGeneration = initial.generation;
  let added = 0;
  let segments = 0;
  let cancelled = false;

  // 取り込み対象をバッチに分割。
  const batches: IngestMail[][] = [];
  for (let off = 0; off < fresh.length; off += BATCH) batches.push(fresh.slice(off, off + BATCH));

  // 埋め込みは API レイテンシ律速なので、最大 embedConcurrency 本を並列実行する。
  // ただしセグメント書込/manifest 更新は順序依存 (seq 採番・版管理) なので直列のまま。
  const concurrency = Math.min(10, Math.max(1, s.embedConcurrency || 3));
  const embedOf = (b: IngestMail[]): Promise<Float32Array[]> => {
    // 埋め込みは「件名 + 本文」をまとめて行う:
    // - "Re: ..." 等で本文が引用だけのメールも件名で意味検索が利く
    // - OneNote のチャンクはページタイトル/見出しが件名側に入るので、それも索引される
    // 件名/本文どちらかが空でも組み立てが破綻しないように "(本文なし)" にフォールバック。
    const bodies = b.map(m => {
      const text = m.isHtml ? cleanBody(htmlToText(m.body)) : cleanBody(m.body);
      const subj = (m.subject || '').trim();
      const body = text || '(本文なし)';
      return subj ? `件名: ${subj}\n\n${body}` : body;
    });
    return embedDocsFor(bodies, s, signal);
  };

  // 先読みパイプライン: 先頭 concurrency 本を起動しておき、順番に await→書込→次を起動。
  const inflight = new Map<number, Promise<Float32Array[]>>();
  for (let i = 0; i < Math.min(concurrency, batches.length); i++) inflight.set(i, embedOf(batches[i]));

  let embedded = 0;
  onProgress?.('embed', 0, fresh.length);

  for (let bi = 0; bi < batches.length; bi++) {
    if (signal?.aborted) { cancelled = true; break; }
    const part = batches[bi];

    let vecs: Float32Array[];
    try {
      vecs = await inflight.get(bi)!;
    } catch (e) {
      if (signal?.aborted || (e instanceof Error && e.name === 'AbortError')) { cancelled = true; break; }
      throw e;
    }
    inflight.delete(bi);
    embedded += part.length;
    onProgress?.('embed', embedded, fresh.length);
    const next = bi + concurrency;
    if (next < batches.length && !signal?.aborted) inflight.set(next, embedOf(batches[next]));

    const records: SegmentRecord[] = part.map((m, i) => ({
      seq: ++seq,
      op: 'upsert',
      messageId: m.messageId,
      internetMessageId: m.internetMessageId,
      conversationId: m.conversationId,
      kind: m.kind,
      label: m.label,
      chunkIdx: m.chunkIdx,
      chunkCount: m.chunkCount,
      docPath: m.docPath,
      subject: m.subject,
      from: m.from,
      to: m.to,
      cc: m.cc,
      date: m.date,
      // 表示用は元の本文 (HTML はそのまま、表示時にサニタイズ)。サイズは上限で抑える。
      body: (m.body ?? '').slice(0, m.isHtml ? 30000 : 8000),
      isHtml: !!m.isHtml,
      // PPTX 関連メタ (kind='pptx' のときのみ意味を持つ)。
      pptxFile: m.pptxFile,
      pptxServerRelUrl: m.pptxServerRelUrl,
      slideNo: m.slideNo,
      slideTitle: m.slideTitle,
      thumbServerRelUrl: m.thumbServerRelUrl,
      transcriptFile: m.transcriptFile,
      vttServerRelUrl: m.vttServerRelUrl,
      recordingServerRelUrl: m.recordingServerRelUrl,
      startSec: m.startSec,
      docFile: m.docFile,
      docServerRelUrl: m.docServerRelUrl,
      srcHash: m.srcHash,
      emb: encodeEmbedding(normalize(vecs[i])),
    }));

    // segment 書き込み: overwrite=false で名前衝突を検出し、衝突したら idx を bump。
    // 別 writer が同じ idx で書こうとした場合の名前衝突を Phase 1 ではここで吸収する。
    const placeholder: Segment = { id: segmentId(idx), generation: manifestGeneration, records };
    const { id, idx: confirmedIdx } = await eng.store.writeSegmentNoOverwrite(placeholder, idx);
    idx = confirmedIdx + 1;
    const seg: Segment = { ...placeholder, id };
    await eng.cache.put(id, seg);
    eng.db.applySegment(seg);

    // manifest 更新: CAS (If-Match) で楽観ロック。412 で他 writer の更新を検出したら
    // 最新を再読込してから自分の追加 (seg id + maxSeq + version) を再適用してリトライ。
    const ourId = id;
    const ourSeq = seq;
    const updatedManifest = await updateManifestWithCas(eng.store, m => ({
      ...m,
      sealed: m.sealed.includes(ourId) ? m.sealed : [...m.sealed, ourId],
      maxSeq: Math.max(m.maxSeq, ourSeq),
      version: m.version + 1,
    }));
    await eng.cache.setManifest(updatedManifest);
    // 他 writer が並行で追加したセグメントがあれば、自分の次の idx も追従して bump。
    const latestIdx = nextSegmentIndex(updatedManifest.sealed);
    if (latestIdx > idx) idx = latestIdx;
    // seq も他 writer の進度に追従 (新規 batch は max(local, manifest)+1 から)
    if (updatedManifest.maxSeq > seq) seq = updatedManifest.maxSeq;

    added += records.length;
    segments++;
    onProgress?.('upload', added, fresh.length);
  }

  // 早期終了 (停止/エラー) で未 await の先読み埋め込みが残ると unhandledrejection に
  // なるため、ここで握りつぶしておく。
  for (const p of inflight.values()) p.catch(() => {});

  return { added, skipped, segments, cancelled };
}

/** 指定メールを削除 (tombstone を 1 件のセグメントとして追記)。 */
export async function deleteFromSegments(messageId: string, siteUrl: string): Promise<void> {
  if (!await getLease(siteUrl).ensureWriter()) {
    throw new Error('書き込み権限がありません (他のメンバーが取り込み中)。');
  }
  const eng = await getEngine(siteUrl);
  await eng.sync.sync();
  const initial = await eng.store.ensureManifest();
  const seq = initial.maxSeq + 1;
  const placeholder: Segment = {
    id: segmentId(nextSegmentIndex(initial.sealed)),
    generation: initial.generation,
    records: [{ seq, op: 'delete', messageId }],
  };
  const { id } = await eng.store.writeSegmentNoOverwrite(placeholder, nextSegmentIndex(initial.sealed));
  const seg: Segment = { ...placeholder, id };
  await eng.cache.put(id, seg);
  eng.db.applySegment(seg);
  const updated = await updateManifestWithCas(eng.store, m => ({
    ...m,
    sealed: m.sealed.includes(id) ? m.sealed : [...m.sealed, id],
    maxSeq: Math.max(m.maxSeq, seq),
    version: m.version + 1,
  }));
  await eng.cache.setManifest(updated);
}
