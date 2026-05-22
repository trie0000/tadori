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
import {
  nextSegmentIndex, segmentId,
  type Segment, type SegmentRecord,
} from '../sync/segments';
import type { RuntimeSettings } from '../api/aiSettings';

export interface IngestMail {
  messageId: string;
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  date: string;
  body: string;
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

  const manifest = await eng.store.ensureManifest();
  let seq = manifest.maxSeq;
  let idx = nextSegmentIndex(manifest.sealed);
  let added = 0;
  let segments = 0;
  let cancelled = false;

  // バッチ単位で「埋め込み → セグメント書込 → manifest 更新」を確定。
  for (let off = 0; off < fresh.length; off += BATCH) {
    if (signal?.aborted) { cancelled = true; break; }
    const part = fresh.slice(off, off + BATCH);
    // 本文が空 (引用/署名のみ等) のメールは埋め込み API が空文字を拒否するため、
    // 件名にフォールバックして空文字を送らない。
    const bodies = part.map(m => cleanBody(m.body) || (m.subject || '').trim() || '(本文なし)');

    let vecs: Float32Array[];
    try {
      vecs = await embedDocsFor(bodies, s, signal);
    } catch (e) {
      if (signal?.aborted || (e instanceof Error && e.name === 'AbortError')) { cancelled = true; break; }
      throw e;
    }

    const records: SegmentRecord[] = part.map((m, i) => ({
      seq: ++seq,
      op: 'upsert',
      messageId: m.messageId,
      subject: m.subject,
      from: m.from,
      to: m.to,
      cc: m.cc,
      date: m.date,
      body: bodies[i],
      emb: encodeEmbedding(normalize(vecs[i])),
    }));

    const id = segmentId(idx++);
    const seg: Segment = { id, generation: manifest.generation, records };
    await eng.store.writeSegment(seg);
    await eng.cache.put(id, seg);
    eng.db.applySegment(seg);

    manifest.sealed.push(id);
    manifest.maxSeq = seq;
    manifest.version += 1;
    manifest.updatedAt = new Date().toISOString();
    await eng.store.writeManifest(manifest);
    await eng.cache.setManifest(manifest);

    added += records.length;
    segments++;
    onProgress?.('upload', added, fresh.length);
  }

  return { added, skipped, segments, cancelled };
}

/** 指定メールを削除 (tombstone を 1 件のセグメントとして追記)。 */
export async function deleteFromSegments(messageId: string, siteUrl: string): Promise<void> {
  if (!await getLease(siteUrl).ensureWriter()) {
    throw new Error('書き込み権限がありません (他のメンバーが取り込み中)。');
  }
  const eng = await getEngine(siteUrl);
  await eng.sync.sync();
  const manifest = await eng.store.ensureManifest();
  const seq = manifest.maxSeq + 1;
  const id = segmentId(nextSegmentIndex(manifest.sealed));
  const seg: Segment = { id, generation: manifest.generation, records: [{ seq, op: 'delete', messageId }] };
  await eng.store.writeSegment(seg);
  await eng.cache.put(id, seg);
  eng.db.applySegment(seg);
  manifest.sealed.push(id);
  manifest.maxSeq = seq;
  manifest.version += 1;
  manifest.updatedAt = new Date().toISOString();
  await eng.store.writeManifest(manifest);
  await eng.cache.setManifest(manifest);
}
