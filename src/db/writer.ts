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
  SEGMENT_CAP, nextSegmentIndex, segmentId,
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
}

export type WritePhase = 'sync' | 'embed' | 'upload';

/** 新規メールを埋め込んでセグメント化し SharePoint へ書き込む。 */
export async function ingestToSegments(
  mails: IngestMail[],
  s: RuntimeSettings,
  siteUrl: string,
  onProgress?: (phase: WritePhase, done: number, total: number) => void,
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
  if (fresh.length === 0) return { added: 0, skipped: mails.length, segments: 0 };

  // 埋め込み (失敗したら書き込まない: ベクトル無しレコードは検索対象外で無意味)。
  // バッチで分割して進捗を出す & 1回の巨大リクエストを避ける。
  const bodies = fresh.map(m => cleanBody(m.body));
  const EMBED_BATCH = 64;
  const vecs: Float32Array[] = [];
  onProgress?.('embed', 0, bodies.length);
  for (let off = 0; off < bodies.length; off += EMBED_BATCH) {
    const part = await embedDocsFor(bodies.slice(off, off + EMBED_BATCH), s);
    for (const v of part) vecs.push(v);
    onProgress?.('embed', Math.min(off + EMBED_BATCH, bodies.length), bodies.length);
  }

  const manifest = await eng.store.ensureManifest();
  let seq = manifest.maxSeq;
  const records: SegmentRecord[] = fresh.map((m, i) => ({
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

  // ≤SEGMENT_CAP 件ずつ新セグメントに分割して UL
  let idx = nextSegmentIndex(manifest.sealed);
  const newIds: string[] = [];
  for (let off = 0; off < records.length; off += SEGMENT_CAP) {
    const id = segmentId(idx++);
    const seg: Segment = { id, generation: manifest.generation, records: records.slice(off, off + SEGMENT_CAP) };
    await eng.store.writeSegment(seg);
    await eng.cache.put(id, seg);
    eng.db.applySegment(seg);
    newIds.push(id);
    onProgress?.('upload', newIds.length, Math.ceil(records.length / SEGMENT_CAP));
  }

  manifest.sealed.push(...newIds);
  manifest.maxSeq = seq;
  manifest.version += 1;
  manifest.updatedAt = new Date().toISOString();
  await eng.store.writeManifest(manifest);
  await eng.cache.setManifest(manifest);

  return { added: fresh.length, skipped: mails.length - fresh.length, segments: newIds.length };
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
