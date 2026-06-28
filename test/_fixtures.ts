// ローカル検証用フィクスチャ。本番の埋め込み/SP/AI には依存しない決定論データ。
// 埋め込みは seed から決定論的に生成し、検索ランキングを再現可能にする。

import { normalize } from '../src/search/cosine';
import { encodeEmbedding } from '../src/lib/float16';
import type { SegmentRecord, Segment } from '../src/sync/segments';

export const DIM = 1024;

/** seed から決定論的な DIM 次元ベクトル (正規化前)。 */
export function vec(seed: number, dim = DIM): Float32Array {
  const a = new Float32Array(dim);
  let s = (seed * 2654435761) >>> 0;
  for (let i = 0; i < dim; i++) { s = (s * 1103515245 + 12345) >>> 0; a[i] = (s / 0xffffffff) - 0.5; }
  return a;
}

export interface MakeRecOpts {
  i: number;
  kind: 'mail' | 'doc' | 'pptx' | 'transcript' | 'onenote';
  folder?: string;
  dim?: number;
  seqBase?: number;
}

/** 本番 writer と同じ形の SegmentRecord を作る (emb は decode 可能な実バイト)。 */
export function makeRecord(o: MakeRecOpts): SegmentRecord {
  const dim = o.dim ?? DIM;
  const folder = o.folder ?? '/sites/x/Shared Documents/F';
  const seedOffset = { mail: 0, doc: 1000, pptx: 5000, transcript: 8000, onenote: 9000 }[o.kind];
  const v = normalize(vec(o.i + seedOffset, dim));
  const path = `${o.kind}://${folder}/f${o.i}#0`;
  return {
    seq: (o.seqBase ?? 0) + o.i + 1,
    op: 'upsert',
    messageId: path,
    conversationId: `${folder}/f${o.i}`,
    kind: o.kind,
    chunkIdx: 0,
    chunkCount: 1,
    subject: `${o.kind} ${o.i}`,
    from: 'tester',
    to: [], cc: [],
    date: '2026-01-01T00:00:00Z',
    body: `本文 ${o.kind} ${o.i} キーワード${o.i}`,
    isHtml: false,
    docServerRelUrl: o.kind === 'doc' ? `${folder}/f${o.i}.pdf` : undefined,
    docFile: o.kind === 'doc' ? `f${o.i}.pdf` : undefined,
    emb: encodeEmbedding(v),
  };
}

export function makeSegment(id: string, records: SegmentRecord[]): Segment {
  return { id, generation: 0, records };
}
