// ブラウザ内ベクトルDB (検索エンジン本体)。
// SharePoint から同期したセグメントを適用してメモリ上にレコードを構築し、
// 総当たり cosine で検索する。1メール=1レコード。
//
// 適用規則: seq 昇順・message-id 単位 last-writer-wins。
//   upsert: そのベクトル/本文で置き換え。delete(tombstone): レコード削除。
//   再適用は seq 比較で no-op になり冪等。
// 検索を ANN(voy 等) に差し替えたくなったら search() だけ置換すればよい。

import { decodeEmbedding } from '../lib/float16';
import { normalize, search as cosineSearch } from '../search/cosine';
import type { Segment, SegmentRecord } from '../sync/segments';

export interface MailRecord {
  messageId: string;
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  date: string;
  body: string;
  vec: Float32Array; // L2 正規化済み
}

export interface DbHit {
  record: MailRecord;
  score: number;
}

export class VectorDb {
  private records = new Map<string, MailRecord>();
  /** message-id ごとに適用済みの最大 seq (削除済みでも保持 = 後続 upsert 復活用)。 */
  private appliedSeq = new Map<string, number>();
  private maxSeq = 0;

  get size(): number { return this.records.size; }
  get watermark(): number { return this.maxSeq; }

  applySegment(seg: Segment): void {
    const recs = [...seg.records].sort((a, b) => a.seq - b.seq);
    for (const r of recs) this.applyRecord(r);
  }

  applyRecord(r: SegmentRecord): void {
    const prev = this.appliedSeq.get(r.messageId) ?? 0;
    if (r.seq <= prev) return; // 古い → last-writer-wins で無視 (冪等)
    if (r.op === 'delete') {
      this.records.delete(r.messageId);
    } else {
      if (!r.emb) return; // upsert は埋め込み必須
      this.records.set(r.messageId, {
        messageId: r.messageId,
        subject: r.subject ?? '(件名なし)',
        from: r.from ?? '',
        to: r.to ?? [],
        cc: r.cc ?? [],
        date: r.date ?? '',
        body: r.body ?? '',
        vec: normalize(decodeEmbedding(r.emb)),
      });
    }
    this.appliedSeq.set(r.messageId, r.seq);
    if (r.seq > this.maxSeq) this.maxSeq = r.seq;
  }

  /** 正規化済みクエリベクトルで Top-K (総当たり cosine)。 */
  search(qvec: Float32Array, topK: number): DbHit[] {
    const q = normalize(qvec);
    const index = Array.from(this.records.values(), r => ({ messageId: r.messageId, vec: r.vec }));
    const hits = cosineSearch(q, index, topK);
    const out: DbHit[] = [];
    for (const h of hits) {
      const rec = this.records.get(h.messageId);
      if (rec) out.push({ record: rec, score: h.score });
    }
    return out;
  }

  clear(): void {
    this.records.clear();
    this.appliedSeq.clear();
    this.maxSeq = 0;
  }
}
