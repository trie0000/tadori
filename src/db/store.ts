// ブラウザ内ベクトルDB (検索エンジン本体)。
// SharePoint から同期したセグメントを適用してメモリ上にレコードを構築し、
// 総当たり cosine で検索する。1メール=1レコード。
//
// 適用規則: seq 昇順・message-id 単位 last-writer-wins。
//   upsert: そのベクトル/本文で置き換え。delete(tombstone): レコード削除。
//   再適用は seq 比較で no-op になり冪等。
// 検索を ANN(voy 等) に差し替えたくなったら search() だけ置換すればよい。

import { decodeEmbedding } from '../lib/float16';
import { normalize } from '../search/cosine';
import { htmlToText } from '../lib/mailhtml';
import { cleanBody } from '../lib/mailtext';
import type { Segment, SegmentRecord } from '../sync/segments';

/** 文字 2-gram の集合 (日本語は空白区切りが無いので char bigram で一致を取る)。 */
function bigrams(text: string): Set<string> {
  const t = (text || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const out = new Set<string>();
  for (let i = 0; i < t.length - 1; i++) out.add(t.slice(i, i + 2));
  return out;
}

/** クエリ bigram のうちドキュメントに含まれる割合 (0..1)。 */
function keywordCoverage(query: Set<string>, doc: Set<string>): number {
  if (query.size === 0) return 0;
  let hit = 0;
  for (const g of query) if (doc.has(g)) hit++;
  return hit / query.size;
}

export interface MailRecord {
  messageId: string;
  internetMessageId: string;
  conversationId: string;
  kind: 'mail' | 'onenote' | 'doc' | 'pptx' | 'transcript';
  chunkIdx?: number;
  chunkCount?: number;
  docPath?: string;
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  date: string;
  body: string;
  isHtml: boolean;
  vec: Float32Array; // L2 正規化済み
  /** PPTX 取り込みのメタ (kind='pptx' のときのみ意味を持つ)。 */
  pptxFile?: string;
  pptxServerRelUrl?: string;
  slideNo?: number;
  slideTitle?: string;
  thumbServerRelUrl?: string;
  /** Teams 文字起こしのメタ (kind='transcript' のときのみ意味を持つ)。 */
  transcriptFile?: string;
  vttServerRelUrl?: string;
  recordingServerRelUrl?: string;
  startSec?: number;
  /** ドキュメント (kind='doc') のメタ。 */
  docFile?: string;
  docServerRelUrl?: string;
  /** ソース内容ハッシュ (差分判定用)。 */
  srcHash?: string;
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

  has(messageId: string): boolean { return this.records.has(messageId); }

  /** messageId でレコードを取得 (無ければ undefined)。PPTX 差分判定で srcHash を引く用。 */
  get(messageId: string): MailRecord | undefined { return this.records.get(messageId); }

  /** 指定 conversationId (= pptx の serverRelativeUrl) に属する messageId 一覧。 */
  messageIdsForConversation(conversationId: string): string[] {
    const out: string[] = [];
    for (const r of this.records.values()) {
      if (r.conversationId === conversationId) out.push(r.messageId);
    }
    return out;
  }

  /** 種別ごとのレコード数 (診断用)。例: { mail: 1200, doc: 30, pptx: 88 } */
  kindCounts(): Record<string, number> {
    const c: Record<string, number> = {};
    for (const r of this.records.values()) {
      const k = r.kind ?? 'mail';
      c[k] = (c[k] ?? 0) + 1;
    }
    return c;
  }

  /** kind ごとのベクトル次元分布 (診断用)。例: { mail: {1024:1200}, doc: {256:30} } */
  dimByKind(): Record<string, Record<number, number>> {
    const out: Record<string, Record<number, number>> = {};
    for (const r of this.records.values()) {
      const k = r.kind ?? 'mail';
      (out[k] ??= {});
      out[k][r.vec.length] = (out[k][r.vec.length] ?? 0) + 1;
    }
    return out;
  }

  /** kind='doc' レコードの docServerRelUrl サンプル (診断用、最大 n 件)。 */
  sampleDocUrls(n = 3): string[] {
    const out: string[] = [];
    for (const r of this.records.values()) {
      if (r.kind === 'doc') { out.push(r.docServerRelUrl || r.conversationId || '(空)'); if (out.length >= n) break; }
    }
    return out;
  }

  /** 直近 applySegment の落とし込み統計 (診断用)。 */
  lastApply: { records: number; applied: number; droppedOldSeq: number; droppedNoEmb: number; deletes: number; byKind: Record<string, number> } =
    { records: 0, applied: 0, droppedOldSeq: 0, droppedNoEmb: 0, deletes: 0, byKind: {} };

  applySegment(seg: Segment): void {
    const recs = [...seg.records].sort((a, b) => a.seq - b.seq);
    const st = { records: recs.length, applied: 0, droppedOldSeq: 0, droppedNoEmb: 0, deletes: 0, byKind: {} as Record<string, number> };
    for (const r of recs) {
      const k = r.kind ?? 'mail';
      st.byKind[k] = (st.byKind[k] ?? 0) + 1;
      const res = this.applyRecord(r);
      if (res === 'oldseq') st.droppedOldSeq++;
      else if (res === 'noemb') st.droppedNoEmb++;
      else if (res === 'delete') st.deletes++;
      else st.applied++;
    }
    this.lastApply = st;
  }

  /** 適用結果コード: '' = upsert 適用 / 'oldseq' / 'noemb' / 'delete'。 */
  applyRecord(r: SegmentRecord): '' | 'oldseq' | 'noemb' | 'delete' {
    const prev = this.appliedSeq.get(r.messageId) ?? 0;
    if (r.seq <= prev) return 'oldseq'; // 古い → last-writer-wins で無視 (冪等)
    this.kwCache.delete(r.messageId); // 内容が変わる → キーワード索引を無効化
    if (r.op === 'delete') {
      this.records.delete(r.messageId);
      this.appliedSeq.set(r.messageId, r.seq);
      if (r.seq > this.maxSeq) this.maxSeq = r.seq;
      return 'delete';
    } else {
      if (!r.emb) return 'noemb'; // upsert は埋め込み必須
      this.records.set(r.messageId, {
        messageId: r.messageId,
        internetMessageId: r.internetMessageId ?? '',
        conversationId: r.conversationId ?? '',
        kind: r.kind ?? 'mail',
        chunkIdx: r.chunkIdx,
        chunkCount: r.chunkCount,
        docPath: r.docPath,
        subject: r.subject ?? '(件名なし)',
        from: r.from ?? '',
        to: r.to ?? [],
        cc: r.cc ?? [],
        date: r.date ?? '',
        body: r.body ?? '',
        isHtml: r.isHtml ?? false,
        vec: normalize(decodeEmbedding(r.emb)),
        pptxFile: r.pptxFile,
        pptxServerRelUrl: r.pptxServerRelUrl,
        slideNo: r.slideNo,
        slideTitle: r.slideTitle,
        thumbServerRelUrl: r.thumbServerRelUrl,
        transcriptFile: r.transcriptFile,
        vttServerRelUrl: r.vttServerRelUrl,
        recordingServerRelUrl: r.recordingServerRelUrl,
        startSec: r.startSec,
        docFile: r.docFile,
        docServerRelUrl: r.docServerRelUrl,
        srcHash: r.srcHash,
      });
    }
    this.appliedSeq.set(r.messageId, r.seq);
    if (r.seq > this.maxSeq) this.maxSeq = r.seq;
  }

  /** Top-K 検索。keywordWeight>0 ならハイブリッド (ベクトル + 文字bigramキーワード) の
   *  加重和でスコアリング (両方 0..1 に正規化)。score は 0..1 で閾値判定にそのまま使える。 */
  search(qvec: Float32Array, topK: number, queryText = '', keywordWeight = 0): DbHit[] {
    const q = normalize(qvec);
    const dim = q.length;
    const recs = [...this.records.values()];
    const useKw = keywordWeight > 0 && queryText.trim().length > 0;
    const qbi = useKw ? bigrams(queryText) : null;
    const w = Math.min(1, Math.max(0, keywordWeight));

    // 診断: kind ごとの「ベクトル次元」分布と、クエリ次元と一致しない (= 類似度 0 に
    //       沈む) 件数を集計してログ出力する。doc/pdf がヒットしない切り分け用。
    const dimByKind: Record<string, Record<number, number>> = {};
    let dimMismatch = 0;
    const mismatchByKind: Record<string, number> = {};

    const scored = recs.map(r => {
      const k = r.kind ?? 'mail';
      (dimByKind[k] ??= {})[r.vec.length] = ((dimByKind[k] ??= {})[r.vec.length] ?? 0) + 1;
      let dot = 0;
      if (r.vec.length === dim) { for (let i = 0; i < dim; i++) dot += q[i] * r.vec[i]; }
      else { dimMismatch++; mismatchByKind[k] = (mismatchByKind[k] ?? 0) + 1; }
      const vcos = Math.max(0, dot); // cosine を 0..1 に
      if (!qbi) return { record: r, score: vcos };
      const kcov = keywordCoverage(qbi, this.kwIndex(r));
      return { record: r, score: (1 - w) * vcos + w * kcov };
    });

    console.log('[tadori] db.search 診断:', {
      queryDim: dim,
      kindごとの次元分布: dimByKind,
      次元不一致でベクトル0になった件数: dimMismatch,
      次元不一致の内訳: mismatchByKind,
    });
    if (dimMismatch > 0) {
      console.warn('[tadori] ★ 一部レコードがクエリと次元不一致 → 類似度 0 で沈みます。' +
        '設定→AI設定の「次元数」が取り込み当時と違う可能性。該当ソースを強制再取り込みで揃えてください。');
    }

    scored.sort((a, b) => b.score - a.score);
    return scored.slice(0, topK);
  }

  // ── キーワード索引 (文字bigram)。レコード単位でキャッシュ ──
  private kwCache = new Map<string, Set<string>>();
  private kwIndex(r: MailRecord): Set<string> {
    let s = this.kwCache.get(r.messageId);
    if (!s) {
      // 引用履歴 / HTML タグはランキングに混ぜない (埋め込み側と同じ前処理)。
      const text = r.isHtml ? htmlToText(r.body) : r.body;
      s = bigrams(`${r.subject} ${cleanBody(text)}`);
      this.kwCache.set(r.messageId, s);
    }
    return s;
  }

  /** 取り込み済みの OneNote ページ ID 集合 (kind=onenote の conversationId)。
   *  取り込みペインで「既に取り込み済みか」を判定するのに使う。 */
  importedOneNotePageIds(): Set<string> {
    const out = new Set<string>();
    for (const r of this.records.values()) {
      if (r.kind === 'onenote' && r.conversationId) out.add(r.conversationId);
    }
    return out;
  }

  /** 取り込み済み OneNote ページ一覧 (1 ページ = 1 エントリ。複数チャンクは集約)。
   *  「OneNote に追記」モーダルの追記先ピッカーで使う。 */
  importedOneNotePages(): Array<{ pageId: string; title: string; location: string; lastModified: string }> {
    const seen = new Map<string, { pageId: string; title: string; location: string; lastModified: string }>();
    for (const r of this.records.values()) {
      if (r.kind !== 'onenote' || !r.conversationId) continue;
      const ex = seen.get(r.conversationId);
      // チャンク 0 の subject = ページタイトル素のまま。それ以外は "title - heading" になっているので
      // chunkIdx === 0 を優先採用。なければそのまま入れる。
      if (!ex || (r.chunkIdx ?? 0) === 0) {
        seen.set(r.conversationId, {
          pageId: r.conversationId,
          title: (r.chunkIdx ?? 0) === 0 ? r.subject : (ex?.title ?? r.subject),
          location: r.from,
          lastModified: r.date,
        });
      }
    }
    return [...seen.values()].sort((a, b) => a.location.localeCompare(b.location) || a.title.localeCompare(b.title));
  }

  /** 同一スレッド (conversationId) のレコードを受信日時の昇順で返す。 */
  byConversation(conversationId: string): MailRecord[] {
    if (!conversationId) return [];
    const out: MailRecord[] = [];
    for (const r of this.records.values()) if (r.conversationId === conversationId) out.push(r);
    out.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
    return out;
  }

  clear(): void {
    this.records.clear();
    this.appliedSeq.clear();
    this.kwCache.clear();
    this.maxSeq = 0;
  }
}
