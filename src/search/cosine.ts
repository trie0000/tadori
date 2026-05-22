// ブルートフォース cosine 類似度検索 (ADR-008)。
// 37,500 件規模では 30〜80ms で実用十分。100 万件超で WASM ANN を再検討。

export interface IndexedVector {
  messageId: string;
  /** 正規化済みベクトル (事前に L2 normalize して内積=cosine にする)。 */
  vec: Float32Array;
}

export interface SearchHit {
  messageId: string;
  score: number;
}

/** L2 正規化。検索前にインデックス側・クエリ側の両方を正規化しておくと、
 *  内積がそのまま cosine 類似度になりループが軽くなる。 */
export function normalize(vec: Float32Array): Float32Array {
  let sum = 0;
  for (let i = 0; i < vec.length; i++) sum += vec[i] * vec[i];
  const norm = Math.sqrt(sum) || 1;
  const out = new Float32Array(vec.length);
  for (let i = 0; i < vec.length; i++) out[i] = vec[i] / norm;
  return out;
}

/** 正規化済みクエリと正規化済みインデックスの Top-K を返す。 */
export function search(
  query: Float32Array,
  index: IndexedVector[],
  topK = 20,
): SearchHit[] {
  const dim = query.length;
  const hits: SearchHit[] = [];
  for (const item of index) {
    const v = item.vec;
    if (v.length !== dim) continue;
    let dot = 0;
    for (let i = 0; i < dim; i++) dot += query[i] * v[i];
    hits.push({ messageId: item.messageId, score: dot });
  }
  hits.sort((a, b) => b.score - a.score);
  return hits.slice(0, topK);
}
