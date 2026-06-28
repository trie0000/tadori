// テスト用 embeddings ルータ・スタブ。クエリ埋め込みをテストが注入したベクトルで返す
// (Azure/Voyage に触れない)。embedDocsFor は使わない想定だがダミーを置く。

let _q: Float32Array = new Float32Array();
export function __setQuery(v: Float32Array): void { _q = v; }

export async function embedQueryFor(_text: string, _s: any, _signal?: AbortSignal): Promise<Float32Array> {
  return _q;
}
export async function embedDocsFor(texts: string[], _s: any, _signal?: AbortSignal): Promise<Float32Array[]> {
  return texts.map(() => new Float32Array());
}
