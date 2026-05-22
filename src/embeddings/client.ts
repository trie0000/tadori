// Azure OpenAI 埋め込みクライアント。PoC 02 で動作確認済みの経路。
// ブラウザ → 中継サーバ (loopback) → 社内プロキシ → Azure OpenAI。

import type { TadoriConfig } from '../config';

export interface EmbedAuth {
  /** api-key ヘッダ方式 (サブスクリプションキー)。 */
  apiKey?: string;
  /** AAD トークン方式 (Authorization: Bearer ...)。 */
  bearer?: string;
}

/** 複数テキストをまとめて埋め込み、Float32Array の配列で返す。
 *  入力順と出力順は Azure OpenAI 仕様で対応 (data[].index で保証)。 */
export async function embedTexts(
  texts: string[],
  cfg: TadoriConfig,
  auth: EmbedAuth,
): Promise<Float32Array[]> {
  if (texts.length === 0) return [];

  const url = `${cfg.relayBaseUrl.replace(/\/+$/, '')}`
    + `/openai/deployments/${encodeURIComponent(cfg.embeddingDeployment)}`
    + `/embeddings?api-version=${encodeURIComponent(cfg.apiVersion)}`;

  const headers: Record<string, string> = { 'Content-Type': 'application/json' };
  if (auth.apiKey) headers['api-key'] = auth.apiKey;
  if (auth.bearer) headers['Authorization'] = auth.bearer.startsWith('Bearer ') ? auth.bearer : `Bearer ${auth.bearer}`;

  const res = await fetch(url, {
    method: 'POST',
    headers,
    credentials: 'omit',
    body: JSON.stringify({ input: texts, dimensions: cfg.dimensions }),
  });

  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`embed failed: HTTP ${res.status} ${body.slice(0, 300)}`);
  }

  const json = await res.json() as { data: { index: number; embedding: number[] }[] };
  const out: Float32Array[] = new Array(texts.length);
  for (const d of json.data) out[d.index] = Float32Array.from(d.embedding);
  return out;
}

/** 単一クエリの埋め込み (検索時に使用)。 */
export async function embedQuery(
  text: string,
  cfg: TadoriConfig,
  auth: EmbedAuth,
): Promise<Float32Array> {
  const [v] = await embedTexts([text], cfg, auth);
  return v;
}
