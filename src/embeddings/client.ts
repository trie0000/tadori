// Azure OpenAI 埋め込みクライアント。PoC 02 で動作確認済みの経路。
// ブラウザ → 中継サーバ (loopback) → 社内プロキシ → Azure OpenAI。

/** 埋め込みに必要な設定の最小集合 (RuntimeSettings が構造的に満たす)。 */
export interface EmbedConfig {
  relayBaseUrl: string;
  embeddingDeployment: string;
  apiVersion: string;
  dimensions: number;
}

export interface EmbedAuth {
  /** api-key ヘッダ方式 (サブスクリプションキー)。 */
  apiKey?: string;
  /** AAD トークン方式 (Authorization: Bearer ...)。 */
  bearer?: string;
}

/** AbortSignal を尊重しつつ ms 待機。abort されたら例外で抜ける。 */
async function sleepRespectingAbort(ms: number, signal?: AbortSignal): Promise<void> {
  if (ms <= 0) return;
  if (signal?.aborted) throw new DOMException('aborted', 'AbortError');
  await new Promise<void>((resolve, reject) => {
    const t = setTimeout(() => { signal?.removeEventListener('abort', onAbort); resolve(); }, ms);
    const onAbort = (): void => { clearTimeout(t); reject(new DOMException('aborted', 'AbortError')); };
    signal?.addEventListener('abort', onAbort, { once: true });
  });
}

/** レスポンスから「次のリトライまでの待機秒数」を推定。
 *  - Retry-After ヘッダ (秒 or HTTP-date) を最優先
 *  - 本文の "try again in N seconds" / "retry after N" 等から抽出
 *  - 何も取れなければ指数バックオフのフォールバック値 (attempt 依存) */
function inferRetryDelayMs(res: Response, bodyText: string, attempt: number): number {
  const ra = res.headers.get('Retry-After');
  if (ra) {
    const n = Number(ra);
    if (!isNaN(n) && n >= 0) return Math.min(n * 1000, 120_000);
    const d = Date.parse(ra);
    if (!isNaN(d)) return Math.max(0, Math.min(d - Date.now(), 120_000));
  }
  const m = bodyText.match(/(?:try again in|retry (?:after|in))\s+(\d+)\s*(?:s|sec|seconds)?/i);
  if (m) return Math.min(Number(m[1]) * 1000, 120_000);
  // フォールバック: 指数バックオフ (2s → 4s → 8s → 16s → 30s)
  return Math.min(2000 * Math.pow(2, attempt), 30_000);
}

/** 複数テキストをまとめて埋め込み、Float32Array の配列で返す。
 *  入力順と出力順は Azure OpenAI 仕様で対応 (data[].index で保証)。
 *  429 (rate limit) と 5xx は最大 maxRetries 回まで自動リトライする。
 *  Retry-After ヘッダ or 本文 "try again in N seconds" を尊重。 */
export async function embedTexts(
  texts: string[],
  cfg: EmbedConfig,
  auth: EmbedAuth,
  signal?: AbortSignal,
  maxRetries = 5,
): Promise<Float32Array[]> {
  if (texts.length === 0) return [];

  const url = `${cfg.relayBaseUrl.replace(/\/+$/, '')}`
    + `/openai/deployments/${encodeURIComponent(cfg.embeddingDeployment)}`
    + `/embeddings?api-version=${encodeURIComponent(cfg.apiVersion)}`;

  const headers: Record<string, string> = { 'Content-Type': 'application/json' };
  if (auth.apiKey) headers['api-key'] = auth.apiKey;
  if (auth.bearer) headers['Authorization'] = auth.bearer.startsWith('Bearer ') ? auth.bearer : `Bearer ${auth.bearer}`;
  const body = JSON.stringify({ input: texts, dimensions: cfg.dimensions });

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    if (signal?.aborted) throw new DOMException('aborted', 'AbortError');
    const res = await fetch(url, { method: 'POST', headers, credentials: 'omit', signal, body });
    if (res.ok) {
      const json = await res.json() as { data: { index: number; embedding: number[] }[] };
      const out: Float32Array[] = new Array(texts.length);
      for (const d of json.data) out[d.index] = Float32Array.from(d.embedding);
      return out;
    }
    const errBody = await res.text().catch(() => '');
    const retriable = res.status === 429 || (res.status >= 500 && res.status < 600);
    if (!retriable || attempt === maxRetries) {
      throw new Error(`embed failed: HTTP ${res.status} ${errBody.slice(0, 300)}`);
    }
    const waitMs = inferRetryDelayMs(res, errBody, attempt);
    // ログは console.warn で出して、呼び出し側からも見えるようにする
    // eslint-disable-next-line no-console
    console.warn(`[embed] HTTP ${res.status}; retrying in ${Math.round(waitMs/1000)}s (attempt ${attempt + 1}/${maxRetries})`);
    await sleepRespectingAbort(waitMs, signal);
  }
  // ループ抜け防止 (理論上到達しない)
  throw new Error('embed failed: max retries exceeded');
}
