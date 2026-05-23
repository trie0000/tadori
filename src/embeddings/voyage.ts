// Voyage AI 埋め込みクライアント (ブラウザ直接呼び出し)。
// Claude API には埋め込みエンドポイントが無いため、開発者モードでの検索ベクトルは
// Anthropic 推奨の Voyage AI を使う。CORS 対応済み (Access-Control-Allow-Origin: *)。
//
// output_dimension で次元を Tadori の格納次元 (既定 256) に合わせる。
// input_type は query / document を使い分けると検索精度が上がる。

const ENDPOINT = 'https://api.voyageai.com/v1/embeddings';

export interface VoyageConfig {
  voyageApiKey: string;
  voyageModel: string;
  dimensions: number;
}

async function sleepRespectingAbort(ms: number, signal?: AbortSignal): Promise<void> {
  if (ms <= 0) return;
  if (signal?.aborted) throw new DOMException('aborted', 'AbortError');
  await new Promise<void>((resolve, reject) => {
    const t = setTimeout(() => { signal?.removeEventListener('abort', onAbort); resolve(); }, ms);
    const onAbort = (): void => { clearTimeout(t); reject(new DOMException('aborted', 'AbortError')); };
    signal?.addEventListener('abort', onAbort, { once: true });
  });
}

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
  return Math.min(2000 * Math.pow(2, attempt), 30_000);
}

async function call(
  texts: string[],
  cfg: VoyageConfig,
  inputType: 'query' | 'document',
  signal?: AbortSignal,
  maxRetries = 5,
): Promise<Float32Array[]> {
  if (texts.length === 0) return [];
  if (!cfg.voyageApiKey) throw new Error('Voyage API キーが未設定です');

  const body = JSON.stringify({
    input: texts,
    model: cfg.voyageModel,
    input_type: inputType,
    output_dimension: cfg.dimensions,
  });

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    if (signal?.aborted) throw new DOMException('aborted', 'AbortError');
    const res = await fetch(ENDPOINT, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${cfg.voyageApiKey}`,
        'content-type': 'application/json',
      },
      signal,
      body,
    });
    if (res.ok) {
      const json = await res.json() as { data: { index: number; embedding: number[] }[] };
      const out: Float32Array[] = new Array(texts.length);
      for (const d of json.data) out[d.index] = Float32Array.from(d.embedding);
      return out;
    }
    const errBody = await res.text().catch(() => '');
    const retriable = res.status === 429 || (res.status >= 500 && res.status < 600);
    if (!retriable || attempt === maxRetries) {
      throw new Error(`Voyage embed 失敗: HTTP ${res.status} ${errBody.slice(0, 300)}`);
    }
    const waitMs = inferRetryDelayMs(res, errBody, attempt);
    // eslint-disable-next-line no-console
    console.warn(`[voyage] HTTP ${res.status}; retrying in ${Math.round(waitMs / 1000)}s (attempt ${attempt + 1}/${maxRetries})`);
    await sleepRespectingAbort(waitMs, signal);
  }
  throw new Error('Voyage embed: max retries exceeded');
}

export function embedVoyageQuery(text: string, cfg: VoyageConfig): Promise<Float32Array[]> {
  return call([text], cfg, 'query');
}

export function embedVoyageDocs(texts: string[], cfg: VoyageConfig, signal?: AbortSignal): Promise<Float32Array[]> {
  return call(texts, cfg, 'document', signal);
}
