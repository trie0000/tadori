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

async function call(
  texts: string[],
  cfg: VoyageConfig,
  inputType: 'query' | 'document',
  signal?: AbortSignal,
): Promise<Float32Array[]> {
  if (texts.length === 0) return [];
  if (!cfg.voyageApiKey) throw new Error('Voyage API キーが未設定です');

  const res = await fetch(ENDPOINT, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${cfg.voyageApiKey}`,
      'content-type': 'application/json',
    },
    signal,
    body: JSON.stringify({
      input: texts,
      model: cfg.voyageModel,
      input_type: inputType,
      output_dimension: cfg.dimensions,
    }),
  });

  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`Voyage embed 失敗: HTTP ${res.status} ${body.slice(0, 300)}`);
  }

  const json = await res.json() as { data: { index: number; embedding: number[] }[] };
  const out: Float32Array[] = new Array(texts.length);
  for (const d of json.data) out[d.index] = Float32Array.from(d.embedding);
  return out;
}

export function embedVoyageQuery(text: string, cfg: VoyageConfig): Promise<Float32Array[]> {
  return call([text], cfg, 'query');
}

export function embedVoyageDocs(texts: string[], cfg: VoyageConfig, signal?: AbortSignal): Promise<Float32Array[]> {
  return call(texts, cfg, 'document', signal);
}
