// AI / Tadori 設定の永続化 (localStorage)。
//
// ★ AI 接続設定は Spira と「方式ごと統一」する ★
// Tadori は Spira と同じ SharePoint オリジン上で動くため localStorage を共有する。
// キー名・モデルIDの扱い・デプロイ名の組み立て方を Spira (src/api/aiSettings.ts) と
// 完全に揃え、どちらのツールで設定しても両方に効くようにする。
//
// 社内 AI (corp / Azure OpenAI 互換):
//   - corp:model はモデルID (例 gpt-4.1-mini) を保存。デプロイ名そのものではない。
//   - デプロイ名 = `<prefix><モデル名(.除去)>` (deploymentIdFor)。
//   - apiVersion はモデル別に導出 (reasoning 系 → preview)。overrides で上書き可。
//   - 個人上書き (key) が無ければ `<key>:default` を読む (Spira が共有デフォルトを
//     起動時に展開したキャッシュ)。Tadori は読むだけ (書き込みはしない)。
//
// provider='claude' は開発者モード限定。OFF のときは corp に丸める。

import { isDeveloperMode } from '../utils/devMode';

export type Provider = 'corp' | 'claude';

export interface CorpAiModel {
  id: string;
  /** reasoning 系は max_completion_tokens / preview apiVersion を使う。 */
  reasoning: boolean;
}

// Spira の CORP_AI_MODELS と同一。
export const CORP_AI_MODELS: CorpAiModel[] = [
  { id: 'gpt-5',        reasoning: true  },
  { id: 'gpt-5-mini',   reasoning: true  },
  { id: 'gpt-5-nano',   reasoning: true  },
  { id: 'o3',           reasoning: true  },
  { id: 'o4-mini',      reasoning: true  },
  { id: 'gpt-4.1',      reasoning: false },
  { id: 'gpt-4.1-mini', reasoning: false },
  { id: 'gpt-4.1-nano', reasoning: false },
  { id: 'gpt-4o',       reasoning: false },
  { id: 'gpt-4o-mini',  reasoning: false },
];

export const CLAUDE_MODELS: Array<{ id: string; label: string }> = [
  { id: 'claude-opus-4-5',   label: 'Claude Opus 4.5' },
  { id: 'claude-sonnet-4-5', label: 'Claude Sonnet 4.5' },
  { id: 'claude-haiku-4-5',  label: 'Claude Haiku 4.5' },
];

// 埋め込みは Tadori 固有 (Spira は埋め込みを使わない)。デプロイ名は corp と同じ
// prefix 方式で組み立てる。
export const EMBEDDING_MODELS: string[] = [
  'text-embedding-3-small',
  'text-embedding-3-large',
  'text-embedding-ada-002',
];

const DEFAULT_PROVIDER: Provider = 'corp';
export const DEFAULT_CORP_MODEL = 'gpt-4.1-mini';
export const DEFAULT_CLAUDE_MODEL = 'claude-sonnet-4-5';
export const DEFAULT_VOYAGE_MODEL = 'voyage-3.5-lite';
export const DEFAULT_EMBEDDING_MODEL = 'text-embedding-3-large';
const DEFAULT_EMBEDDING_API_VERSION = '2024-02-01';

const KEY = {
  // Spira と共有
  provider:         'spira:ai:provider',
  corpKey:          'spira:ai:corp:key',
  corpModel:        'spira:ai:corp:model',
  corpBaseUrl:      'spira:ai:corp:base-url',
  corpDeployPrefix: 'spira:ai:corp:deploy-prefix',
  corpOverrides:    'spira:ai:corp:overrides',
  claudeKey:        'spira:ai:claude:key',
  claudeModel:      'spira:ai:claude:model',
  // Tadori 固有
  embeddingModel:      'tadori:embedding-model',
  embeddingApiVersion: 'tadori:api-version',
  voyageKey:           'tadori:voyage:key',
  voyageModel:         'tadori:voyage:model',
  dimensions:          'tadori:dimensions',
  listTitle:           'tadori:list-title',
  mlAddresses:         'tadori:ml-addresses',
  ingestIntervalSec:   'tadori:ingest-interval-sec',
  embedConcurrency:    'tadori:embed-concurrency',
  ragTopK:             'tadori:rag-topk',
  ragMinScore:         'tadori:rag-min-score',
  ragKeywordWeight:    'tadori:rag-keyword-weight',
  enterSends:          'tadori:enter-sends',
  rerankEnabled:       'tadori:rerank-enabled',
  rerankCandidates:    'tadori:rerank-candidates',
} as const;

const DEFAULT_SUFFIX = ':default';

function lsGet(k: string): string {
  try { return localStorage.getItem(k) ?? ''; } catch { return ''; }
}
function lsSet(k: string, v: string): void {
  try { localStorage.setItem(k, v); } catch { /* quota */ }
}
/** 個人上書き → 共有デフォルトキャッシュ (Spira が展開) の順で読む。 */
function lsGetEff(k: string): string {
  return lsGet(k) || lsGet(k + DEFAULT_SUFFIX);
}

// ─── corp デプロイ名 / エンドポイント解決 (Spira と同一ロジック) ───────────────

export function findCorpAiModel(modelId: string): CorpAiModel | null {
  return CORP_AI_MODELS.find(m => m.id === modelId) ?? null;
}

export function deploymentIdFor(modelId: string): string {
  const prefix = lsGetEff(KEY.corpDeployPrefix);
  return prefix + modelId.replace(/\./g, '');
}

interface CorpOverride { baseUrl?: string; apiVersion?: string; deploymentId?: string; }
function getOverrides(): Record<string, CorpOverride> {
  const raw = lsGetEff(KEY.corpOverrides);
  if (!raw) return {};
  try {
    const o = JSON.parse(raw);
    if (o && typeof o === 'object') return o as Record<string, CorpOverride>;
  } catch { /* ignore */ }
  return {};
}

function corpBaseUrl(): string {
  return lsGetEff(KEY.corpBaseUrl).replace(/\/$/, '');
}

interface ResolvedEndpoint { baseUrl: string; apiVersion: string; deploymentId: string; }
function resolveCorpChat(modelId: string): ResolvedEndpoint {
  const m = findCorpAiModel(modelId);
  const defaultApiVersion = m?.reasoning ? '2024-12-01-preview' : '2024-06-01';
  const ov = getOverrides()[modelId] || {};
  return {
    baseUrl: (ov.baseUrl || corpBaseUrl()).replace(/\/$/, ''),
    apiVersion: ov.apiVersion || defaultApiVersion,
    deploymentId: ov.deploymentId || deploymentIdFor(modelId),
  };
}

function getCorpModel(): string {
  const stored = lsGetEff(KEY.corpModel);
  if (stored && CORP_AI_MODELS.some(m => m.id === stored)) return stored;
  return DEFAULT_CORP_MODEL;
}

function getEmbeddingModel(): string {
  return lsGetEff(KEY.embeddingModel) || DEFAULT_EMBEDDING_MODEL;
}

// ─── RuntimeSettings ───────────────────────────────────────────────────────

export interface RuntimeSettings {
  provider: Provider;

  // corp 生値 (UI 往復用)
  apiKey: string;
  chatModel: string;
  corpDeployPrefix: string;
  corpOverridesRaw: string;
  embeddingModel: string;

  // corp 導出値 (リクエスト用)
  chatBaseUrl: string;       // チャット URL のベース (per-model override 反映)
  chatDeployment: string;    // チャットデプロイ名
  chatApiVersion: string;    // チャット apiVersion

  // 埋め込み (EmbedConfig を構造的に満たす)
  relayBaseUrl: string;        // 埋め込み URL のベース (= corp ベース URL)
  embeddingDeployment: string; // 埋め込みデプロイ名
  apiVersion: string;          // 埋め込み apiVersion
  dimensions: number;

  // claude (開発者モード)
  claudeApiKey: string;
  claudeModel: string;

  // voyage 埋め込み (開発者モード)
  voyageApiKey: string;
  voyageModel: string;

  // tadori 取り込み
  listTitle: string;
  mlAddresses: string[];
  ingestIntervalSec: number;
  embedConcurrency: number;
  ragTopK: number;
  ragMinScore: number;
  /** ハイブリッド検索の重み (0=ベクトルのみ / 1=キーワードのみ)。 */
  ragKeywordWeight: number;
  /** Enter キー単独で送信するか (true=Enter送信/Shift+Enter改行、false=Ctrl/⌘+Enter送信)。 */
  enterSends: boolean;
  /** 再ランカーで検索候補を AI に並べ替えさせる (+1 AI コール、精度向上)。 */
  rerankEnabled: boolean;
  /** 再ランカーへ渡す候補件数 (この件数まで多めに取得し、LLM で並べ替えて上位 ragTopK を採用)。 */
  rerankCandidates: number;
}

/** provider を解決。開発者モード OFF のときは 'claude' を 'corp' に丸める。 */
export function resolveProvider(): Provider {
  const raw = lsGetEff(KEY.provider);
  if (raw === 'claude' && isDeveloperMode()) return 'claude';
  return DEFAULT_PROVIDER;
}

export function loadSettings(): RuntimeSettings {
  const chatModel = getCorpModel();
  const chatEp = resolveCorpChat(chatModel);
  const embeddingModel = getEmbeddingModel();

  return {
    provider: resolveProvider(),

    apiKey: lsGet(KEY.corpKey),
    chatModel,
    corpDeployPrefix: lsGetEff(KEY.corpDeployPrefix),
    corpOverridesRaw: lsGetEff(KEY.corpOverrides),
    embeddingModel,

    chatBaseUrl: chatEp.baseUrl,
    chatDeployment: chatEp.deploymentId,
    chatApiVersion: chatEp.apiVersion,

    relayBaseUrl: corpBaseUrl(),
    embeddingDeployment: deploymentIdFor(embeddingModel),
    apiVersion: lsGet(KEY.embeddingApiVersion) || DEFAULT_EMBEDDING_API_VERSION,
    dimensions: Number(lsGet(KEY.dimensions) || '1024') || 1024,

    claudeApiKey: lsGet(KEY.claudeKey),
    claudeModel: lsGetEff(KEY.claudeModel) || DEFAULT_CLAUDE_MODEL,

    voyageApiKey: lsGet(KEY.voyageKey),
    voyageModel: lsGet(KEY.voyageModel) || DEFAULT_VOYAGE_MODEL,

    listTitle: lsGet(KEY.listTitle) || 'Tadori 受信メール',
    mlAddresses: parseAddressList(lsGet(KEY.mlAddresses)),
    ingestIntervalSec: Number(lsGet(KEY.ingestIntervalSec) || '30') || 30,
    embedConcurrency: Math.min(10, Math.max(1, Number(lsGet(KEY.embedConcurrency) || '3') || 3)),
    ragTopK: Math.min(20, Math.max(1, Number(lsGet(KEY.ragTopK) || '8') || 8)),
    ragMinScore: parseMinScore(lsGet(KEY.ragMinScore)),
    ragKeywordWeight: parseWeight(lsGet(KEY.ragKeywordWeight)),
    enterSends: lsGet(KEY.enterSends) === '1',
    rerankEnabled: lsGet(KEY.rerankEnabled) === '1',
    rerankCandidates: Math.min(30, Math.max(5, Number(lsGet(KEY.rerankCandidates) || '15') || 15)),
  };
}

/** 生値のみ永続化する (導出値は保存しない)。 */
export function saveSettings(s: Partial<RuntimeSettings>): void {
  if (s.provider !== undefined)         lsSet(KEY.provider, s.provider);
  if (s.apiKey !== undefined)           lsSet(KEY.corpKey, s.apiKey.trim());
  if (s.chatModel !== undefined)        lsSet(KEY.corpModel, s.chatModel);
  if (s.relayBaseUrl !== undefined)     lsSet(KEY.corpBaseUrl, s.relayBaseUrl.trim());
  if (s.corpDeployPrefix !== undefined) lsSet(KEY.corpDeployPrefix, s.corpDeployPrefix.trim());
  if (s.corpOverridesRaw !== undefined) lsSet(KEY.corpOverrides, s.corpOverridesRaw.trim());
  if (s.embeddingModel !== undefined)   lsSet(KEY.embeddingModel, s.embeddingModel);
  if (s.apiVersion !== undefined)       lsSet(KEY.embeddingApiVersion, s.apiVersion.trim());
  if (s.dimensions !== undefined)       lsSet(KEY.dimensions, String(s.dimensions));
  if (s.claudeApiKey !== undefined)     lsSet(KEY.claudeKey, s.claudeApiKey.trim());
  if (s.claudeModel !== undefined)      lsSet(KEY.claudeModel, s.claudeModel);
  if (s.voyageApiKey !== undefined)     lsSet(KEY.voyageKey, s.voyageApiKey.trim());
  if (s.voyageModel !== undefined)      lsSet(KEY.voyageModel, s.voyageModel);
  if (s.listTitle !== undefined)        lsSet(KEY.listTitle, s.listTitle);
  if (s.mlAddresses !== undefined)      lsSet(KEY.mlAddresses, s.mlAddresses.join('\n'));
  if (s.ingestIntervalSec !== undefined) lsSet(KEY.ingestIntervalSec, String(s.ingestIntervalSec));
  if (s.embedConcurrency !== undefined)  lsSet(KEY.embedConcurrency, String(Math.min(10, Math.max(1, s.embedConcurrency))));
  if (s.ragTopK !== undefined)           lsSet(KEY.ragTopK, String(Math.min(20, Math.max(1, Math.round(s.ragTopK)))));
  if (s.ragMinScore !== undefined)       lsSet(KEY.ragMinScore, String(Math.min(1, Math.max(0, s.ragMinScore))));
  if (s.ragKeywordWeight !== undefined)  lsSet(KEY.ragKeywordWeight, String(Math.min(1, Math.max(0, s.ragKeywordWeight))));
  if (s.enterSends !== undefined)        lsSet(KEY.enterSends, s.enterSends ? '1' : '');
  if (s.rerankEnabled !== undefined)     lsSet(KEY.rerankEnabled, s.rerankEnabled ? '1' : '');
  if (s.rerankCandidates !== undefined)  lsSet(KEY.rerankCandidates, String(Math.min(30, Math.max(5, Math.round(s.rerankCandidates)))));
}

function parseMinScore(raw: string): number {
  if (raw === '') return 0.3;
  const n = Number(raw);
  return isNaN(n) ? 0.3 : Math.min(1, Math.max(0, n));
}

function parseWeight(raw: string): number {
  if (raw === '') return 0.4;
  const n = Number(raw);
  return isNaN(n) ? 0.4 : Math.min(1, Math.max(0, n));
}

export function parseAddressList(raw: string): string[] {
  return raw.split(/[\n,;]+/).map(s => s.trim()).filter(Boolean);
}
