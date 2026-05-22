// Tadori 実行時設定。環境依存値はここに集約する。
// PoC で確定済みの前提:
//   - SharePoint は Cookie 認証 (credentials:'include') で REST READ 可能
//   - 埋め込みは PowerShell 中継サーバ経由で Azure OpenAI を叩く
//     (ブラウザは社内プロキシを per-request 指定できないため)

export interface TadoriConfig {
  /** 中継サーバの listen URL (scripts/tadori-ai-relay.ps1)。 */
  relayBaseUrl: string;
  /** Azure OpenAI のデプロイ名 (モデル名と一致しない場合あり)。 */
  embeddingDeployment: string;
  /** Azure OpenAI API バージョン。 */
  apiVersion: string;
  /** Matryoshka 短縮の次元数。ADR-004 で 256。 */
  dimensions: number;
  /** ベクトルを格納する SharePoint List の表示名。 */
  listTitle: string;
  /** 取り込みループの間隔 (ms)。ADR では 30 秒。 */
  ingestIntervalMs: number;
  /** try-claim のタイムアウト (ms)。10 分。 */
  claimTimeoutMs: number;
  /** 1 回の埋め込みバッチ件数。 */
  embedBatchSize: number;
}

export const DEFAULT_CONFIG: TadoriConfig = {
  relayBaseUrl: 'http://localhost:18080',
  embeddingDeployment: 'text-embedding-3-small',
  apiVersion: '2024-02-01',
  dimensions: 256,
  listTitle: '受信メールリスト',
  ingestIntervalMs: 30_000,
  claimTimeoutMs: 10 * 60_000,
  embedBatchSize: 16,
};

/** SharePoint List に Tadori が追加する列の内部名。 */
export const COLUMNS = {
  isMl: 'is_ml',
  ragStatus: 'rag_status',
  claimedBy: 'claimed_by',
  claimedAt: 'claimed_at',
  embeddedAt: 'embedded_at',
  embedding: 'embedding',
} as const;

export type RagStatus = 'claimed' | 'embedded' | 'indexed' | 'error';
