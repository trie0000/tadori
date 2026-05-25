// ユーザの質問を LLM が解析し、検索戦略 (ベクトル文 / 必須キーワード) を返す。
// 例: 「申請番号 APP-2026-1234 の承認状況は?」
//   → vectorQuery: "申請の承認状況",  keywords: ["APP-2026-1234"]
// ID やプロジェクト名のような distinctive な文字列は完全一致で必須にする一方、
// 検索の文脈 (どんな意味か) はベクトル検索に任せて両者の良いとこ取りを狙う。

import { streamClaude } from '../api/aiClaude';
import { recordChat } from '../usage/tracker';
import { chatYen } from '../usage/pricing';
import type { ChatHistoryMsg, ChatUsage } from './client';
import type { RuntimeSettings } from '../api/aiSettings';

export interface QueryPlan {
  /** ベクトル検索に使う再構成済みクエリ。元の質問でも可。 */
  vectorQuery: string;
  /** ベクトル検索結果のうち、これらすべてを (大文字小文字無視で) 含むレコードに絞る。
   *  空配列なら絞らない (通常のハイブリッド)。 */
  keywords: string[];
  /** LLM が判定した検索の性格 (UI 表示やログ用)。 */
  mode: 'keyword' | 'semantic' | 'mixed';
}

const ROUTER_SYSTEM = [
  'あなたは社内 RAG 検索のクエリルータです。ユーザの質問を解析し、',
  '次の JSON を 1 行で返してください (それ以外の出力は禁止):',
  '',
  '{"mode":"keyword|semantic|mixed","vectorQuery":"<意味検索用のクエリ>","keywords":["<必須完全一致>", ...]}',
  '',
  'ルール:',
  '- keywords には「申請番号 / チケット ID / プロジェクトコード / 製品名 / 固有名詞 / 日付指定」等の',
  '  必ず含まれるべき文字列だけを入れる (3 文字以上、最大 4 個まで)。',
  '- 数字単体 (例: "2026" "100") やよくある単語 (例: "メール" "件" "について") は keywords に入れない。',
  '- vectorQuery には質問の「意味的な主題」を 1 文で表す。元の文がそのまま使えるならそれでよい。',
  '  IDや固有名詞は keywords 側に出すので vectorQuery には含めなくてもよい。',
  '- 純粋に ID/コード/固有名詞だけで探す質問 → mode="keyword"。意味で探す質問 → "semantic"。両方混在 → "mixed"。',
  '',
  '★ フォローアップ質問 (直前会話を踏まえた省略表現) の解決 ★',
  '- 「直前の会話」が与えられた場合、ユーザの質問に含まれる指示語 (それ / あれ / この / 上記 等) や、',
  '  「要約して」「もっと詳しく」「続きは?」のような前提が省略された質問は、',
  '  直前会話から主題を補って vectorQuery を組み立てること。',
  '  例: 直前 user="APP-2026-1234 の承認状況は?" / 今回 user="要約して"',
  '      → vectorQuery="APP-2026-1234 の承認状況の要約", keywords=["APP-2026-1234"]',
  '- 直前会話と無関係な新規質問の場合は、履歴を無視してその質問だけを解析する。',
  '',
  '- 出力は厳密に有効な JSON。前後に説明文や ```等の装飾は付けない。',
].join('\n');

const FALLBACK = (q: string): QueryPlan => ({ vectorQuery: q, keywords: [], mode: 'semantic' });

/** JSON を厳格パースし、形が壊れていたら null。 */
function parsePlan(text: string): QueryPlan | null {
  // LLM が ```json ... ``` で囲ってきても拾えるよう緩く抽出。
  const m = text.match(/\{[\s\S]*\}/);
  if (!m) return null;
  try {
    const obj = JSON.parse(m[0]) as Partial<QueryPlan>;
    const vectorQuery = typeof obj.vectorQuery === 'string' ? obj.vectorQuery.trim() : '';
    const keywords = Array.isArray(obj.keywords)
      ? obj.keywords.filter((k): k is string => typeof k === 'string' && k.trim().length >= 2).map(k => k.trim()).slice(0, 4)
      : [];
    const mode = obj.mode === 'keyword' || obj.mode === 'mixed' || obj.mode === 'semantic' ? obj.mode : (keywords.length > 0 ? 'mixed' : 'semantic');
    if (!vectorQuery && keywords.length === 0) return null;
    return { vectorQuery: vectorQuery || keywords.join(' '), keywords, mode };
  } catch { return null; }
}

/** 直近会話を ROUTER 用に短く整形する。アシスタント回答は要点だけ拾えればよいので短くトリム。 */
function formatHistoryForRouter(history?: ChatHistoryMsg[]): string {
  if (!history || history.length === 0) return '';
  // 直近 4 メッセージ (= 2 ターン分) で十分。長すぎるとコスト増+ノイズ。
  const recent = history.slice(-4);
  const lines = recent.map(m => {
    const tag = m.role === 'user' ? 'ユーザ' : 'アシスタント';
    // アシスタントの長文回答は冒頭 300 字に圧縮 (主題の特定には十分)
    const max = m.role === 'assistant' ? 300 : 500;
    const c = m.content.length > max ? m.content.slice(0, max) + '…' : m.content;
    return `${tag}: ${c}`;
  });
  return lines.join('\n');
}

/** 質問を LLM に投げて検索プランを得る。失敗時は安全なフォールバック (元クエリ + キーワード無し)。
 *  チャット応答生成 (generateAnswer) と同じプロバイダ・モデルを使う。
 *  history を渡すと、フォローアップ質問 (「要約して」「もっと詳しく」「それは誰?」等) の
 *  指示語・省略を解決した vectorQuery を組み立てるようルータに指示する。 */
export async function classifyQuery(
  question: string,
  s: RuntimeSettings,
  signal?: AbortSignal,
  onUsage?: (u: ChatUsage) => void,
  history?: ChatHistoryMsg[],
): Promise<QueryPlan> {
  const q = question.trim();
  if (!q) return FALLBACK(q);

  const histBlock = formatHistoryForRouter(history);
  const userPrompt = histBlock
    ? `直前の会話 (古い順):\n${histBlock}\n\n---\n\n今回の質問:\n${q}`
    : `質問:\n${q}`;
  try {
    let text = '';
    const estIn = Math.ceil((ROUTER_SYSTEM.length + userPrompt.length) / 3);

    if (s.provider === 'claude') {
      text = await streamClaude({
        apiKey: s.claudeApiKey, model: s.claudeModel, system: ROUTER_SYSTEM,
        messages: [{ role: 'user', content: userPrompt }], onText: () => { /* noop */ }, signal,
      });
      const outTok = Math.ceil(text.length / 3);
      recordChat(s.claudeModel, estIn, outTok);
      onUsage?.({ model: s.claudeModel, inputTokens: estIn, outputTokens: outTok, yen: chatYen(s.claudeModel, estIn + outTok) });
    } else {
      const url = `${s.chatBaseUrl.replace(/\/+$/, '')}`
        + `/openai/deployments/${encodeURIComponent(s.chatDeployment)}`
        + `/chat/completions?api-version=${encodeURIComponent(s.chatApiVersion)}`;
      const headers: Record<string, string> = { 'Content-Type': 'application/json' };
      if (s.apiKey) headers['api-key'] = s.apiKey;
      const res = await fetch(url, {
        method: 'POST', headers, credentials: 'omit', signal,
        body: JSON.stringify({
          messages: [
            { role: 'system', content: ROUTER_SYSTEM },
            { role: 'user', content: userPrompt },
          ],
          stream: false,
        }),
      });
      if (!res.ok) return FALLBACK(q);
      const json = await res.json() as { choices?: { message?: { content?: string } }[] };
      text = json.choices?.[0]?.message?.content ?? '';
      const outTok = Math.ceil(text.length / 3);
      recordChat(s.chatModel, estIn, outTok);
      onUsage?.({ model: s.chatModel, inputTokens: estIn, outputTokens: outTok, yen: chatYen(s.chatModel, estIn + outTok) });
    }
    const plan = parsePlan(text);
    return plan ?? FALLBACK(q);
  } catch {
    return FALLBACK(q);
  }
}
