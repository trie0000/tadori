// RAG 回答合成クライアント。中継サーバ経由で Azure OpenAI chat/completions を叩く。
// 上位メールを文脈に詰め、出典 [n] 付きの回答をストリーミングで返す。

import type { RuntimeSettings } from '../api/aiSettings';
import { streamClaude } from '../api/aiClaude';
import { recordChat } from '../usage/tracker';
import { estimateTokens } from '../usage/pricing';

export interface RagSource {
  /** 1 始まりの出典番号 (回答中の [n] と対応)。 */
  n: number;
  subject: string;
  from: string;
  date: string;
  body: string;
}

const SYSTEM_PROMPT = [
  'あなたは社内メーリングリストの過去ログに基づいて回答するアシスタントです。',
  '与えられた「参照メール」だけを根拠に、日本語で簡潔に回答してください。',
  '根拠にしたメールは文末に [1] [2] のように出典番号を付けて示します。',
  '参照メールに答えが無い場合は、推測せず「該当するメールが見つかりませんでした」と正直に答えてください。',
].join('\n');

function buildUserPrompt(question: string, sources: RagSource[]): string {
  const ctx = sources.map(s =>
    `[${s.n}] 件名: ${s.subject}\n送信者: ${s.from} / ${s.date}\n本文:\n${s.body}`,
  ).join('\n\n---\n\n');
  return `参照メール:\n\n${ctx}\n\n---\n\n質問: ${question}`;
}

const TITLE_SYSTEM = [
  '次のユーザーの質問を、日本語で15文字以内の短い見出しに要約してください。',
  '見出しの語句だけを返し、句読点・記号・引用符・接頭辞は付けないでください。',
].join('\n');

function cleanTitle(t: string): string {
  return (t || '')
    .split('\n')[0]
    .replace(/^["'「『（(]+|["'」』）)]+$/g, '')
    .trim()
    .slice(0, 30);
}

/** 質問を短い見出しに要約 (チャット履歴のタイトル用)。失敗時は空文字。 */
export async function summarizeTitle(question: string, s: RuntimeSettings, signal?: AbortSignal): Promise<string> {
  const inTok = estimateTokens(TITLE_SYSTEM) + estimateTokens(question);
  try {
    if (s.provider === 'claude') {
      const t = await streamClaude({
        apiKey: s.claudeApiKey, model: s.claudeModel, system: TITLE_SYSTEM,
        messages: [{ role: 'user', content: question }], onText: () => { /* noop */ }, signal,
      });
      recordChat(s.claudeModel, inTok, estimateTokens(t));
      return cleanTitle(t);
    }

    const url = `${s.chatBaseUrl.replace(/\/+$/, '')}`
      + `/openai/deployments/${encodeURIComponent(s.chatDeployment)}`
      + `/chat/completions?api-version=${encodeURIComponent(s.chatApiVersion)}`;
    const headers: Record<string, string> = { 'Content-Type': 'application/json' };
    if (s.apiKey) headers['api-key'] = s.apiKey;
    const res = await fetch(url, {
      method: 'POST', headers, credentials: 'omit', signal,
      body: JSON.stringify({
        messages: [
          { role: 'system', content: TITLE_SYSTEM },
          { role: 'user', content: question },
        ],
        stream: false,
      }),
    });
    if (!res.ok) return '';
    const json = await res.json() as { choices?: { message?: { content?: string } }[] };
    const content = json.choices?.[0]?.message?.content ?? '';
    recordChat(s.chatModel, inTok, estimateTokens(content));
    return cleanTitle(content);
  } catch {
    return '';
  }
}

/** chat/completions をストリーミング呼び出し。onDelta で逐次トークンを受け取る。 */
export async function generateAnswer(
  question: string,
  sources: RagSource[],
  s: RuntimeSettings,
  onDelta: (text: string) => void,
  signal?: AbortSignal,
): Promise<string> {
  const userPrompt = buildUserPrompt(question, sources);
  const inputTokens = estimateTokens(SYSTEM_PROMPT) + estimateTokens(userPrompt);

  if (s.provider === 'claude') {
    const answer = await streamClaude({
      apiKey: s.claudeApiKey,
      model: s.claudeModel,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userPrompt }],
      onText: onDelta,
      signal,
    });
    recordChat(s.claudeModel, inputTokens, estimateTokens(answer));
    return answer;
  }

  const url = `${s.chatBaseUrl.replace(/\/+$/, '')}`
    + `/openai/deployments/${encodeURIComponent(s.chatDeployment)}`
    + `/chat/completions?api-version=${encodeURIComponent(s.chatApiVersion)}`;

  const headers: Record<string, string> = { 'Content-Type': 'application/json' };
  if (s.apiKey) headers['api-key'] = s.apiKey;

  const res = await fetch(url, {
    method: 'POST',
    headers,
    credentials: 'omit',
    signal,
    // temperature は送らない: モデルによっては既定(1)以外を拒否する
    // ("Unsupported value: 'temperature' ... Only the default value is supported")。
    body: JSON.stringify({
      messages: [
        { role: 'system', content: SYSTEM_PROMPT },
        { role: 'user', content: userPrompt },
      ],
      stream: true,
    }),
  });

  if (!res.ok || !res.body) {
    const body = await res.text().catch(() => '');
    throw new Error(`chat failed: HTTP ${res.status} ${body.slice(0, 300)}`);
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  let answer = '';

  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split('\n');
    buffer = lines.pop() ?? '';
    for (const line of lines) {
      const t = line.trim();
      if (!t.startsWith('data:')) continue;
      const data = t.slice(5).trim();
      if (data === '[DONE]') { recordChat(s.chatModel, inputTokens, estimateTokens(answer)); return answer; }
      try {
        const json = JSON.parse(data) as { choices?: { delta?: { content?: string } }[] };
        const delta = json.choices?.[0]?.delta?.content;
        if (delta) { answer += delta; onDelta(delta); }
      } catch { /* keep-alive 等は無視 */ }
    }
  }
  recordChat(s.chatModel, inputTokens, estimateTokens(answer));
  return answer;
}
