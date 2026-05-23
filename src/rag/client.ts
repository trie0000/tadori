// RAG 回答合成クライアント。中継サーバ経由で Azure OpenAI chat/completions を叩く。
// 上位メールを文脈に詰め、出典 [n] 付きの回答をストリーミングで返す。

import type { RuntimeSettings } from '../api/aiSettings';
import { streamClaude } from '../api/aiClaude';
import { recordChat } from '../usage/tracker';
import { chatYen, estimateTokens } from '../usage/pricing';

export interface ChatUsage { model: string; inputTokens: number; outputTokens: number; yen: number; }

const RERANK_SYSTEM = [
  'あなたは社内メーリングリスト検索の関連度判定アシスタントです。',
  'ユーザーの質問に対して、提示された候補メールを「関連が高い順」に並べ替えてください。',
  '出力は最も関連が高い順のインデックス番号 (0 始まり) をカンマ区切りで 1 行のみ。',
  '説明・前置き・括弧・引用符・追加テキストは一切不要です。',
  '例: 3,0,7,2,1',
].join('\n');

export interface RerankCandidate { subject: string; from: string; date: string; body: string; }

function buildRerankPrompt(question: string, candidates: RerankCandidate[]): string {
  const blocks = candidates.map((c, i) => {
    const body = (c.body || '').slice(0, 400).replace(/\s+/g, ' ');
    return `[${i}] 件名: ${c.subject}\n送信者: ${c.from} / ${c.date}\n本文: ${body}`;
  }).join('\n\n');
  return `質問: ${question}\n\n候補:\n\n${blocks}\n\n最も関連が高い順のインデックス (カンマ区切り):`;
}

function parseRerankIndices(text: string, n: number): number[] {
  // 数字シーケンスを順に抽出 → 範囲内 + 重複除去
  const nums = (text.match(/\d+/g) || []).map(Number).filter(x => x >= 0 && x < n);
  const seen = new Set<number>();
  const order: number[] = [];
  for (const x of nums) if (!seen.has(x)) { seen.add(x); order.push(x); }
  // 抜けた候補を末尾に補完 (LLM が全件並べなかった場合)
  for (let i = 0; i < n; i++) if (!seen.has(i)) order.push(i);
  return order;
}

/** 候補メール群を LLM で関連度順に並べ替え、元インデックスの並びを返す。
 *  失敗時は元順序 ([0..n-1])。利用料は onUsage に通知 (チャットと同じ会計)。 */
export async function rerankByLLM(
  question: string,
  candidates: RerankCandidate[],
  s: RuntimeSettings,
  signal?: AbortSignal,
  onUsage?: (u: ChatUsage) => void,
): Promise<number[]> {
  if (candidates.length <= 1) return candidates.map((_, i) => i);
  const userPrompt = buildRerankPrompt(question, candidates);
  const inTok = estimateTokens(RERANK_SYSTEM) + estimateTokens(userPrompt);
  const identity = candidates.map((_, i) => i);
  try {
    let text = '';
    if (s.provider === 'claude') {
      text = await streamClaude({
        apiKey: s.claudeApiKey, model: s.claudeModel, system: RERANK_SYSTEM,
        messages: [{ role: 'user', content: userPrompt }], onText: () => { /* noop */ }, signal,
      });
      recordChat(s.claudeModel, inTok, estimateTokens(text));
      onUsage?.({ model: s.claudeModel, inputTokens: inTok, outputTokens: estimateTokens(text), yen: chatYen(s.claudeModel, inTok + estimateTokens(text)) });
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
            { role: 'system', content: RERANK_SYSTEM },
            { role: 'user', content: userPrompt },
          ],
          stream: false,
        }),
      });
      if (!res.ok) return identity;
      const json = await res.json() as { choices?: { message?: { content?: string } }[] };
      text = json.choices?.[0]?.message?.content ?? '';
      const outTok = estimateTokens(text);
      recordChat(s.chatModel, inTok, outTok);
      onUsage?.({ model: s.chatModel, inputTokens: inTok, outputTokens: outTok, yen: chatYen(s.chatModel, inTok + outTok) });
    }
    const order = parseRerankIndices(text, candidates.length);
    return order.length === candidates.length ? order : identity;
  } catch {
    return identity; // 失敗・中断時は元の順序を維持
  }
}


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
  '与えられた「参照メール」だけを根拠に、日本語で回答してください。',
  '直前までの会話があれば文脈として踏まえ、フォローアップ質問にも答えてください。',
  '',
  '出力の1行目には必ず「TITLE: <この質問を表す15文字以内の短い見出し>」だけを書き、',
  '2行目以降に回答本文を書いてください。見出しに記号・引用符は付けないでください。',
  '',
  '回答本文は Markdown で「階層」を活かして読みやすく整形してください:',
  '- 短い回答 (1〜2文) は箇条書きにせず普通の文で答える。',
  '- 長くなる場合は構造化する:',
  '  - 大きな話題の切り替わりには見出し (##) を使う。',
  '  - 経緯・時系列・人物別の情報は、「日付」「名前」などを親項目 (- ) にして、',
  '    その詳細を 2 スペース字下げの箇条書きでネストする。',
  '    例:',
  '      - **2026-05-07 (平田)**',
  '        - スクリーンショットの内容を確認',
  '        - SSH ポートの差異について照会 [1]',
  '  - 同レベルの箇条書きを 5 件以上ベタ並べにしない。階層・段落で構造化する。',
  '- 重要なキーワード (日付/人名/ID/数値/結論) は **太字** で目立たせる。',
  '- 段落の間は空行を入れる。必要に応じて区切り線 (---) を使う。',
  '- 結論や決定事項は引用ブロック (>) で強調してもよい。',
  '',
  '根拠にしたメールは文末に [1] [2] のように出典番号を付けて示します。',
  '参照メールに答えが無い場合は、推測せず「該当するメールが見つかりませんでした」と正直に答えてください。',
  '',
  '回答の最後の行に必ず「@@SUGGEST@@ 質問1 || 質問2 || 質問3」の形式で、',
  'この内容に関連する短いフォローアップ質問を3つ (各15文字程度) 付けてください。',
].join('\n');

function buildUserPrompt(question: string, sources: RagSource[], onenoteCtx?: string): string {
  const ctx = sources.map(s =>
    `[${s.n}] 件名: ${s.subject}\n送信者: ${s.from} / ${s.date}\n本文:\n${s.body}`,
  ).join('\n\n---\n\n');
  const oneSec = onenoteCtx ? `\n\n---\n\n${onenoteCtx}` : '';
  return `参照メール:\n\n${ctx}${oneSec}\n\n---\n\n質問: ${question}`;
}

const ONENOTE_APPEND_INSTRUCTIONS = [
  '',
  '【OneNote 追記の自動提案】',
  'ユーザの質問文に「OneNote に追記」「ノートに書いて/メモして」「ナレッジに追加」「議事録に書いて」など',
  'OneNote ページへの追記の意図が読み取れる場合のみ、回答の本文と @@SUGGEST@@ の間 (空行で区切る) に',
  '次の 1 行を出力すること:',
  '',
  '@@ONENOTE_APPEND@@ {"pageId":"<id>","heading":"<40字以内>","body":"<MarkdownでOneNoteに追記する本文。改行は \\n (バックスラッシュ + n) で JSON エスケープする>"}',
  '',
  '- pageId は後述の「OneNote 追記候補ページ一覧」から最も適切な 1 件を選ぶ:',
  '  1) ユーザが特定ページを名指ししていればそのページ',
  '  2) なければ「現在開いているページ」',
  '  3) どちらもなければ内容に最も関係する 1 件',
  '- heading は短い見出し (例: "FAQ: 春の懇親会の詳細")。',
  '- body は OneNote にそのまま貼られる Markdown。改行は \\n (バックスラッシュ + n、JSON 文字列の正規エスケープ) で表現する。',
  '- 質問に追記の意図が無い場合は出力しないこと (通常の回答だけ返す)。',
].join('\n');

export interface ChatHistoryMsg { role: 'user' | 'assistant'; content: string; }

function cleanTitle(t: string): string {
  return (t || '')
    .split('\n')[0]
    .replace(/^["'「『（(]+|["'」』）)]+$/g, '')
    .trim()
    .slice(0, 30);
}

const TITLE_MARK = 'TITLE:';
const SUGGEST_MARK = '@@SUGGEST@@';
const APPEND_MARK  = '@@ONENOTE_APPEND@@';
const MAX_MARK_LEN = Math.max(SUGGEST_MARK.length, APPEND_MARK.length);

export interface OneNoteAppendSuggestion { pageId: string; heading: string; body: string; }

/** ストリームを整形するパーサ。
 *  - 先頭の "TITLE: xxx" 行 → onTitle (本文には出さない)
 *  - "@@ONENOTE_APPEND@@ {json}" を 1 行で → onAppendSuggestion (本文には出さない)
 *  - 末尾の "@@SUGGEST@@ a || b || c" → onSuggest (本文には出さない)
 *  本文だけを onDelta へ流す。マーカーがチャンク境界で割れても拾えるよう末尾を保留する。 */
function makeStreamParser(
  onDelta: (t: string) => void,
  onTitle?: (t: string) => void,
  onSuggest?: (qs: string[]) => void,
  onAppendSuggestion?: (s: OneNoteAppendSuggestion) => void,
) {
  let pre = '';
  let bodyStarted = false;
  let titleSent = false;
  let body = '';           // onDelta へ確定送信済み
  let buf = '';            // 本文の全蓄積 (マーカー検出用)
  let mode: 'body' | 'append' | 'suggest' = 'body';
  let appendRaw = '';
  let suggestRaw = '';
  let appendEmitted = false;

  function emitAppend(): void {
    if (appendEmitted || !onAppendSuggestion) { appendRaw = ''; return; }
    appendEmitted = true;
    try {
      const obj = JSON.parse(appendRaw.trim()) as { pageId?: string; heading?: string; body?: string };
      // LLM が誤って "\\n" を出力した場合の保険: 残ったリテラル "\n" (バックスラッシュ+n) を実改行に。
      // JSON.parse が正しい "\n" を実改行へ変換した後にこの置換が走るので、二重変換は起きない。
      const unescapeNl = (s: string): string => s.replace(/\\n/g, '\n').replace(/\\t/g, '\t');
      onAppendSuggestion({
        pageId: String(obj.pageId ?? ''),
        heading: unescapeNl(String(obj.heading ?? '')),
        body: unescapeNl(String(obj.body ?? '')),
      });
    } catch { /* JSON 不正は無視 (回答自体は通常通り表示) */ }
    appendRaw = '';
  }

  function emitSuggest(): void {
    if (!suggestRaw || !onSuggest) return;
    const qs = suggestRaw.split(/\|\||\n/).map(x => x.replace(/^[\s\-・*]+/, '').trim()).filter(Boolean).slice(0, 4);
    if (qs.length) onSuggest(qs);
  }

  function pushBody(text: string): void {
    if (mode === 'suggest') { suggestRaw += text; return; }
    if (mode === 'append') {
      // APPEND は 1 行で完結する想定。\n を見つけたら確定して body モードへ戻る。
      const nl = text.indexOf('\n');
      if (nl < 0) { appendRaw += text; return; }
      appendRaw += text.slice(0, nl);
      emitAppend();
      mode = 'body';
      const tail = text.slice(nl + 1);
      if (tail) pushBody(tail);
      return;
    }
    // body モード
    buf += text;
    const sIdx = buf.indexOf(SUGGEST_MARK);
    const aIdx = buf.indexOf(APPEND_MARK);
    let mIdx = -1;
    let which: 'suggest' | 'append' | null = null;
    if (sIdx >= 0 && aIdx >= 0) { which = sIdx < aIdx ? 'suggest' : 'append'; mIdx = Math.min(sIdx, aIdx); }
    else if (sIdx >= 0) { which = 'suggest'; mIdx = sIdx; }
    else if (aIdx >= 0) { which = 'append'; mIdx = aIdx; }
    if (mIdx >= 0 && which) {
      const head = buf.slice(0, mIdx);
      const toEmit = head.slice(body.length);
      if (toEmit) { body += toEmit; onDelta(toEmit); }
      const markLen = which === 'suggest' ? SUGGEST_MARK.length : APPEND_MARK.length;
      const rest = buf.slice(mIdx + markLen);
      buf = head; // 確定済み本文だけ残す
      if (which === 'suggest') {
        mode = 'suggest';
        suggestRaw += rest;
      } else {
        mode = 'append';
        const nl = rest.indexOf('\n');
        if (nl < 0) { appendRaw += rest; return; }
        appendRaw += rest.slice(0, nl);
        emitAppend();
        mode = 'body';
        const tail = rest.slice(nl + 1);
        if (tail) pushBody(tail);
      }
      return;
    }
    // マーカー未検出 → チャンク境界で割れる可能性があるので末尾を保留
    const safe = Math.max(body.length, buf.length - (MAX_MARK_LEN - 1));
    const toEmit = buf.slice(body.length, safe);
    if (toEmit) { body += toEmit; onDelta(toEmit); }
  }

  return {
    feed(chunk: string): void {
      if (bodyStarted) { pushBody(chunk); return; }
      pre += chunk;
      const consistent = pre.length < TITLE_MARK.length ? TITLE_MARK.startsWith(pre) : pre.startsWith(TITLE_MARK);
      if (!consistent) { bodyStarted = true; pushBody(pre); return; }
      const nl = pre.indexOf('\n');
      if (nl === -1) return;
      const title = cleanTitle(pre.slice(0, nl).replace(/^TITLE:\s*/i, ''));
      if (title && !titleSent) { onTitle?.(title); titleSent = true; }
      bodyStarted = true;
      pushBody(pre.slice(nl + 1).replace(/^\n+/, ''));
    },
    flush(): void {
      if (!bodyStarted) { if (!pre.startsWith(TITLE_MARK)) pushBody(pre); bodyStarted = true; }
      if (mode === 'body') { const tail = buf.slice(body.length); if (tail) { body += tail; onDelta(tail); } }
      if (mode === 'append' && appendRaw) emitAppend();
      emitSuggest();
    },
    get body(): string { return body; },
  };
}

export interface GenerateOptions {
  signal?: AbortSignal;
  onTitle?: (title: string) => void;
  history?: ChatHistoryMsg[];
  onSuggest?: (qs: string[]) => void;
  onUsage?: (u: ChatUsage) => void;
  /** OneNote 追記候補ページ一覧を userPrompt に同梱する文字列。指定時は SYSTEM_PROMPT も拡張される。 */
  onenoteCtx?: string;
  /** ストリームに @@ONENOTE_APPEND@@ が出てきたときのコールバック。 */
  onAppendSuggestion?: (s: OneNoteAppendSuggestion) => void;
}

/** chat/completions をストリーミング呼び出し。onDelta で本文を逐次受け取る。
 *  onTitle で 1 行目のタイトル、onSuggest で末尾のフォローアップ質問を受け取る。
 *  history で直前までの会話 (マルチターン文脈) を渡せる。戻り値は本文。 */
export async function generateAnswer(
  question: string,
  sources: RagSource[],
  s: RuntimeSettings,
  onDelta: (text: string) => void,
  opts: GenerateOptions = {},
): Promise<string> {
  const { signal, onTitle, history, onSuggest, onUsage, onenoteCtx, onAppendSuggestion } = opts;
  const systemPrompt = onenoteCtx ? (SYSTEM_PROMPT + '\n' + ONENOTE_APPEND_INSTRUCTIONS) : SYSTEM_PROMPT;
  const userPrompt = buildUserPrompt(question, sources, onenoteCtx);
  const hist = history ?? [];
  const inputTokens = estimateTokens(systemPrompt) + estimateTokens(userPrompt)
    + hist.reduce((n, m) => n + estimateTokens(m.content), 0);
  const sep = makeStreamParser(onDelta, onTitle, onSuggest, onAppendSuggestion);

  if (s.provider === 'claude') {
    await streamClaude({
      apiKey: s.claudeApiKey,
      model: s.claudeModel,
      system: systemPrompt,
      messages: [...hist, { role: 'user', content: userPrompt }],
      onText: (t) => sep.feed(t),
      signal,
    });
    sep.flush();
    const outTok = estimateTokens(sep.body);
    recordChat(s.claudeModel, inputTokens, outTok);
    onUsage?.({ model: s.claudeModel, inputTokens, outputTokens: outTok, yen: chatYen(s.claudeModel, inputTokens + outTok) });
    return sep.body;
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
        { role: 'system', content: systemPrompt },
        ...hist,
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
      if (data === '[DONE]') {
        sep.flush();
        const out = estimateTokens(sep.body);
        recordChat(s.chatModel, inputTokens, out);
        onUsage?.({ model: s.chatModel, inputTokens, outputTokens: out, yen: chatYen(s.chatModel, inputTokens + out) });
        return sep.body;
      }
      try {
        const json = JSON.parse(data) as { choices?: { delta?: { content?: string } }[] };
        const delta = json.choices?.[0]?.delta?.content;
        if (delta) sep.feed(delta);
      } catch { /* keep-alive 等は無視 */ }
    }
  }
  sep.flush();
  const outTok = estimateTokens(sep.body);
  recordChat(s.chatModel, inputTokens, outTok);
  onUsage?.({ model: s.chatModel, inputTokens, outputTokens: outTok, yen: chatYen(s.chatModel, inputTokens + outTok) });
  return sep.body;
}
