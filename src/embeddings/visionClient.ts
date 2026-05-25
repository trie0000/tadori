// PPTX マニュアル取り込み用 Vision LLM クライアント。
// スライド 1 枚の PNG + 補助テキスト (Shape 抽出原文 + 表) を GPT-5 vision に
// 投げ、図形配置・矢印・色分けまで踏まえた Markdown 化された記述を受け取る。
//
// チャット応答に使う Azure OpenAI と同じデプロイ (s.chatDeployment) を流用する想定。
// GPT-5 は vision 入力対応 (社内 Azure OpenAI でデプロイ済みである前提)。
//
// 設計参照: docs/pptx-rag-design.md §7

import type { RuntimeSettings } from '../api/aiSettings';
import { recordChat } from '../usage/tracker';
import { chatYen } from '../usage/pricing';

export interface VisionSlideInput {
  /** PNG バイナリ (base64 エンコード済み)。relay から受け取ったまま渡す。 */
  pngBase64: string;
  /** Shape から抽出した正確なテキスト (改行区切り)。空文字なら省略可。 */
  rawText: string;
  /** 検出された表データ。各表 = 2D 配列 (行 → セル配列)。空配列なら省略。 */
  tables: string[][][];
  /** スライド番号 (1-origin)。プロンプトの参照用。 */
  slideNo: number;
  /** スライドタイトル (placeholder から抽出)。空でも可。 */
  title?: string;
  /** スピーカーノート (本文には含めるが Vision プロンプトには付加情報として渡す)。 */
  notes?: string;
}

export interface VisionResult {
  /** Markdown 化されたスライド記述 (RAG embedding 対象)。 */
  markdown: string;
  /** トークン使用量見積 (チャットメーター集計用)。 */
  inputTokens: number;
  outputTokens: number;
}

const SYSTEM_PROMPT = [
  'あなたは技術マニュアルの図解 (PowerPoint スライド画像) を読み取って、',
  '内容を構造化された Markdown に変換するアシスタントです。',
  '',
  '## 入力',
  '- 画像: PowerPoint スライド 1 枚を 1920x1080 でラスタライズした PNG',
  '- 補助情報: スライドの Shape から抽出した正確なテキストと、検出された表',
  '  (画像 OCR の誤読を補正する目的。両者を相互参照して正確な内容を出力する)',
  '',
  '## 出力ルール',
  '1. スライドのタイトル (見出しに見えるもの) は `# <タイトル>` で書く',
  '2. 概要 / セクションは `## <見出し>` で構造化する',
  '3. 図形による「フロー図」「構成図」は、矢印の向き・色分け・位置関係から論理を',
  '   読み取り、番号付き箇条書きや表で再表現する',
  '   例: A→B→C のフローなら "1. A: 〜  2. B: 〜  3. C: 〜"',
  '4. 表 (table) は Markdown 表 (`|...|`) に変換する',
  '5. 色分けに意味がある場合 (必須=赤、任意=青 等) は本文中に「※ 赤枠は必須」のように注記',
  '6. 装飾的な要素 (背景画像、ロゴ、ヘッダフッタ等) は無視',
  '7. このスライドが取り扱う「主題」を最後に 1 行サマリで書く: `> 要点: <50字以内>`',
  '8. 補助情報のテキストと画像中の文字に食い違いがある場合は、補助情報を優先する',
  '   (補助情報は Shape から API で取得した正確な文字列、画像 OCR より信頼できる)',
  '',
  '## 禁止',
  '- 「私は AI です」のような前置き',
  '- ```markdown``` のコードブロック装飾 (生の Markdown のみ)',
  '- 推測の追加 (画像と補助情報に無い情報は書かない)',
  '- 空のスライドや内容が読み取れない場合は `# (空のスライド)` だけ書いて終わる',
].join('\n');

function formatAuxText(input: VisionSlideInput): string {
  const lines: string[] = [];
  lines.push(`スライド番号: ${input.slideNo}`);
  if (input.title) lines.push(`タイトル placeholder: ${input.title}`);

  if (input.rawText) {
    lines.push('');
    lines.push('Shape から抽出したテキスト (順不同):');
    // 各テキストブロックの冒頭に "- " を付けて見やすく
    for (const line of input.rawText.split('\n').map(l => l.trim()).filter(Boolean)) {
      lines.push(`  - ${line}`);
    }
  }

  if (input.tables.length > 0) {
    lines.push('');
    lines.push('検出された表 (構造化済み):');
    for (let i = 0; i < input.tables.length; i++) {
      const tbl = input.tables[i];
      lines.push(`  表 #${i + 1}:`);
      for (const row of tbl) {
        lines.push(`    | ${row.join(' | ')} |`);
      }
    }
  }

  if (input.notes) {
    lines.push('');
    lines.push('スピーカーノート (本文ではないが補足情報):');
    for (const line of input.notes.split('\n').map(l => l.trim()).filter(Boolean)) {
      lines.push(`  - ${line}`);
    }
  }

  return lines.join('\n');
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

/** スライド 1 枚 → Vision LLM → Markdown 記述。
 *  429/5xx は最大 5 回までリトライ。Abort 対応。 */
export async function describeSlide(
  input: VisionSlideInput,
  s: RuntimeSettings,
  signal?: AbortSignal,
): Promise<VisionResult> {
  const auxText = formatAuxText(input);
  const userTextPart = [
    '次のスライド画像を解析し、システム指示に従って Markdown 化してください。',
    '',
    '--- 補助情報 ---',
    auxText,
    '--- ここまで補助情報 ---',
    '',
    '画像と補助情報を統合し、Markdown のみを出力してください。',
  ].join('\n');

  const url = `${s.chatBaseUrl.replace(/\/+$/, '')}`
    + `/openai/deployments/${encodeURIComponent(s.chatDeployment)}`
    + `/chat/completions?api-version=${encodeURIComponent(s.chatApiVersion)}`;
  const headers: Record<string, string> = { 'Content-Type': 'application/json' };
  if (s.apiKey) headers['api-key'] = s.apiKey;

  const body = {
    messages: [
      { role: 'system', content: SYSTEM_PROMPT },
      {
        role: 'user',
        content: [
          { type: 'text', text: userTextPart },
          {
            type: 'image_url',
            image_url: {
              url: `data:image/png;base64,${input.pngBase64}`,
              // 高解像度モード (図解の細部を読み取らせるため)
              detail: 'high',
            },
          },
        ],
      },
    ],
    stream: false,
  };

  const maxRetries = 5;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    if (signal?.aborted) throw new DOMException('aborted', 'AbortError');
    const res = await fetch(url, {
      method: 'POST',
      headers,
      credentials: 'omit',
      signal,
      body: JSON.stringify(body),
    });

    if (res.ok) {
      const json = await res.json() as {
        choices?: { message?: { content?: string } }[];
        usage?: { prompt_tokens?: number; completion_tokens?: number };
      };
      const md = (json.choices?.[0]?.message?.content ?? '').trim();
      const inTok = json.usage?.prompt_tokens ?? Math.ceil((SYSTEM_PROMPT.length + userTextPart.length) / 3) + 1200;
      const outTok = json.usage?.completion_tokens ?? Math.ceil(md.length / 3);
      // 使用量計上 (チャットメーターと同じバケットに記録)
      recordChat(s.chatModel, inTok, outTok);
      return {
        markdown: md,
        inputTokens: inTok,
        outputTokens: outTok,
      };
    }

    // リトライ判定
    const status = res.status;
    const bodyText = await res.text().catch(() => '');
    const retryable = status === 429 || (status >= 500 && status < 600);
    if (!retryable || attempt >= maxRetries) {
      throw new Error(`Vision LLM failed (slide ${input.slideNo}): HTTP ${status} ${bodyText.slice(0, 300)}`);
    }
    const delayMs = inferRetryDelayMs(res, bodyText, attempt);
    await sleepRespectingAbort(delayMs, signal);
  }
  throw new Error(`Vision LLM failed (slide ${input.slideNo}): retries exhausted`);
}

/** 円換算ヘルパ (UI 表示用)。チャットモデルと同じ単価で計上。 */
export function visionYen(s: RuntimeSettings, r: VisionResult): number {
  return chatYen(s.chatModel, r.inputTokens + r.outputTokens);
}
