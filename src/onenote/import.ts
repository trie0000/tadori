// OneNote 取り込み: relay 経由で COM から階層/ページを取得し、チャンク化して
// ベクトル DB に投入する。メール取り込み (outlook/import.ts) と同じ作法。

import type { IngestMail } from '../db/writer';
import { splitIntoChunks } from '../lib/chunk';

export interface OneNotePage {
  pageId: string;
  title: string;
  notebook: string;
  section: string;
  lastModified: string;
  body: string;
}
export interface OneNoteSection { id: string; name: string; pages: { id: string; name: string; lastModified?: string }[]; }
export interface OneNoteNotebook { id: string; name: string; sections: OneNoteSection[]; }

function trim(s: string): string { return s.replace(/\/+$/, ''); }

/** relay から OneNote の階層 (ノートブック → セクション → ページ) を取得。 */
export async function fetchOneNoteHierarchy(relayBaseUrl: string, signal?: AbortSignal): Promise<OneNoteNotebook[]> {
  if (!relayBaseUrl) throw new Error('中継サーバ URL が未設定です (AI 接続で設定)');
  const res = await fetch(`${trim(relayBaseUrl)}/tadori/onenote/hierarchy`, { method: 'GET', signal });
  if (!res.ok) {
    const b = await res.text().catch(() => '');
    throw new Error(`OneNote 階層取得失敗: HTTP ${res.status} ${b.slice(0, 300)}`);
  }
  const json = await res.json() as { ok?: boolean; notebooks?: OneNoteNotebook[] };
  return json.notebooks ?? [];
}

/** 指定したページ ID 群の本文を抽出 (ids 省略時は max まで全部)。
 *  ID 多数のときは URL 長と「1 ページ崩れで全滅」を避けるため batchSize ずつに分割して投げる。
 *  onProgress は (done, total) でバッチ確定ごとに呼ばれる。 */
export async function fetchOneNotePages(
  relayBaseUrl: string,
  opts: { ids?: string[]; since?: string; max?: number; batchSize?: number },
  signal?: AbortSignal,
  onProgress?: (done: number, total: number) => void,
): Promise<OneNotePage[]> {
  if (!relayBaseUrl) throw new Error('中継サーバ URL が未設定です (AI 接続で設定)');
  const ids = opts.ids ?? [];
  // ids 未指定 (全件取得) は単発リクエストでそのまま投げる。
  if (ids.length === 0) {
    const p = new URLSearchParams();
    if (opts.since) p.set('since', opts.since);
    if (opts.max) p.set('max', String(opts.max));
    const res = await fetch(`${trim(relayBaseUrl)}/tadori/onenote/pages?${p.toString()}`, { method: 'GET', signal });
    if (!res.ok) { const b = await res.text().catch(() => ''); throw new Error(`OneNote ページ取得失敗: HTTP ${res.status} ${b.slice(0, 300)}`); }
    const j = await res.json() as { pages?: OneNotePage[] };
    return j.pages ?? [];
  }
  const batchSize = Math.max(1, opts.batchSize ?? 20);
  const out: OneNotePage[] = [];
  let firstError: Error | null = null;
  let okBatches = 0, failedBatches = 0;
  for (let off = 0; off < ids.length; off += batchSize) {
    if (signal?.aborted) throw new DOMException('aborted', 'AbortError');
    const slice = ids.slice(off, off + batchSize);
    const p = new URLSearchParams();
    p.set('ids', slice.join(';'));
    if (opts.since) p.set('since', opts.since);
    if (opts.max) p.set('max', String(opts.max));
    try {
      const res = await fetch(`${trim(relayBaseUrl)}/tadori/onenote/pages?${p.toString()}`, { method: 'GET', signal });
      if (!res.ok) {
        const b = await res.text().catch(() => '');
        const err = new Error(`HTTP ${res.status} ${b.slice(0, 300)}`);
        if (!firstError) firstError = err;
        failedBatches++;
      } else {
        const j = await res.json() as { pages?: OneNotePage[] };
        if (j.pages?.length) out.push(...j.pages);
        okBatches++;
      }
    } catch (e) {
      if (signal?.aborted) throw e;
      if (!firstError && e instanceof Error) firstError = e;
      failedBatches++;
    }
    onProgress?.(Math.min(off + batchSize, ids.length), ids.length);
  }
  // 全バッチ失敗かつ 1 件も取れていない → 例外を投げる。一部でも取れたら結果を返す。
  if (out.length === 0 && firstError) {
    throw new Error(`OneNote ページ取得失敗 (全 ${failedBatches} バッチ失敗): ${firstError.message}`);
  }
  return out;
}

/** OneNote ページ群をチャンク化して IngestMail 配列に変換 (既存パイプラインに流す)。 */
export function pagesToIngestMails(pages: OneNotePage[]): IngestMail[] {
  const out: IngestMail[] = [];
  for (const p of pages) {
    const docPath = `onenote://${p.notebook}/${p.section}/${p.title}`;
    const chunks = splitIntoChunks(p.body, { maxChars: 800, overlap: 80 });
    if (chunks.length === 0) continue;
    chunks.forEach((c, i) => {
      out.push({
        messageId: `${p.pageId}#${i}`,
        internetMessageId: '',
        conversationId: p.pageId,   // 親ドキュメント = ページ ID
        kind: 'onenote',
        chunkIdx: i,
        chunkCount: chunks.length,
        docPath,
        subject: c.heading ? `${p.title} - ${c.heading}` : p.title,
        from: `${p.notebook} › ${p.section}`,
        to: [],
        cc: [],
        date: p.lastModified || new Date().toISOString(),
        body: c.text,
        isHtml: false,
      });
    });
  }
  return out;
}

/** OneNote ページに追記する 1 ブロック。
 *  type: h=見出し / p=段落 / ul=箇条書き / ol=番号付き / q=引用 / blank=空行 (段落区切り)
 *  level: ネスト深さ (0 = トップ。Markdown の半角スペース 2 つで +1) */
export interface AppendBlock { type: 'h' | 'p' | 'ul' | 'ol' | 'q' | 'blank'; text: string; level?: number; }

/** 既存 OneNote ページの末尾に新規 Outline として追記する (relay 経由)。
 *  user を渡すと relay 側で「Tadori 追記 by user [日時]」というバナー行が
 *  見出しの直前に挿入される (誰が追記したかをノート単独で識別できるように)。 */
export async function appendOneNotePage(
  relayBaseUrl: string,
  args: { pageId: string; heading?: string; blocks: AppendBlock[]; user?: string },
  signal?: AbortSignal,
): Promise<void> {
  if (!relayBaseUrl) throw new Error('中継サーバ URL が未設定です');
  if (!args.pageId) throw new Error('pageId がありません');
  const res = await fetch(`${trim(relayBaseUrl)}/tadori/onenote/append`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    body: JSON.stringify(args),
    signal,
  });
  if (!res.ok) {
    const b = await res.text().catch(() => '');
    throw new Error(`OneNote 追記失敗: HTTP ${res.status} ${b.slice(0, 300)}`);
  }
}

/** Markdown 文字列を簡易ブロック列へ。# 見出し / - 箇条書き / 1. 番号 / > 引用 / それ以外は段落。
 *  AI 回答を OneNote に貼るための「最低限の構造保持」変換。完全な markdown ではない。
 *  インデント (先頭の半角スペース) を 2 文字 = 1 レベルとして level に変換 (タブは 4 文字相当)。
 *  空行は type=blank として出力 (OneNote 側で段落区切りに使う)。 */
export function markdownToBlocks(md: string): AppendBlock[] {
  const out: AppendBlock[] = [];
  const lines = md.replace(/\r\n/g, '\n').split('\n');
  let para: string[] = [];
  let paraLevel = 0;
  const flushPara = (): void => {
    if (para.length) { out.push({ type: 'p', text: inlineMd(para.join(' ').trim()), level: paraLevel }); para = []; paraLevel = 0; }
  };
  const indentLevel = (raw: string): number => {
    let n = 0;
    for (const ch of raw) {
      if (ch === ' ') n += 1;
      else if (ch === '\t') n += 4;
      else break;
    }
    return Math.floor(n / 2);
  };
  for (const raw of lines) {
    const trimmed = raw.trim();
    if (!trimmed) { flushPara(); out.push({ type: 'blank', text: '', level: 0 }); continue; }
    const level = indentLevel(raw);
    const h = trimmed.match(/^(#{1,6})\s+(.+)$/);
    if (h) { flushPara(); out.push({ type: 'h', text: inlineMd(h[2]), level: 0 }); continue; }
    const ul = trimmed.match(/^[-*]\s+(.+)$/);
    if (ul) { flushPara(); out.push({ type: 'ul', text: inlineMd(ul[1]), level }); continue; }
    const ol = trimmed.match(/^\d+\.\s+(.+)$/);
    if (ol) { flushPara(); out.push({ type: 'ol', text: inlineMd(ol[1]), level }); continue; }
    const q = trimmed.match(/^>\s*(.*)$/);
    if (q) { flushPara(); out.push({ type: 'q', text: inlineMd(q[1]), level: 0 }); continue; }
    if (para.length === 0) paraLevel = level;
    para.push(trimmed);
  }
  flushPara();
  // 末尾の blank は不要なので削る (見栄え用に最後に余白を入れない)
  while (out.length && out[out.length - 1].type === 'blank') out.pop();
  return out;
}

// インライン markdown を OneNote が解釈する HTML タグへ。**太字** / *斜体* / `code` / [label](url) に対応。
function inlineMd(s: string): string {
  // XSS 対策で先に &<> をエスケープ → その後マークアップを差し戻す。
  let t = s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  // [label](url) → <a href="url">label</a> (url 内の "&amp;" は HTML 属性値として正しい)
  t = t.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');
  t = t.replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>');
  t = t.replace(/\*([^*]+)\*/g, '<i>$1</i>');
  t = t.replace(/`([^`]+)`/g, '<span style="font-family:Consolas,monospace;background:#f4f4f4">$1</span>');
  return t;
}

/** 指定したページ ID 群について OneNote の `onenote:` リンクを取得 (relay 経由)。
 *  追記時の「出典」セクションに貼るためのもの。失敗したものは Map に含まれない。 */
export async function fetchOneNoteLinks(
  relayBaseUrl: string, pageIds: string[], signal?: AbortSignal,
): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  if (!relayBaseUrl || pageIds.length === 0) return out;
  try {
    const p = new URLSearchParams(); p.set('ids', pageIds.join(';'));
    const res = await fetch(`${trim(relayBaseUrl)}/tadori/onenote/links?${p.toString()}`, { method: 'GET', signal });
    if (!res.ok) return out;
    const j = await res.json() as { links?: Record<string, string> };
    for (const [k, v] of Object.entries(j.links ?? {})) if (v) out.set(k, v);
  } catch { /* relay 不在等は無視 (リンクなしで本文だけ書き込む) */ }
  return out;
}

/** OneNote のアクティブウィンドウで現在表示中のページ ID を取得。
 *  失敗時 (relay 未起動 / OneNote 未起動 / プロパティ取得不能) は空文字を返す。 */
export async function fetchCurrentOneNotePageId(relayBaseUrl: string, signal?: AbortSignal): Promise<string> {
  if (!relayBaseUrl) return '';
  try {
    const res = await fetch(`${trim(relayBaseUrl)}/tadori/onenote/current`, { method: 'GET', signal });
    if (!res.ok) return '';
    const j = await res.json() as { pageId?: string };
    return j.pageId ?? '';
  } catch { return ''; }
}

/** OneNote 上でページを表示。 */
export async function openOneNotePage(relayBaseUrl: string, pageId: string, signal?: AbortSignal): Promise<void> {
  if (!relayBaseUrl) throw new Error('中継サーバ URL が未設定です');
  if (!pageId) throw new Error('pageId がありません');
  const res = await fetch(`${trim(relayBaseUrl)}/tadori/onenote/open?id=${encodeURIComponent(pageId)}`, { method: 'GET', signal });
  if (!res.ok) {
    const b = await res.text().catch(() => '');
    throw new Error(`OneNote 表示失敗: HTTP ${res.status} ${b.slice(0, 200)}`);
  }
}
