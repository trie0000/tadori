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

/** OneNote ページに追記する 1 ブロック。h=見出し / p=段落 / ul=箇条書き / ol=番号付き / q=引用。 */
export interface AppendBlock { type: 'h' | 'p' | 'ul' | 'ol' | 'q'; text: string; }

/** 既存 OneNote ページの末尾に新規 Outline として追記する (relay 経由)。 */
export async function appendOneNotePage(
  relayBaseUrl: string,
  args: { pageId: string; heading?: string; blocks: AppendBlock[] },
  signal?: AbortSignal,
): Promise<void> {
  if (!relayBaseUrl) throw new Error('中継サーバ URL が未設定です');
  if (!args.pageId) throw new Error('pageId がありません');
  const res = await fetch(`${trim(relayBaseUrl)}/tadori/onenote/append`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(args),
    signal,
  });
  if (!res.ok) {
    const b = await res.text().catch(() => '');
    throw new Error(`OneNote 追記失敗: HTTP ${res.status} ${b.slice(0, 300)}`);
  }
}

/** Markdown 文字列を簡易ブロック列へ。# 見出し / - 箇条書き / 1. 番号 / > 引用 / それ以外は段落。
 *  AI 回答を OneNote に貼るための「最低限の構造保持」変換。完全な markdown ではない。 */
export function markdownToBlocks(md: string): AppendBlock[] {
  const out: AppendBlock[] = [];
  const lines = md.replace(/\r\n/g, '\n').split('\n');
  let para: string[] = [];
  const flushPara = (): void => {
    if (para.length) { out.push({ type: 'p', text: inlineMd(para.join(' ').trim()) }); para = []; }
  };
  for (const raw of lines) {
    const line = raw.trim();
    if (!line) { flushPara(); continue; }
    const h = line.match(/^(#{1,6})\s+(.+)$/);
    if (h) { flushPara(); out.push({ type: 'h', text: inlineMd(h[2]) }); continue; }
    const ul = line.match(/^[-*]\s+(.+)$/);
    if (ul) { flushPara(); out.push({ type: 'ul', text: inlineMd(ul[1]) }); continue; }
    const ol = line.match(/^\d+\.\s+(.+)$/);
    if (ol) { flushPara(); out.push({ type: 'ol', text: inlineMd(ol[1]) }); continue; }
    const q = line.match(/^>\s*(.*)$/);
    if (q) { flushPara(); out.push({ type: 'q', text: inlineMd(q[1]) }); continue; }
    para.push(line);
  }
  flushPara();
  return out;
}

// インライン markdown を OneNote が解釈する HTML タグへ。**太字** / *斜体* / `code` のみ対応。
function inlineMd(s: string): string {
  // XSS 対策で先に &<> をエスケープ → その後マークアップを差し戻す。
  let t = s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  t = t.replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>');
  t = t.replace(/\*([^*]+)\*/g, '<i>$1</i>');
  t = t.replace(/`([^`]+)`/g, '<span style="font-family:Consolas,monospace;background:#f4f4f4">$1</span>');
  return t;
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
