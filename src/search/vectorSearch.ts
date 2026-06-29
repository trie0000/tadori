// 新方式の検索エントリ。SharePoint のセグメントをブラウザ内ベクトルDBへ同期し、
// クエリを埋め込み → 総当たり cosine で Top-K。List 列方式 (旧 searchMails) の置換。

import { getEngine } from '../db/engine';
import { embedQueryFor } from '../embeddings/router';
import type { MailRecord } from '../db/store';
import type { RuntimeSettings } from '../api/aiSettings';
import { getExcludedOneNotePageIds } from '../onenote/exclude';
import { makeInScope, type SourceScope } from './sourceScope';

/** 診断メッセージを relay コンソールに表示させる (ブラウザ Console が読みづらい時用)。
 *  fire-and-forget。relay 未起動なら黙って無視。 */
function relayLog(s: RuntimeSettings, msg: string): void {
  const base = s.relayBaseUrl?.replace(/\/+$/, '');
  if (!base) return;
  try {
    void fetch(`${base}/tadori/log`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ msg }),
    }).catch(() => { /* relay 未起動等は無視 */ });
  } catch { /* noop */ }
}

export interface MailHit {
  messageId: string;
  internetMessageId: string;
  conversationId: string;
  kind: 'mail' | 'onenote' | 'doc' | 'pptx' | 'transcript';
  label?: string;
  chunkIdx?: number;
  chunkCount?: number;
  docPath?: string;
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  date: string;
  body: string;
  isHtml: boolean;
  score: number;
  /** PPTX 取り込みのメタ (kind='pptx' のときのみ意味を持つ)。 */
  pptxFile?: string;
  pptxServerRelUrl?: string;
  slideNo?: number;
  slideTitle?: string;
  thumbServerRelUrl?: string;
  /** Teams 文字起こしのメタ (kind='transcript' のときのみ意味を持つ)。 */
  transcriptFile?: string;
  vttServerRelUrl?: string;
  recordingServerRelUrl?: string;
  startSec?: number;
  /** ドキュメント (kind='doc') のメタ。 */
  docFile?: string;
  docServerRelUrl?: string;
}

function toHit(record: MailRecord, score: number): MailHit {
  return {
    messageId: record.messageId,
    internetMessageId: record.internetMessageId,
    conversationId: record.conversationId,
    kind: record.kind ?? 'mail',
    label: record.label,
    chunkIdx: record.chunkIdx,
    chunkCount: record.chunkCount,
    docPath: record.docPath,
    subject: record.subject,
    from: record.from,
    to: record.to,
    cc: record.cc,
    date: record.date,
    body: record.body,
    isHtml: record.isHtml,
    score,
    pptxFile: record.pptxFile,
    pptxServerRelUrl: record.pptxServerRelUrl,
    slideNo: record.slideNo,
    slideTitle: record.slideTitle,
    thumbServerRelUrl: record.thumbServerRelUrl,
    transcriptFile: record.transcriptFile,
    vttServerRelUrl: record.vttServerRelUrl,
    recordingServerRelUrl: record.recordingServerRelUrl,
    startSec: record.startSec,
    docFile: record.docFile,
    docServerRelUrl: record.docServerRelUrl,
  };
}

/** 手動再同期 (取り込み後など)。 */
export async function resyncVectors(siteUrl: string): Promise<void> {
  const eng = await getEngine(siteUrl);
  await eng.sync.sync();
}

export interface SearchOptions {
  /** ベクトル検索に使うクエリ (LLM ルータが組み立てた意味用クエリ)。省略時は question をそのまま使う。 */
  vectorQuery?: string;
  /** 検索結果のうち、これらの文字列を「すべて含む」レコードに絞る (大文字小文字無視 / 件名+本文)。
   *  空配列なら絞らない。完全一致が 0 件の時は厳密フィルタを外して通常検索にフォールバック。 */
  mustContain?: string[];
  /** 検索対象とする kind を絞る (例: ['mail', 'onenote'] のみ)。
   *  undefined または空配列なら絞らない (全 kind を対象)。
   *  ユーザがチャットボックス近くの「+ ソース選択」UI で設定する。 */
  kinds?: Array<'mail' | 'onenote' | 'doc' | 'pptx' | 'transcript'>;
  /** 種別ごとのサブ項目絞り込み (メール=アドレス / OneNote=ラベル / doc・pptx・会議=フォルダURL)。
   *  各種別とも空/未指定なら絞り込みなし (その種別は全件)。チャットの「＋」ピッカーで設定。 */
  scope?: SourceScope;
  /** 用語辞書によるクエリ展開語 (同義語/略語)。ベクトル用クエリと bigram 用テキスト両方に
   *  畳み込んで、表記違いの取りこぼしを減らす。チャット側で expandQueryTerms から渡す。 */
  glossaryTerms?: string[];
}

export async function searchVectors(
  question: string,
  s: RuntimeSettings,
  siteUrl: string,
  topK: number,
  opts: SearchOptions = {},
): Promise<MailHit[]> {
  const eng = await getEngine(siteUrl);
  if (eng.db.size === 0) return [];
  // 用語辞書の展開語を、ベクトル用クエリと bigram 用テキスト両方に畳み込む。
  const extra = (opts.glossaryTerms ?? []).filter(t => t && t.trim().length >= 2);
  const extraStr = extra.length ? ' ' + extra.join(' ') : '';
  const vecQ = ((opts.vectorQuery || question).trim() || question) + extraStr;
  const kwText = question + extraStr;   // db.search の bigram 側に渡す
  const qvec = await embedQueryFor(vecQ, s);
  const excluded = getExcludedOneNotePageIds();
  const must = (opts.mustContain ?? []).filter(k => k.trim().length >= 2).map(k => k.toLowerCase());
  // kind フィルタ (UI のソース選択チップから渡される)
  const kindFilter = opts.kinds && opts.kinds.length > 0 ? new Set(opts.kinds) : null;

  // 診断ログ: DB の種別内訳 / 次元分布 / kind フィルタ / doc スコープを relay へ。
  const kc = eng.db.kindCounts();
  const dimk = eng.db.dimByKind();
  relayLog(s, `検索開始 q="${vecQ.slice(0, 40)}" queryDim=${qvec.length} ` +
    `manifest_seg数=${eng.sync.lastStats.manifestSealed} DL数=${eng.sync.lastStats.downloaded} ` +
    `DB件数=${eng.db.size} 種別=${JSON.stringify(kc)} ` +
    `kindFilter=${JSON.stringify(opts.kinds ?? '全部')} ` +
    `scope=${opts.scope && Object.keys(opts.scope).length > 0 ? JSON.stringify(opts.scope) : '無し(全部)'} ` +
    `次元分布=${JSON.stringify(dimk)}`);
  console.log('[tadori] search:', { dbSize: eng.db.size, kinds: kc, dimByKind: dimk, kindFilter: opts.kinds, scope: opts.scope });
  // 種別ごとのサブ項目スコープ判定 (各種別とも空/未指定なら全通し)。本番もテストも同一関数。
  const docInScope = makeInScope(opts.scope);

  // レコードが必須キーワードを「全部」含むかチェック (件名+本文を対象、大文字小文字無視)。
  const containsAll = (record: { subject: string; body: string }): boolean => {
    if (must.length === 0) return true;
    const hay = (record.subject + '\n' + record.body).toLowerCase();
    for (const k of must) if (!hay.includes(k)) return false;
    return true;
  };

  // 1 OneNote ページが多数チャンクに分かれている場合、固定 over-pull だとデデュープ後に
  // topK 未満で返るリスクがある (codex review 指摘)。
  // dedup 後の件数が topK に達するまで pull を 2 倍ずつ拡張しつつ再検索する。
  // db.search はメモリ内の全件ソートなので、pull が増えてもコストはほぼ一定 (slice の長さが変わるだけ)。
  // must フィルタが厳しい時は対象が激減するので、pull の初期値も大きめにとる。
  const initialPull = Math.max(topK * 3, topK + excluded.size + 20, must.length > 0 ? topK * 10 : 0);
  let pull = initialPull;
  const dbSize = eng.db.size;
  let deduped: ReturnType<typeof eng.db.search> = [];
  let mustHit = 0; // must を満たすレコードが何件取れたか (デバッグ/フォールバック判定用)

  const runOnce = (applyMust: boolean): ReturnType<typeof eng.db.search> => {
    const raw = eng.db.search(qvec, pull, kwText, s.ragKeywordWeight)
      .filter(({ record }) => !(record.kind === 'onenote' && excluded.has(record.conversationId)))
      .filter(({ record }) => !kindFilter || kindFilter.has(record.kind ?? 'mail'))
      .filter(({ record }) => docInScope(record))
      .filter(({ record }) => !applyMust || containsAll(record));
    const seenPageIds = new Set<string>();
    const out: ReturnType<typeof eng.db.search> = [];
    for (const h of raw) {
      if (h.record.kind === 'onenote' && h.record.conversationId) {
        if (seenPageIds.has(h.record.conversationId)) continue;
        seenPageIds.add(h.record.conversationId);
      }
      out.push(h);
      if (out.length >= topK) break;
    }
    return out;
  };

  for (let attempt = 0; attempt < 5; attempt++) {
    deduped = runOnce(must.length > 0);
    mustHit = deduped.length;
    if (deduped.length >= topK) break;
    if (pull >= dbSize) break;
    pull = Math.min(pull * 2, dbSize);
  }

  // must 指定があったのに 0 件 → キーワード厳密一致を諦めて通常検索にフォールバック (取りこぼし防止)。
  if (must.length > 0 && mustHit === 0) {
    pull = initialPull;
    for (let attempt = 0; attempt < 5; attempt++) {
      deduped = runOnce(false);
      if (deduped.length >= topK) break;
      if (pull >= dbSize) break;
      pull = Math.min(pull * 2, dbSize);
    }
  }

  // ヒット件数を種別別 + スコア帯で relay へ。
  const byKind: Record<string, number> = {};
  for (const { record } of deduped) { const k = record.kind ?? 'mail'; byKind[k] = (byKind[k] ?? 0) + 1; }
  const top = deduped.slice(0, 5).map(d => `${d.record.kind ?? 'mail'}:${d.score.toFixed(3)}`).join(' ');
  relayLog(s, `検索ヒット ${deduped.length}件 種別別=${JSON.stringify(byKind)} 上位5=[${top}]`);

  return deduped.map(({ record, score }) => toHit(record, score));
}

/** 同一スレッド (conversationId) の全メールを時系列で返す (経緯要約用)。 */
export async function getThread(siteUrl: string, conversationId: string): Promise<MailHit[]> {
  const eng = await getEngine(siteUrl);
  // 除外指定された OneNote ページは要約対象からも外す (チャンク全部 = 親ページごと)。
  const excluded = getExcludedOneNotePageIds();
  if (excluded.has(conversationId)) return [];
  return eng.db.byConversation(conversationId).map(r => toHit(r, 1));
}
