// RAG チャットパネル。左ペイン (セッション一覧) + スレッド表示 + コンポーザ。
// 送信 → searchVectors → generateAnswer (streaming) → 出典カード。
// セッションは localStorage に保存し、左ペインから再表示できる (src/ui/sessions.ts)。

import { el } from '../lib/dom';
import { icons } from './icons';
import { toast } from './toast';
import { searchVectors, getThread } from '../search/vectorSearch';
import { loadRules, matchesAnyRule } from '../search/exclusionRules';
import { htmlToText, renderMailBody, splitHtmlReplyHistory } from '../lib/mailhtml';
import { cleanBody, splitReplyHistory } from '../lib/mailtext';
import { generateAnswer, rerankByLLM, type RagSource, type ChatHistoryMsg, type OneNoteAppendSuggestion } from '../rag/client';
import { classifyQuery } from '../rag/queryRouter';
import { loadSettings, saveSettings, CORP_AI_MODELS, CLAUDE_MODELS } from '../api/aiSettings';
import { isDeveloperMode } from '../utils/devMode';
import { renderMarkdown } from '../lib/markdown';
import { openMailInOutlook } from '../outlook/import';
import { openOneNotePage, appendOneNotePage, markdownToBlocks, fetchCurrentOneNotePageId, fetchOneNoteLinks, fetchOneNoteHierarchy, fetchTadoriOutlines, replaceTadoriOutline, type TadoriOutline } from '../onenote/import';
import { currentUser } from '../usage/tracker';
import { getEngine } from '../db/engine';
import { getExcludedOneNotePageIds } from '../onenote/exclude';
import { openModal } from './modal';
import { confirmModal } from './modal';
import {
  listSessions, getSession, appendTurn, setTitle, deleteSession, newSessionId,
  type ChatSession, type SavedHit,
} from './sessions';

/** クエリに含まれる 2 文字以上の連続部分をスニペット内で <mark> 強調 (DOM 構築で安全)。
 *  日本語は形態素を持たないので、クエリの部分文字列一致で貪欲にハイライトする。 */
function highlightInto(host: HTMLElement, text: string, query: string): void {
  const q = (query || '').toLowerCase().replace(/\s+/g, '');
  if (!q || q.length < 2) { host.textContent = text; return; }
  const lower = text.toLowerCase();
  let i = 0;
  while (i < text.length) {
    // text[i..] から始まる最長の、query の部分文字列になっている連続を探す
    let len = 0;
    for (let l = Math.min(text.length - i, 12); l >= 2; l--) {
      if (q.includes(lower.slice(i, i + l))) { len = l; break; }
    }
    if (len >= 2) {
      host.appendChild(el('mark', {}, [text.slice(i, i + len)]));
      i += len;
    } else {
      // 連続する非ヒット文字をまとめてテキストノードに
      let j = i + 1;
      while (j < text.length) {
        let hit = false;
        for (let l = Math.min(text.length - j, 12); l >= 2; l--) {
          if (q.includes(lower.slice(j, j + l))) { hit = true; break; }
        }
        if (hit) break;
        j++;
      }
      host.appendChild(document.createTextNode(text.slice(i, j)));
      i = j;
    }
  }
}

/** 「処理中」を示すアニメーション付きインジケータ。 */
function thinkingEl(text: string): HTMLElement {
  return el('span', { class: 'tdr-thinking' }, [
    text,
    el('span', { class: 'tdr-dot' }),
    el('span', { class: 'tdr-dot' }),
    el('span', { class: 'tdr-dot' }),
  ]);
}

const SIDEBAR_W_KEY = 'tadori:sidebar-w';
const SIDEBAR_MIN = 180;
const SIDEBAR_MAX = 520;

interface TurnRefs {
  turnEl: HTMLElement;
  answerText: HTMLElement;
  metaEl: HTMLElement;
  aBody: HTMLElement;
}

export function createChatPanel(root: HTMLElement, siteUrl: string): HTMLElement {
  const thread = el('div', { class: 'tdr-thread' });
  const input  = el('textarea', { class: 'tdr-note-input', placeholder: 'メーリングリストについて質問…', rows: '1' });
  const sendBtn = el('button', { class: 'tdr-note-submit', 'aria-label': '送信', html: icons.send(14) });
  const hintEl = el('div', { class: 'tdr-note-hint' }, ['']);
  const refreshHint = (): void => {
    hintEl.textContent = loadSettings().enterSends
      ? 'Enter で送信 / Shift+Enter で改行'
      : '⌘+Enter または Ctrl+Enter で送信';
  };
  refreshHint();
  input.addEventListener('focus', refreshHint);

  const emptyState = el('div', { class: 'tdr-empty' }, [
    el('div', { class: 'big' }, ['辿り']),
    el('p', {}, ['社内メーリングリストを自然言語で検索できます。']),
    el('p', { style: 'font-size:var(--fs-sm);color:var(--ink-4);margin-top:var(--s-3)' }, [
      '例: 「先月の懇親会の日程はいつですか?」',
    ]),
  ]);

  let abort: AbortController | null = null;
  let generating = false;
  let currentId = newSessionId();   // 現在のセッション (最初の送信まで未保存)
  let hasTurns = false;

  // 生成中は送信ボタンを停止ボタンに切り替える。
  function setGenerating(on: boolean): void {
    generating = on;
    sendBtn.classList.toggle('is-stop', on);
    sendBtn.innerHTML = on ? icons.stop(14) : icons.send(14);
    sendBtn.setAttribute('aria-label', on ? '停止' : '送信');
  }

  // ── 左ペイン (セッション一覧) ──
  const sessionList = el('div', { class: 'tdr-session-list' });
  const newBtn = el('button', { class: 'tdr-new-session' }, [
    el('span', { html: icons.plus(15) }), 'New session',
  ]);
  newBtn.addEventListener('click', startNewSession);
  const searchBtn = el('button', { class: 'tdr-search-btn', 'aria-label': 'チャット履歴を検索', title: 'チャット履歴を検索', html: icons.search(16) });
  searchBtn.addEventListener('click', () => openHistorySearch());

  function refreshList(): void {
    sessionList.replaceChildren();
    const sessions = listSessions();
    if (sessions.length === 0) {
      sessionList.appendChild(el('div', { class: 'tdr-session-empty' }, ['履歴はまだありません']));
      return;
    }
    for (const s of sessions) {
      const titleEl = el('span', { class: 'tdr-session-title', title: s.title }, [s.title]);
      const item = el('div', { class: 'tdr-session' + (s.id === currentId ? ' is-active' : '') }, [
        el('span', { class: 'tdr-session-ic', html: icons.message(14) }),
        titleEl,
      ]);
      item.addEventListener('click', () => openSession(s.id));

      const renameBtn = el('button', { class: 'tdr-session-edit', 'aria-label': '名前を変更', title: '名前を変更', html: icons.edit(13) });
      renameBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        beginRename(s.id, titleEl, s.title);
      });
      item.appendChild(renameBtn);

      const del = el('button', { class: 'tdr-session-del', 'aria-label': '削除', title: '削除', html: icons.trash(13) });
      del.addEventListener('click', (e) => {
        e.stopPropagation();
        confirmModal({
          root,
          title: 'セッションを削除',
          message: `このチャット履歴を削除しますか?\n「${s.title}」`,
          primaryLabel: '削除',
          primaryVariant: 'danger',
          onConfirm: () => {
            deleteSession(s.id);
            if (s.id === currentId) startNewSession();
            else refreshList();
          },
        });
      });
      item.appendChild(del);
      sessionList.appendChild(item);
    }
  }

  /** タイトル要素を入力欄に差し替えてインライン編集。Enter で保存、Esc / blur で確定/取消。 */
  function beginRename(id: string, titleEl: HTMLElement, current: string): void {
    const inp = el('input', { class: 'tdr-session-edit-input', value: current }) as HTMLInputElement;
    titleEl.replaceWith(inp);
    inp.focus(); inp.select();
    let done = false;
    const commit = (save: boolean): void => {
      if (done) return; done = true;
      const v = inp.value.trim();
      if (save && v && v !== current) setTitle(id, v);
      refreshList(); // 元の DOM に戻す
    };
    inp.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); commit(true); }
      else if (e.key === 'Escape') { e.preventDefault(); commit(false); }
    });
    inp.addEventListener('blur', () => commit(true));
    inp.addEventListener('click', e => e.stopPropagation());
  }

  function startNewSession(): void {
    abort?.abort();
    currentId = newSessionId();
    hasTurns = false;
    thread.replaceChildren(emptyState);
    refreshList();
    input.focus();
  }

  function openSession(id: string): void {
    const s = getSession(id);
    if (!s) return;
    abort?.abort();
    currentId = id;
    renderSession(s);
    refreshList();
  }

  /** チャット履歴の横断検索 (Spira のチケット検索 UI を踏襲)。 */
  function openHistorySearch(): void {
    const existing = root.querySelector<HTMLElement>('.tdr-search-backdrop');
    if (existing) { existing.querySelector<HTMLInputElement>('.tdr-search-input')?.focus(); return; }

    const inp = el('input', {
      type: 'text', class: 'tdr-search-input', placeholder: 'タイトル・質問・回答で検索…',
      autocomplete: 'off', spellcheck: 'false',
    }) as HTMLInputElement;
    const summary = el('div', { class: 'tdr-search-summary' });
    const closeBtn = el('button', { class: 'tdr-iconbtn tdr-search-close', 'aria-label': '閉じる', html: icons.close(16) });
    const head = el('div', { class: 'tdr-search-head' }, [
      el('span', { class: 'tdr-search-icon', html: icons.search(18) }),
      inp, closeBtn,
    ]);
    const body = el('div', { class: 'tdr-search-body' });
    const foot = el('div', { class: 'tdr-search-foot' }, [summary]);
    const modal = el('div', { class: 'tdr-modal tdr-search-modal', role: 'dialog', 'aria-modal': 'true' }, [head, body, foot]);
    const backdrop = el('div', { class: 'tdr-backdrop tdr-search-backdrop' }, [modal]);

    const close = (): void => { backdrop.remove(); document.removeEventListener('keydown', onKey); };
    const onKey = (e: KeyboardEvent): void => { if (e.key === 'Escape') close(); };
    backdrop.addEventListener('mousedown', e => { if (e.target === backdrop) close(); });
    closeBtn.addEventListener('click', close);
    document.addEventListener('keydown', onKey);
    root.appendChild(backdrop);
    setTimeout(() => { inp.focus(); }, 0);

    const hint = (title: string, desc: string): HTMLElement => el('div', { class: 'tdr-search-hint' }, [
      el('div', { class: 'tdr-search-hint-title' }, [title]),
      el('div', { class: 'tdr-search-hint-desc' }, [desc]),
    ]);

    let timer: ReturnType<typeof setTimeout> | null = null;
    let last = '';
    const run = (q: string): void => {
      const trimmed = q.trim();
      if (trimmed === last) return;
      last = trimmed;
      body.replaceChildren();
      if (!trimmed) {
        summary.textContent = '';
        body.appendChild(hint('検索ワードを入力してください', 'チャット履歴のタイトル・質問・回答から探します。'));
        return;
      }
      const results = searchSessions(trimmed);
      summary.textContent = results.length === 0 ? '0 件' : `${results.length} セッション`;
      if (results.length === 0) { body.appendChild(hint('結果なし', '別のキーワードで試してください。')); return; }
      const list = el('div', { class: 'tdr-search-results' });
      for (const r of results) {
        const card = el('div', { class: 'tdr-search-result' });
        const titleEl = el('div', { class: 'tdr-search-result-title' });
        highlightInto(titleEl, r.session.title, trimmed);
        card.appendChild(titleEl);
        if (r.snippet) {
          const snip = el('div', { class: 'tdr-search-result-snippet' });
          highlightInto(snip, r.snippet, trimmed);
          card.appendChild(snip);
        }
        card.appendChild(el('div', { class: 'tdr-search-result-meta' }, [`${r.matches} 件ヒット ・ ${r.session.turns.length} ターン`]));
        card.addEventListener('click', () => { close(); openSession(r.session.id); });
        list.appendChild(card);
      }
      body.appendChild(list);
    };
    inp.addEventListener('input', () => { if (timer) clearTimeout(timer); timer = setTimeout(() => run(inp.value), 150); });
    inp.addEventListener('keydown', e => { if (e.key === 'Enter') { if (timer) { clearTimeout(timer); timer = null; } run(inp.value); } });
    body.appendChild(hint('検索ワードを入力してください', 'チャット履歴のタイトル・質問・回答から探します。'));
  }

  function searchSessions(q: string): Array<{ session: ChatSession; matches: number; snippet: string }> {
    const ql = q.toLowerCase();
    const out: Array<{ session: ChatSession; matches: number; snippet: string }> = [];
    for (const s of listSessions()) {
      let matches = 0;
      let snippet = '';
      if (s.title.toLowerCase().includes(ql)) { matches++; if (!snippet) snippet = s.title; }
      for (const t of s.turns) {
        if (t.q.toLowerCase().includes(ql)) { matches++; if (!snippet) snippet = `Q: ${t.q}`; }
        const al = t.answer.toLowerCase();
        if (al.includes(ql)) {
          matches++;
          if (!snippet) {
            const idx = al.indexOf(ql);
            const start = Math.max(0, idx - 30); const end = Math.min(t.answer.length, idx + ql.length + 60);
            snippet = (start > 0 ? '…' : '') + t.answer.slice(start, end).replace(/\n+/g, ' ') + (end < t.answer.length ? '…' : '');
          }
        }
      }
      if (matches > 0) out.push({ session: s, matches, snippet });
    }
    out.sort((a, b) => b.matches - a.matches || (a.session.updatedAt < b.session.updatedAt ? 1 : -1));
    return out;
  }

  // ── スレッド描画 ──
  function buildTurn(q: string): TurnRefs {
    const turnEl = el('div', { class: 'tdr-turn' });
    turnEl.appendChild(el('div', { class: 'tdr-q' }, [q]));
    const answerText = el('div', { class: 'tdr-answer' });
    const metaEl    = el('div', { class: 'tdr-a-meta' });
    const aBody     = el('div', { class: 'tdr-a-body' }, [metaEl, answerText]);
    turnEl.appendChild(el('div', { class: 'tdr-a' }, [
      el('div', { class: 'tdr-a-avatar' }, ['T']),
      aBody,
    ]));
    thread.appendChild(turnEl);
    return { turnEl, answerText, metaEl, aBody };
  }

  function finalizeTurn(refs: TurnRefs, fullMarkdown: string, hits: SavedHit[], ms: number, relayBaseUrl: string, query = '', yen?: number, createdAt?: string): void {
    refs.answerText.innerHTML = renderMarkdown(fullMarkdown).replace(
      /\[(\d+)\]/g,
      (_, n) => `<span class="cite" data-n="${n}">[${n}]</span>`,
    );
    const metaChildren: HTMLElement[] = [];
    if (createdAt) metaChildren.push(el('span', { class: 'tdr-turn-time', title: createdAt }, [fmtTurnTime(createdAt)]));
    metaChildren.push(el('span', {}, [`${hits.length} 件参照`]));
    metaChildren.push(el('span', { class: 'mono' }, [`${ms} ms`]));
    refs.metaEl.replaceChildren(...metaChildren);
    // コピーボタン / OneNote 追記 / 利用料を 1 行で。
    const actions = el('div', { class: 'tdr-turn-actions' }, [
      makeCopyBtn(fullMarkdown),
      makeAppendOneNoteBtn(query, fullMarkdown, relayBaseUrl, hits),
    ]);
    if (yen != null) {
      actions.appendChild(el('span', { class: 'tdr-turn-cost', title: 'このやり取りの AI 利用料 (目安)' }, [fmtYen(yen)]));
    }
    refs.aBody.appendChild(actions);
    // 回答中に [n] として実際に引用された番号を抽出。
    const cited = new Set<number>();
    for (const m of fullMarkdown.matchAll(/\[(\d+)\]/g)) cited.add(Number(m[1]));
    if (hits.length) appendSources(refs.aBody, hits, relayBaseUrl, query, cited);
    wireCiteJump(refs.aBody);
  }

  function fmtYen(n: number): string {
    const v = n < 0.1 ? n.toFixed(3) : n.toFixed(2);
    return '¥' + v;
  }

  /** タイムスタンプを表示用に整形 (今日なら HH:mm、別日なら M/D HH:mm)。 */
  function fmtTurnTime(iso: string): string {
    const d = new Date(iso); if (isNaN(d.getTime())) return '';
    const hm = `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
    const today = new Date();
    const sameDay = d.getFullYear() === today.getFullYear() && d.getMonth() === today.getMonth() && d.getDate() === today.getDate();
    return sameDay ? hm : `${d.getMonth() + 1}/${d.getDate()} ${hm}`;
  }

  // 回答中の [n] 引用クリックで該当の出典カードへスクロール + 展開 + ハイライト。
  // カードは [data-n] で引けるので、グループ分け (引用 / 候補) されていても OK。
  function wireCiteJump(aBody: HTMLElement): void {
    const cites = aBody.querySelectorAll<HTMLElement>('.cite');
    if (!cites.length) return;
    cites.forEach(c => {
      c.addEventListener('click', () => {
        const n = c.dataset.n;
        if (!n) return;
        const card = aBody.querySelector<HTMLElement>(`.tdr-hit[data-n="${n}"]`);
        if (!card) return;
        // カードが属するグループの header / list を畳みから展開。
        const list = card.closest<HTMLElement>('.tdr-sources');
        const hdr = list?.previousElementSibling as HTMLElement | null;
        hdr?.classList.remove('collapsed');
        list?.classList.remove('collapsed');
        card.scrollIntoView({ behavior: 'smooth', block: 'center' });
        card.classList.add('is-flash');
        setTimeout(() => card.classList.remove('is-flash'), 1200);
      });
    });
  }

  function renderSuggest(aBody: HTMLElement, qs: string[]): void {
    if (!qs.length) return;
    const row = el('div', { class: 'tdr-suggest' }, [
      el('span', { class: 'tdr-suggest-h' }, ['関連する質問']),
      ...qs.map(q => {
        const chip = el('button', { class: 'tdr-suggest-chip' }, [q]);
        chip.addEventListener('click', () => { input.value = q; autosize(); input.focus(); });
        return chip;
      }),
    ]);
    aBody.appendChild(row);
  }

  /** OneNote 追記候補ページ一覧 (最大 30 件) + 現在開いているページを system プロンプト用に整形。
   *  ページが 1 件も無ければ空文字を返す (LLM に余計な指示を出さない)。 */
  async function buildOneNoteCtx(relayBaseUrl: string, signal?: AbortSignal): Promise<string> {
    const eng = await getEngine(siteUrl);
    const excluded = getExcludedOneNotePageIds();
    const pages = eng.db.importedOneNotePages().filter(p => !excluded.has(p.pageId));
    if (pages.length === 0) return '';
    const sorted = [...pages].sort((a, b) => (b.lastModified || '').localeCompare(a.lastModified || ''));
    // 現在開いているページは取り込み済み 30 件のうちに入っていなかったら個別に注入する
    // (codex review 指摘: 30 件のリスト外だと AI の優先度 (2) =「現在開いているページ」が機能しなくなる)。
    const cur = await fetchCurrentOneNotePageId(relayBaseUrl, signal).catch(() => '');
    const recent = sorted.slice(0, 30);
    if (cur && !recent.some(p => p.pageId === cur)) {
      const old = sorted.find(p => p.pageId === cur);
      if (old) recent.unshift(old); // 直近 30 件の先頭に押し込む
    }
    const lines = recent.map(p => `[${p.pageId}] ${p.location} ・ ${p.title}`).join('\n');
    const hit = cur ? recent.find(p => p.pageId === cur) : null;
    const curHint = hit ? `\n\nOneNote で現在開いているページ: ${hit.pageId} (${hit.location} ・ ${hit.title})` : '';
    return `OneNote 追記候補ページ一覧 (取り込み済み):\n${lines}${curHint}`;
  }

  /** 出典フッターを Markdown で組み立てる。
   *  - 引用された [n] のヒットだけを採用 (回答中で実際に参照されたもの)。
   *  - メール: "[n] **件名** — 送信者 (YYYY-MM-DD HH:mm)"
   *  - OneNote: "[n] [ノートブック › セクション ・ タイトル](onenote:link)" (relay からリンクを引いた場合)
   *  リンクが取れなかった OneNote ヒットや、relay 不在時はラベルのみ。 */
  function buildSourcesFooter(
    answer: string, hits: SavedHit[], links: Map<string, string>,
    onenoteTitles: Map<string, string> = new Map(),
  ): string {
    const cited = new Set<number>();
    for (const m of answer.matchAll(/\[(\d+)\]/g)) cited.add(Number(m[1]));
    if (cited.size === 0 || hits.length === 0) return '';
    const fmtDate = (iso: string): string => {
      const d = new Date(iso); if (isNaN(d.getTime())) return iso || '';
      const Y = d.getFullYear(), M = String(d.getMonth() + 1).padStart(2, '0'), D = String(d.getDate()).padStart(2, '0');
      const hh = String(d.getHours()).padStart(2, '0'), mm = String(d.getMinutes()).padStart(2, '0');
      return `${Y}-${M}-${D} ${hh}:${mm}`;
    };
    const lines: string[] = ['', '## 出典'];
    hits.forEach((h, i) => {
      const n = i + 1;
      if (!cited.has(n)) return;
      if (h.kind === 'onenote') {
        // OneNote: タイトルに " - " が含まれる場合に subject の分割で誤検出するので
        // (codex review 指摘)、まず importedOneNotePages() の権威データを引く。
        // 見つからない時のフォールバックは chunkIdx===0 (素のタイトル) を優先、
        // それ以外は heading suffix を 1 回だけ剥がす。
        const pageId = h.conversationId || '';
        const titleFromDb = pageId ? onenoteTitles.get(pageId) : undefined;
        let pageTitle: string;
        if (titleFromDb) {
          pageTitle = titleFromDb;
        } else if ((h.chunkIdx ?? 0) === 0) {
          pageTitle = h.subject || '';
        } else {
          // " - heading" 末尾だけを 1 回剥がす (title 内の他の "-" は壊さない)。
          // pagesToIngestMails は subject = `${title} - ${heading}` の形で組むので、
          // 最後の " - " 以降を heading とみなす近似で十分。
          const subj = h.subject || '';
          const lastSep = subj.lastIndexOf(' - ');
          pageTitle = lastSep > 0 ? subj.slice(0, lastSep) : subj;
        }
        const loc = h.from || '';
        const label = loc ? `${loc} ・ ${pageTitle}` : pageTitle;
        const link = links.get(pageId) || '';
        lines.push(link ? `- [${n}] [${label}](${link})` : `- [${n}] ${label}`);
      } else {
        // mail: 件名 + 送信者 + 送信日時。
        const subj = (h.subject || '').replace(/\n/g, ' ');
        lines.push(`- [${n}] **${subj}** — ${h.from} (${fmtDate(h.date)})`);
      }
    });
    return lines.length > 2 ? lines.join('\n') : '';
  }

  function makeAppendOneNoteBtn(question: string, answer: string, relayBaseUrl: string, hits: SavedHit[]): HTMLElement {
    const btn = el('button', { class: 'tdr-copy', 'aria-label': 'OneNote に追記', title: 'AI の回答を OneNote に FAQ として追記' }, [
      el('span', { class: 'ic', html: icons.notebook(14) }),
      el('span', { class: 'lbl' }, ['OneNote に追記']),
    ]);
    btn.addEventListener('click', () => { void openAppendOneNoteModal(question, answer, relayBaseUrl, undefined, hits); });
    return btn;
  }

  /** AI 回答を OneNote に「FAQ エントリ」として追記する modal。
   *  追記先ページ選択 + 見出し編集 + Markdown 本文編集 + プレビュー + 確定。
   *  prefill が指定された場合 (AI が追記候補を生成した時) はそちらの値を初期値にする。 */
  async function openAppendOneNoteModal(
    question: string,
    answer: string,
    relayBaseUrl: string,
    prefill?: OneNoteAppendSuggestion,
    hits: SavedHit[] = [],
  ): Promise<void> {
    const eng = await getEngine(siteUrl);
    const excluded = getExcludedOneNotePageIds();
    const pages = eng.db.importedOneNotePages().filter(p => !excluded.has(p.pageId));
    if (pages.length === 0) {
      toast(root, 'OneNote 追記先がありません。先に「設定 → 取り込み」で OneNote ページを取り込んでください。', 'warn');
      return;
    }

    // "@@SUGGEST@@" 以降は OneNote に貼っても邪魔なので除去。引用番号 [n] は出典フッターと
    // 対応させるため残す。
    const answerMain = answer.replace(/@@SUGGEST@@[\s\S]*$/, '').trim();

    // OneNote ページタイトル参照用 (footer のタイトル分割問題対策)。
    const onenoteTitleMap = new Map<string, string>();
    for (const p of eng.db.importedOneNotePages()) onenoteTitleMap.set(p.pageId, p.title);

    // 出典フッター v1: リンク無しでまず作る (relay の応答を待たない)。
    // 後でリンク取得が返ってきたら本文を差し替える (codex review 指摘: 不調な relay で
    // モーダルが開かなくなる問題を回避)。
    const defaultHeading = prefill?.heading?.trim() || `Q: ${question}`;
    const initialFooter = buildSourcesFooter(answer, hits, new Map<string, string>(), onenoteTitleMap);
    const defaultBody = initialFooter ? `${answerMain}\n${initialFooter}` : answerMain;
    const headingInput = el('input', { type: 'text', class: 'tdr-input', value: defaultHeading }) as HTMLInputElement;
    const bodyArea = el('textarea', { class: 'tdr-input', rows: '12', style: 'min-height:280px;font-family:var(--font-mono);font-size:var(--fs-sm)' }) as HTMLTextAreaElement;
    bodyArea.value = defaultBody;
    const NEW_PAGE_SENTINEL = '__tdr_new_page__';
    const pageSelect = el('select', { class: 'tdr-input' }) as HTMLSelectElement;
    for (const p of pages) {
      const opt = el('option', { value: p.pageId }, [`${p.location} ・ ${p.title}`]) as HTMLOptionElement;
      pageSelect.appendChild(opt);
    }
    // 新規ページ作成オプション (relay から階層が取れた時のみ実体として機能)。
    const newPageOpt = el('option', { value: NEW_PAGE_SENTINEL }, ['+ 新規ページを作成…']) as HTMLOptionElement;
    pageSelect.appendChild(newPageOpt);
    const currentHint = el('p', { class: 'tdr-hint', style: 'margin-top:var(--s-1)' }, ['取り込み済みかつ除外していない OneNote ページから選択します。']);

    // 新規ページ作成用の入力欄 (デフォルト非表示)。
    const sectionSelect = el('select', { class: 'tdr-input' }) as HTMLSelectElement;
    const titleInput = el('input', { type: 'text', class: 'tdr-input', placeholder: '新規ページのタイトル' }) as HTMLInputElement;
    const newPageBlock = el('div', { style: 'display:none;border-left:3px solid var(--accent-strong);padding:var(--s-3) var(--s-4);background:var(--accent-soft);border-radius:var(--r-2);margin-top:var(--s-2);display:none' }, [
      el('div', { style: 'margin-bottom:var(--s-3)' }, [
        el('label', { class: 'tdr-label' }, ['保存先セクション']),
        sectionSelect,
      ]),
      el('div', {}, [
        el('label', { class: 'tdr-label' }, ['新規ページタイトル']),
        titleInput,
      ]),
    ]);
    // 階層を lazy 取得してセクション一覧に展開。失敗時は「新規ページ」オプションを無効化。
    let sectionsReady = false;
    void (async () => {
      try {
        const notebooks = await fetchOneNoteHierarchy(relayBaseUrl);
        for (const nb of notebooks) {
          for (const sec of nb.sections) {
            const opt = el('option', { value: sec.id }, [`${nb.name} › ${sec.name}`]) as HTMLOptionElement;
            sectionSelect.appendChild(opt);
          }
        }
        sectionsReady = sectionSelect.options.length > 0;
        if (!sectionsReady) newPageOpt.disabled = true;
      } catch {
        newPageOpt.disabled = true;
        newPageOpt.textContent = '+ 新規ページを作成… (relay 未接続のため使えません)';
      }
    })();

    if (prefill?.pageId && pages.some(p => p.pageId === prefill.pageId)) {
      // AI が選んだページがあればそれを既定にし、現在開いているページ取得は省略 (上書きされないように)。
      pageSelect.value = prefill.pageId;
      const hit = pages.find(p => p.pageId === prefill.pageId)!;
      currentHint.textContent = `AI が選んだページを既定にしています: ${hit.location} ・ ${hit.title}`;
    } else {
      // OneNote で「今開いているページ」を既定にする (取り込み済みなら)。
      void (async () => {
        const cur = await fetchCurrentOneNotePageId(relayBaseUrl).catch(() => '');
        if (!cur) return;
        const hit = pages.find(p => p.pageId === cur);
        if (hit) {
          pageSelect.value = cur;
          currentHint.textContent = `現在開いているページを既定に選択しています: ${hit.location} ・ ${hit.title}`;
          // update モードが既に有効になっている場合、ここで pageSelect を書き換えても
          // change ハンドラが走らないと Outline 一覧が古いままになるので明示的に発火 (codex review 指摘)。
          pageSelect.dispatchEvent(new Event('change', { bubbles: true }));
        } else {
          currentHint.textContent = '現在開いている OneNote ページはまだ取り込まれていないため、別のページを選んでください。';
        }
      })();
    }

    // ─── 更新モード用の UI: 既存の Tadori 追記 Outline を選んで上書きする ───
    let mode: 'append' | 'update' = 'append';
    const outlineSelect = el('select', { class: 'tdr-input' }) as HTMLSelectElement;
    const outlineStatus = el('p', { class: 'tdr-hint', style: 'margin-top:var(--s-1)' }, ['Tadori が追記した Outline を一覧から選んでください。']);
    const currentContentPane = el('div', { class: 'tdr-onenote-preview', style: 'border:1px solid var(--line);border-radius:var(--r-2);padding:var(--s-4);background:var(--paper-2);min-height:120px;max-height:200px;overflow:auto;font-size:var(--fs-sm);line-height:1.7;white-space:pre-wrap' });
    const updateBlock = el('div', { style: 'display:none;border-left:3px solid var(--accent-strong);padding:var(--s-3) var(--s-4);background:var(--accent-soft);border-radius:var(--r-2);margin-top:var(--s-2)' }, [
      el('div', { style: 'margin-bottom:var(--s-3)' }, [
        el('label', { class: 'tdr-label' }, ['更新する Tadori 追記 Outline']),
        outlineSelect,
        outlineStatus,
      ]),
      el('div', {}, [
        el('label', { class: 'tdr-label' }, ['現在 OneNote に書かれている内容']),
        currentContentPane,
      ]),
    ]);

    let currentOutlines: TadoriOutline[] = [];
    // 選択中ページの Tadori 追記 Outline 一覧を取得して outlineSelect に流し込む。
    async function reloadOutlinesForPage(): Promise<void> {
      const pageId = pageSelect.value;
      outlineSelect.replaceChildren();
      currentContentPane.textContent = '';
      currentOutlines = [];
      if (!pageId || pageId === NEW_PAGE_SENTINEL) {
        outlineStatus.textContent = '既存ページを選択してください。';
        return;
      }
      outlineStatus.textContent = '読み込み中…';
      try {
        const list = await fetchTadoriOutlines(relayBaseUrl, pageId);
        currentOutlines = list;
        if (list.length === 0) {
          outlineStatus.textContent = 'このページに Tadori 追記はまだありません。';
          return;
        }
        list.forEach((o, i) => {
          const label = `${o.heading || '(見出しなし)'}  —  ${o.banner.slice(0, 80)}`;
          const opt = el('option', { value: String(i) }, [label]) as HTMLOptionElement;
          outlineSelect.appendChild(opt);
        });
        outlineStatus.textContent = `${list.length} 件の Tadori 追記が見つかりました。`;
        applySelectedOutlineToEditor();
      } catch (e) {
        outlineStatus.textContent = `取得失敗: ${e instanceof Error ? e.message : String(e)}`;
      }
    }
    // outlineSelect の選択を見て、編集エリア (heading) と「現在内容」プレビューを更新する。
    // plainText はタグ除去済みのテキストで、bullets/links/太字 等のマークダウン構造が失われている。
    // そのまま bodyArea に入れて編集 → 上書きすると構造を失った状態で書き戻されてしまうため、
    // bodyArea は空にして「現在内容を見ながら新しく書き直す」方式にする (codex review 指摘の round-trip 損失対策)。
    function applySelectedOutlineToEditor(): void {
      const idx = Number(outlineSelect.value);
      const sel = currentOutlines[idx];
      if (!sel) return;
      currentContentPane.textContent = sel.plainText;
      headingInput.value = sel.heading;
      bodyArea.value = '';
      bodyArea.placeholder = '現在内容 (左) を見ながら新しい本文を Markdown で書いてください。空のまま [更新で上書き] すると本文なしの Outline になります。';
      renderPreview();
    }
    outlineSelect.addEventListener('change', applySelectedOutlineToEditor);

    // モード切替トグル
    const modeToggle = el('div', { style: 'display:flex;gap:var(--s-4);padding:var(--s-3) var(--s-4);background:var(--paper-2);border-radius:var(--r-2);border:1px solid var(--line)' });
    const mkModeRadio = (val: 'append' | 'update', label: string, desc: string): HTMLElement => {
      const radio = el('input', { type: 'radio', name: 'tdr-append-mode', value: val, style: 'margin-right:var(--s-2)' }) as HTMLInputElement;
      if (val === mode) radio.checked = true;
      const lbl = el('label', { style: 'display:flex;align-items:center;gap:var(--s-2);cursor:pointer;flex:1' }, [
        radio,
        el('div', {}, [
          el('div', { style: 'font-weight:600' }, [label]),
          el('div', { class: 'tdr-hint', style: 'margin:0' }, [desc]),
        ]),
      ]);
      radio.addEventListener('change', () => {
        if (radio.checked) { mode = val; applyMode(); }
      });
      return lbl;
    };
    modeToggle.append(
      mkModeRadio('append', '新規追記 / 新規ページ作成', '回答内容をそのまま追記。新規ページも作れます。'),
      mkModeRadio('update', '既存の Tadori 追記を更新', '過去に Tadori が追加したブロックを上書き (手書きは触らない)。'),
    );

    function applyMode(): void {
      if (mode === 'append') {
        updateBlock.style.display = 'none';
        newPageOpt.disabled = false;
        // pageSelect の値がそのまま使える状態に
        refreshNewPageBlock();
        confirmBtn.textContent = 'OneNote に追記';
      } else {
        // update モード: 新規ページオプションを無効化、新規ページ block を必ず隠す
        newPageBlock.style.display = 'none';
        newPageOpt.disabled = true;
        if (pageSelect.value === NEW_PAGE_SENTINEL) pageSelect.value = pages[0]?.pageId ?? '';
        updateBlock.style.display = '';
        void reloadOutlinesForPage();
        confirmBtn.textContent = '更新で上書き';
      }
    }

    // pageSelect 切替で新規ページ入力欄の表示を切る (append モード時) / 更新モードなら Outline 一覧を再読込。
    const refreshNewPageBlock = (): void => {
      newPageBlock.style.display = (mode === 'append' && pageSelect.value === NEW_PAGE_SENTINEL) ? '' : 'none';
    };
    pageSelect.addEventListener('change', () => {
      refreshNewPageBlock();
      if (mode === 'update') void reloadOutlinesForPage();
    });
    refreshNewPageBlock();

    const opUser = currentUser();
    const escapeText = (t: string): string => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

    const preview = el('div', { class: 'tdr-onenote-preview', style: 'border:1px solid var(--line);border-radius:var(--r-2);padding:var(--s-4);background:var(--paper-2);min-height:200px;max-height:300px;overflow:auto;font-size:var(--fs-sm);line-height:1.7' });
    const renderPreview = (): void => {
      const h = headingInput.value.trim();
      const md = bodyArea.value;
      const ts = new Date();
      const stamp = `${ts.getFullYear()}-${String(ts.getMonth() + 1).padStart(2, '0')}-${String(ts.getDate()).padStart(2, '0')} ${String(ts.getHours()).padStart(2, '0')}:${String(ts.getMinutes()).padStart(2, '0')}`;
      // Tadori が誰の操作で追記したかを示すバナー (OneNote 側でも同じ書式で書き込まれる)
      const bannerHtml = `<div style="color:#888;font-size:11px;margin-bottom:4px"><b>[Tadori 追記]</b> by ${escapeText(opUser)} [${stamp}]</div>`;
      const headingHtml = h ? `<div style="font-weight:700;margin-bottom:6px">${escapeText(h)}</div>` : '';
      preview.innerHTML = bannerHtml + headingHtml + renderMarkdown(md);
    };
    headingInput.addEventListener('input', renderPreview);
    bodyArea.addEventListener('input', renderPreview);
    renderPreview();

    const body = el('div', { class: 'tdr-modal-body', style: 'display:flex;flex-direction:column;gap:var(--s-4)' }, [
      modeToggle,
      el('div', {}, [
        el('label', { class: 'tdr-label' }, ['追記先ページ']), pageSelect,
        currentHint,
        newPageBlock,
        updateBlock,
      ]),
      el('div', {}, [
        el('label', { class: 'tdr-label' }, ['見出し']), headingInput,
      ]),
      el('div', { style: 'display:grid;grid-template-columns:1fr 1fr;gap:var(--s-4)' }, [
        el('div', {}, [
          el('label', { class: 'tdr-label' }, ['本文 (Markdown)']), bodyArea,
        ]),
        el('div', {}, [
          el('label', { class: 'tdr-label' }, ['プレビュー']), preview,
        ]),
      ]),
    ]);

    const cancelBtn  = el('button', { class: 'tdr-btn' }, ['キャンセル']);
    const confirmBtn = el('button', { class: 'tdr-btn tdr-btn--primary' }, ['OneNote に追記']);
    const footer = el('div', { class: 'tdr-modal-footer' }, [cancelBtn, confirmBtn]);

    const handle = openModal({ root, title: 'OneNote に追記', body, footer, xlarge: true });
    cancelBtn.addEventListener('click', () => handle.close());

    // モーダルを開いた後で OneNote ハイパーリンクを非同期取得 → 本文をリビルド。
    // 取得が遅い/失敗してもモーダル自体は止めない。ユーザが本文を編集していなければ差し替える。
    const onenotePageIds = hits.filter(h => h.kind === 'onenote' && h.conversationId).map(h => h.conversationId);
    if (onenotePageIds.length > 0) {
      void (async () => {
        const oneLinks = await fetchOneNoteLinks(relayBaseUrl, [...new Set(onenotePageIds)]).catch(() => new Map<string, string>());
        if (oneLinks.size === 0) return;
        const newFooter = buildSourcesFooter(answer, hits, oneLinks, onenoteTitleMap);
        const newBody = newFooter ? `${answerMain}\n${newFooter}` : answerMain;
        // ユーザが既に編集してたら触らない。
        if (bodyArea.value === defaultBody) { bodyArea.value = newBody; renderPreview(); }
      })();
    }
    confirmBtn.addEventListener('click', async () => {
      const heading = headingInput.value.trim();
      const blocks = markdownToBlocks(bodyArea.value);
      if (!heading && blocks.length === 0) { toast(root, '見出しまたは本文を入力してください', 'warn'); return; }
      const origText = confirmBtn.textContent;
      confirmBtn.disabled = true; cancelBtn.disabled = true;
      confirmBtn.textContent = mode === 'update' ? '更新中…' : '追記中…';
      try {
        if (mode === 'update') {
          const pageId = pageSelect.value;
          const idx = Number(outlineSelect.value);
          const sel = currentOutlines[idx];
          if (!pageId || !sel) { toast(root, '更新する Outline を選択してください', 'warn'); throw new Error('skip'); }
          await replaceTadoriOutline(relayBaseUrl, { pageId, outlineId: sel.outlineId, heading, blocks, user: opUser });
          toast(root, 'OneNote の Tadori 追記を更新しました', 'ok');
          handle.close();
        } else {
          const selected = pageSelect.value;
          const args: Parameters<typeof appendOneNotePage>[1] = { heading, blocks, user: opUser };
          if (selected === NEW_PAGE_SENTINEL) {
            const sectionId = sectionSelect.value;
            const newTitle = titleInput.value.trim();
            if (!sectionId) { toast(root, '保存先セクションを選択してください', 'warn'); throw new Error('skip'); }
            if (!newTitle) { toast(root, '新規ページタイトルを入力してください', 'warn'); titleInput.focus(); throw new Error('skip'); }
            args.createInSection = sectionId;
            args.newPageTitle = newTitle;
          } else {
            args.pageId = selected;
          }
          await appendOneNotePage(relayBaseUrl, args);
          toast(root, args.createInSection ? 'OneNote に新規ページを作成して追記しました' : 'OneNote に追記しました', 'ok');
          handle.close();
        }
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        if (msg !== 'skip') toast(root, `失敗: ${msg}`, 'error');
        confirmBtn.disabled = false; cancelBtn.disabled = false;
        confirmBtn.textContent = origText ?? 'OneNote に追記';
      }
    });
  }

  function makeCopyBtn(text: string): HTMLElement {
    const btn = el('button', { class: 'tdr-copy', 'aria-label': '回答をコピー', title: '回答をコピー' }, [
      el('span', { class: 'ic', html: icons.copy(14) }),
      el('span', { class: 'lbl' }, ['コピー']),
    ]);
    btn.addEventListener('click', () => {
      void navigator.clipboard?.writeText(text).then(() => {
        btn.classList.add('is-done');
        btn.replaceChildren(el('span', { class: 'ic', html: icons.check(14) }), el('span', { class: 'lbl' }, ['コピーしました']));
        setTimeout(() => {
          btn.classList.remove('is-done');
          btn.replaceChildren(el('span', { class: 'ic', html: icons.copy(14) }), el('span', { class: 'lbl' }, ['コピー']));
        }, 1500);
      }).catch(() => toast(root, 'コピーに失敗しました', 'error'));
    });
    return btn;
  }

  function renderSession(s: ChatSession): void {
    thread.replaceChildren();
    hasTurns = s.turns.length > 0;
    if (!hasTurns) { thread.appendChild(emptyState); return; }
    const relayBaseUrl = loadSettings().relayBaseUrl;
    for (const t of s.turns) {
      const refs = buildTurn(t.q);
      finalizeTurn(refs, t.answer, t.hits, t.ms, relayBaseUrl, t.q, t.yen, t.createdAt);
    }
    scrollBottom();
  }

  /** 直近 3 ターンを履歴として渡す (フォローアップ質問の文脈用)。回答はトリム。 */
  function buildHistory(): ChatHistoryMsg[] {
    const sess = getSession(currentId);
    if (!sess) return [];
    const out: ChatHistoryMsg[] = [];
    for (const t of sess.turns.slice(-3)) {
      out.push({ role: 'user', content: t.q });
      out.push({ role: 'assistant', content: t.answer.slice(0, 600) });
    }
    return out;
  }

  /** 1 ターンの共通処理: ヒット取得 → (任意で再ランカー) → 生成 → 整形 → 保存。 */
  async function converse(opts: {
    displayQ: string;
    llmQuestion: string;
    loadingLabel: string;
    getHits: (signal: AbortSignal) => Promise<SavedHit[]>;
    /** 再ランカー (RAG 検索向け) を適用するか。経緯要約等は false。 */
    rerankable?: boolean;
  }): Promise<void> {
    if (generating) return;
    if (!hasTurns) { emptyState.remove(); hasTurns = true; }

    const refs = buildTurn(opts.displayQ);
    scrollBottom();
    refs.answerText.replaceChildren(thinkingEl(opts.loadingLabel));

    abort?.abort();
    abort = new AbortController();
    const signal = abort.signal;
    const s = loadSettings();
    const history = buildHistory();
    setGenerating(true);

    let full = '';
    let aiTitle = '';
    let suggestions: string[] = [];
    let hits: SavedHit[] = [];
    let yen: number | undefined = undefined;
    let appendSuggestion: OneNoteAppendSuggestion | null = null;
    let t0 = performance.now();
    const createdAt = new Date().toISOString();
    // ストリーム中の rAF キャンセル用 (try/catch をまたぐので外で宣言)
    let pendingRafId: number | null = null;
    const cancelStreamRender = (): void => {
      if (pendingRafId != null) { cancelAnimationFrame(pendingRafId); pendingRafId = null; }
    };

    const save = (): void => {
      if (!full.trim()) return;
      const ms = Math.round(performance.now() - t0);
      finalizeTurn(refs, full, hits, ms, s.relayBaseUrl, opts.displayQ, yen, createdAt);
      if (suggestions.length) renderSuggest(refs.aBody, suggestions);
      const saved = appendTurn(currentId, { q: opts.displayQ, answer: full, hits, ms, yen, createdAt });
      if (saved.turns.length === 1 && aiTitle) setTitle(currentId, aiTitle);
      refreshList();
      // AI が OneNote 追記候補を出した場合は、回答 finalize 直後に確認モーダルを自動で開く。
      // 旧仕様の「回答下に通知バーを置いて 1 クリック」よりも 1 アクション減る。
      if (appendSuggestion) {
        void openAppendOneNoteModal(opts.displayQ, full, s.relayBaseUrl, appendSuggestion, hits);
      }
    };

    const addUsage = (u: { yen: number }): void => { yen = (yen ?? 0) + u.yen; };

    try {
      hits = await opts.getHits(signal);
      if (hits.length === 0) {
        refs.answerText.textContent = '該当するメールが見つかりませんでした。';
        return;
      }

      // 再ランカー: 取得済み候補を LLM で並べ替えて、上位 ragTopK を残す。
      if (opts.rerankable && s.rerankEnabled && hits.length > 1) {
        refs.answerText.replaceChildren(thinkingEl('候補を絞り込み中'));
        const candidates = hits.map(h => ({
          subject: h.subject, from: h.from, date: h.date,
          body: cleanBody(h.isHtml ? htmlToText(h.body) : h.body),
        }));
        const order = await rerankByLLM(opts.llmQuestion, candidates, s, signal, addUsage);
        const reordered = order.map(i => hits[i]).filter(Boolean) as SavedHit[];
        hits = reordered.slice(0, s.ragTopK);
      }

      refs.answerText.replaceChildren(thinkingEl('回答を生成中'));

      // RAG にはプレーンテキスト + 引用履歴を剥がした「新規発言だけ」を渡す。
      // Re: メール間で引用が重複してプロンプトを膨らませるのを防ぐ。
      const sources: RagSource[] = hits.map((h, i) => ({
        n: i + 1, subject: h.subject, from: h.from, date: h.date,
        body: cleanBody(h.isHtml ? htmlToText(h.body) : h.body),
      }));

      // OneNote 追記候補を AI に判断させるための文脈 (取り込み済みページ一覧 + 現在開いているページ)。
      // 一覧が無ければプロンプトに含めない (LLM への余計な指示を増やさない)。
      const onenoteCtx = await buildOneNoteCtx(s.relayBaseUrl, signal).catch(() => '');

      t0 = performance.now();
      let firstDelta = true;
      // ストリーム中の markdown レンダリングは rAF で 1 フレーム 1 回までスロットル。
      // 毎トークン renderMarkdown を回すと重いし、フレーム境界に揃えるとちらつきも減る。
      // 保留中の rAF は finalize 前にキャンセル (codex review 指摘: 最終 delta と [DONE] が
      // 同フレームで来ると finalize 後に rAF が走り、wireCiteJump で付与したクリックハンドラが消える)。
      const scheduleStreamRender = (): void => {
        if (pendingRafId != null) return;
        pendingRafId = requestAnimationFrame(() => {
          pendingRafId = null;
          refs.answerText.innerHTML = renderMarkdown(full).replace(
            /\[(\d+)\]/g,
            (_, n) => `<span class="cite" data-n="${n}">[${n}]</span>`,
          );
          scrollBottom();
        });
      };
      await generateAnswer(opts.llmQuestion, sources, s, delta => {
        full += delta;
        if (firstDelta) { refs.answerText.textContent = ''; firstDelta = false; } // 「生成中」を消す
        scheduleStreamRender();
      }, {
        signal,
        onTitle: title => { aiTitle = title; },
        history,
        onSuggest: qs => { suggestions = qs; },
        onUsage: addUsage,
        onenoteCtx: onenoteCtx || undefined,
        onAppendSuggestion: sug => { appendSuggestion = sug; },
      });

      cancelStreamRender();
      save();

    } catch (err: unknown) {
      if (signal.aborted || (err instanceof Error && err.name === 'AbortError')) {
        cancelStreamRender();
        save(); // 停止までに生成した分は確定・保存
        return;
      }
      refs.answerText.textContent = 'エラーが発生しました。';
      toast(root, `エラー: ${err instanceof Error ? err.message : String(err)}`, 'error');
    } finally {
      setGenerating(false);
      abort = null;
    }
  }

  async function submit(): Promise<void> {
    if (generating) return;
    const q = input.value.trim();
    if (!q) return;
    input.value = '';
    autosize();
    await converse({
      displayQ: q,
      llmQuestion: q,
      loadingLabel: 'クエリを解析中',
      rerankable: true,
      getHits: async (signal) => {
        const s = loadSettings();
        // LLM クエリルータ: ID/固有名詞などは keywords 完全一致 + 意味文はベクトル検索、と分担。
        // フォローアップ質問 (「要約して」等) の指示語を直前会話から解決させるため history を渡す。
        // 失敗時は元クエリそのままで通常検索にフォールバック (queryRouter 内で FALLBACK 処理)。
        const plan = await classifyQuery(q, s, signal, undefined, buildHistory());
        const topK = s.rerankEnabled ? Math.max(s.rerankCandidates, s.ragTopK) : s.ragTopK;
        const raw = await searchVectors(q, s, siteUrl, topK, {
          vectorQuery: plan.vectorQuery,
          mustContain: plan.keywords,
        });
        const rules = loadRules();
        const afterExclude = rules.length ? raw.filter(h => !matchesAnyRule(h, rules)) : raw;
        let filtered = afterExclude.filter(h => h.score >= s.ragMinScore);
        if (filtered.length === 0 && afterExclude.length > 0) filtered = [afterExclude[0]];
        return filtered as SavedHit[];
      },
    });
  }

  /** 同一スレッド (conversationId) の全メールを時系列で要約する。 */
  async function summarizeThread(hit: SavedHit): Promise<void> {
    if (generating || !hit.conversationId) return;
    await converse({
      displayQ: `「${hit.subject}」の経緯`,
      llmQuestion: '上記は同一スレッドのメールを古い順に並べたものです。やり取りの経緯・決定事項・現状・残課題を時系列で簡潔に要約してください。',
      loadingLabel: 'スレッドを読み込み中',
      getHits: async () => await getThread(siteUrl, hit.conversationId) as SavedHit[],
    });
  }

  /** 出典カード 1 件を生成 (n は引用番号 = hits 内の元インデックス+1)。 */
  function renderHitCard(h: SavedHit, n: number, relayBaseUrl: string, query: string): HTMLElement {
    const kind = h.kind || 'mail';
    const plain = h.isHtml ? htmlToText(h.body) : h.body;
    const hitEl = el('div', { class: `tdr-hit tdr-hit--${kind}` });
    hitEl.dataset.n = String(n); // [n] クリックでの引き当て用
    const badgeIcon = kind === 'onenote' ? icons.notebook(12) : kind === 'doc' ? icons.folder(12) : icons.message(12);
    const badgeLabel = kind === 'onenote' ? 'OneNote' : kind === 'doc' ? '文書' : 'メール';
    const head = el('div', { class: 'tdr-hit-head' }, [
      el('span', { class: 'tdr-hit-num' }, [String(n)]),
      el('span', { class: 'tdr-hit-badge', title: badgeLabel, 'aria-label': badgeLabel, html: badgeIcon }),
      el('span', { class: 'tdr-hit-subject' }, [h.subject]),
      el('span', { class: 'tdr-hit-score' }, [h.score.toFixed(3)]),
    ]);
    if (kind === 'mail' && h.internetMessageId) {
      const openBtn = el('button', {
        class: 'tdr-hit-open', 'aria-label': 'Outlook で開く', title: 'Outlook で開く',
        html: icons.external(14),
      });
      openBtn.addEventListener('click', async (e) => {
        e.stopPropagation();
        openBtn.disabled = true;
        try { await openMailInOutlook(relayBaseUrl, h.internetMessageId); }
        catch (err: unknown) { toast(root, `Outlook 表示に失敗: ${err instanceof Error ? err.message : String(err)}`, 'error'); }
        finally { openBtn.disabled = false; }
      });
      head.appendChild(openBtn);
    } else if (kind === 'onenote' && h.conversationId) {
      // OneNote では conversationId = ページID。relay 経由で OneNote 上に表示。
      const openBtn = el('button', {
        class: 'tdr-hit-open', 'aria-label': 'OneNote で開く', title: 'OneNote で開く',
        html: icons.external(14),
      });
      openBtn.addEventListener('click', async (e) => {
        e.stopPropagation();
        openBtn.disabled = true;
        try { await openOneNotePage(relayBaseUrl, h.conversationId); }
        catch (err: unknown) { toast(root, `OneNote 表示に失敗: ${err instanceof Error ? err.message : String(err)}`, 'error'); }
        finally { openBtn.disabled = false; }
      });
      head.appendChild(openBtn);
    }
    if (h.conversationId) {
      const sumBtn = el('button', {
        class: 'tdr-hit-open',
        'aria-label': kind === 'onenote' ? 'ページ全体を要約' : '経緯を要約',
        title: kind === 'onenote' ? '同じページのチャンクをまとめて要約' : '同じスレッドのやり取りを時系列で要約',
        html: icons.list(14),
      });
      sumBtn.addEventListener('click', (e) => { e.stopPropagation(); void summarizeThread(h); });
      head.appendChild(sumBtn);
    }
    hitEl.appendChild(head);
    const meta = kind === 'onenote'
      ? `${h.from}  ${h.date.slice(0, 10)}${typeof h.chunkIdx === 'number' && h.chunkCount && h.chunkCount > 1 ? `  ・ チャンク #${h.chunkIdx + 1}/${h.chunkCount}` : ''}`
      : `${h.from}  ${h.date.slice(0, 10)}`;
    hitEl.appendChild(el('div', { class: 'tdr-hit-from' }, [meta]));
    const snippetEl = el('div', { class: 'tdr-hit-snippet' });
    highlightInto(snippetEl, plain.slice(0, 140) + (plain.length > 140 ? '…' : ''), query);
    hitEl.appendChild(snippetEl);

    let detail: HTMLElement | null = null;
    hitEl.addEventListener('click', () => {
      hitEl.classList.toggle('is-open');
      if (!detail) {
        detail = el('div', { class: 'tdr-hit-detail' });
        detail.appendChild(kind === 'onenote' ? renderOneNoteHeader(h) : renderMailHeader(h));
        if (kind === 'onenote') {
          const wrap = el('div', { class: 'tdr-hit-detail-body' });
          renderMailBody(wrap, h.body, false);
          detail.appendChild(wrap);
        } else {
          detail.appendChild(renderMailBodyWithHistoryToggle(h));
        }
        hitEl.appendChild(detail);
      } else { detail.hidden = !detail.hidden; }
    });
    return hitEl;
  }

  /** OneNote ページ用ヘッダ (件名/最終更新/ノートブック › セクション/チャンク)。 */
  function renderOneNoteHeader(h: SavedHit): HTMLElement {
    const fmtDate = (iso: string): string => {
      const d = new Date(iso); if (isNaN(d.getTime())) return iso || '';
      const Y = d.getFullYear(), M = d.getMonth() + 1, D = d.getDate();
      const hh = String(d.getHours()).padStart(2, '0'), mm = String(d.getMinutes()).padStart(2, '0');
      return `${Y}-${String(M).padStart(2, '0')}-${String(D).padStart(2, '0')} ${hh}:${mm}`;
    };
    const row = (label: string, value: string): HTMLElement | null => {
      if (!value) return null;
      return el('div', { class: 'tdr-hit-hdr-row' }, [
        el('span', { class: 'tdr-hit-hdr-label' }, [label]),
        el('span', { class: 'tdr-hit-hdr-value' }, [value]),
      ]);
    };
    const chunkInfo = (typeof h.chunkIdx === 'number' && h.chunkCount && h.chunkCount > 1)
      ? `#${h.chunkIdx + 1} / 全 ${h.chunkCount}` : '';
    const rows = [
      row('タイトル', h.subject),
      row('場所', h.from),
      row('最終更新', fmtDate(h.date)),
      row('チャンク', chunkInfo),
    ].filter((x): x is HTMLElement => x !== null);
    return el('div', { class: 'tdr-hit-hdr' }, rows);
  }

  /** 本文を「新規発言」+ 折りたたみ式「過去のやり取り」で描画する。 */
  function renderMailBodyWithHistoryToggle(h: SavedHit): HTMLElement {
    const wrap = el('div', { class: 'tdr-hit-detail-body' });
    const sp = h.isHtml ? splitHtmlReplyHistory(h.body) : splitReplyHistory(h.body);
    const head = sp.head.trim();
    const tail = sp.tail.trim();
    // 履歴が無い or 新規発言が空 (=全部履歴) なら分割せずそのまま描画。
    if (!tail || !head) { renderMailBody(wrap, h.body, h.isHtml); return wrap; }

    const headHost = el('div', { class: 'tdr-hit-body-head' });
    renderMailBody(headHost, head, h.isHtml);
    wrap.appendChild(headHost);

    const btn = el('button', { class: 'tdr-hit-more', type: 'button' }, [
      el('span', { class: 'ic', html: icons.chevron(13) }),
      el('span', { class: 'lbl' }, ['過去のやり取りを表示']),
    ]);
    const tailHost = el('div', { class: 'tdr-hit-body-tail' });
    tailHost.hidden = true;
    renderMailBody(tailHost, tail, h.isHtml);

    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const opening = tailHost.hidden;
      tailHost.hidden = !opening;
      btn.classList.toggle('is-open', opening);
      const lbl = btn.querySelector('.lbl');
      if (lbl) lbl.textContent = opening ? '過去のやり取りを隠す' : '過去のやり取りを表示';
    });

    wrap.append(btn, tailHost);
    return wrap;
  }

  /** 参考メール展開時のヘッダ (件名 / 送信日 / From / To / Cc)。 */
  function renderMailHeader(h: SavedHit): HTMLElement {
    const fmtDate = (iso: string): string => {
      const d = new Date(iso); if (isNaN(d.getTime())) return iso || '';
      const Y = d.getFullYear(), M = d.getMonth() + 1, D = d.getDate();
      const hh = String(d.getHours()).padStart(2, '0'), mm = String(d.getMinutes()).padStart(2, '0');
      return `${Y}-${String(M).padStart(2, '0')}-${String(D).padStart(2, '0')} ${hh}:${mm}`;
    };
    const row = (label: string, value: string | string[]): HTMLElement | null => {
      const v = Array.isArray(value) ? value.filter(Boolean).join(', ') : value;
      if (!v) return null;
      return el('div', { class: 'tdr-hit-hdr-row' }, [
        el('span', { class: 'tdr-hit-hdr-label' }, [label]),
        el('span', { class: 'tdr-hit-hdr-value' }, [v]),
      ]);
    };
    const rows = [
      row('件名', h.subject),
      row('送信日', fmtDate(h.date)),
      row('From', h.from),
      row('To', h.to || []),
      row('Cc', h.cc || []),
    ].filter((x): x is HTMLElement => x !== null);
    return el('div', { class: 'tdr-hit-hdr' }, rows);
  }

  /** 出典グループ (ヘッダ + リスト) を 1 つ container に追加。 */
  function renderSourceGroup(
    container: HTMLElement, label: string, entries: Array<{ h: SavedHit; n: number }>,
    relayBaseUrl: string, query: string, defaultCollapsed: boolean,
  ): void {
    const cls = defaultCollapsed ? ' collapsed' : '';
    const hdr  = el('div', { class: `tdr-sources-h${cls}` }, [
      el('span', { html: icons.chevron(14) }),
      el('span', {}, [`${label} (${entries.length})`]),
    ]);
    const list = el('div', { class: `tdr-sources${cls}` });
    hdr.addEventListener('click', () => { hdr.classList.toggle('collapsed'); list.classList.toggle('collapsed'); });
    for (const { h, n } of entries) list.appendChild(renderHitCard(h, n, relayBaseUrl, query));
    container.append(hdr, list);
  }

  function appendSources(
    container: HTMLElement, hits: SavedHit[], relayBaseUrl: string, query = '', cited?: Set<number>,
  ): void {
    const all = hits.map((h, i) => ({ h, n: i + 1 }));
    const hasCitedSplit = !!cited && cited.size > 0 && cited.size < hits.length;
    if (hasCitedSplit) {
      const citedEntries = all.filter(e => cited!.has(e.n));
      const otherEntries = all.filter(e => !cited!.has(e.n));
      // 引用されたカードは既定で展開、その他は折りたたみ。
      renderSourceGroup(container, '引用された参照メール', citedEntries, relayBaseUrl, query, false);
      renderSourceGroup(container, '他に検索された候補', otherEntries, relayBaseUrl, query, true);
    } else {
      // 引用ゼロ or 全件引用なら 1 グループにまとめる (既定で折りたたみ)。
      renderSourceGroup(container, '参照メール', all, relayBaseUrl, query, true);
    }
  }

  function scrollBottom(): void { thread.scrollTop = thread.scrollHeight; }

  function autosize(): void {
    input.style.height = 'auto';
    input.style.height = Math.min(input.scrollHeight, window.innerHeight * 0.4) + 'px';
  }

  input.addEventListener('input', autosize);
  input.addEventListener('keydown', (e: KeyboardEvent) => {
    if (e.key !== 'Enter' || e.isComposing) return; // IME 変換中は無視
    const enterSends = loadSettings().enterSends;
    if (enterSends) {
      // Enter 単独で送信 / Shift+Enter は改行
      if (!e.shiftKey && !e.metaKey && !e.ctrlKey) { e.preventDefault(); void submit(); }
    } else {
      // ⌘/Ctrl+Enter で送信
      if (e.metaKey || e.ctrlKey) { e.preventDefault(); void submit(); }
    }
  });
  // 生成中は停止、それ以外は送信。
  sendBtn.addEventListener('click', () => { if (generating) abort?.abort(); else void submit(); });

  // 定型プロンプトチップ (クリックで入力欄に挿入)。
  const QUICK_PROMPTS = ['最近の重要なお知らせは?', '今後のイベント・予定を教えて', '締め切り・期限のある依頼は?', 'システム・メンテナンスの予定は?'];
  const chipsRow = el('div', { class: 'tdr-chips' }, QUICK_PROMPTS.map(p => {
    const chip = el('button', { class: 'tdr-chip' }, [p]);
    chip.addEventListener('click', () => { input.value = p; autosize(); input.focus(); });
    return chip;
  }));

  // ── レイアウト: 左ペイン | ドラッグ可能な仕切り | チャット ──
  thread.appendChild(emptyState);
  refreshList();

  const sidebar = el('aside', { class: 'tdr-sidebar' }, [
    el('div', { class: 'tdr-sidebar-head' }, [newBtn, searchBtn]),
    sessionList,
  ]);
  const storedW = Number(localStorage.getItem(SIDEBAR_W_KEY) || '');
  sidebar.style.width = `${clampW(storedW || 260)}px`;

  const divider = el('div', { class: 'tdr-divider', 'aria-label': '幅を調整', title: 'ドラッグで幅を調整' });
  attachDividerDrag(divider, sidebar);

  const chatCol = el('div', { class: 'tdr-chat' }, [
    thread,
    el('div', { class: 'tdr-composer' }, [
      el('div', { class: 'tdr-composer-inner' }, [
        chipsRow,
        buildModelPicker(),
        el('div', { class: 'tdr-note-form' }, [input, sendBtn]),
        hintEl,
      ]),
    ]),
  ]);

  return el('div', { class: 'tdr-main' }, [sidebar, divider, chatCol]);
}

/** 入力ボックス上のモデル切替 (Spira と同じ作法)。社内AI + (開発者モード時) Claude。 */
function buildModelPicker(): HTMLSelectElement {
  const sel = el('select', { class: 'tdr-model-pick', title: 'プロバイダ / モデル' }) as HTMLSelectElement;

  function sync(): void {
    const s = loadSettings();
    const cur = `${s.provider}:${s.provider === 'claude' ? s.claudeModel : s.chatModel}`;
    sel.replaceChildren();
    const corp = el('optgroup', { label: '社内 AI' },
      CORP_AI_MODELS.map(m => el('option', { value: `corp:${m.id}` }, [m.id])));
    sel.appendChild(corp);
    if (isDeveloperMode()) {
      const claude = el('optgroup', { label: 'Claude' },
        CLAUDE_MODELS.map(m => el('option', { value: `claude:${m.id}` }, [m.label])));
      sel.appendChild(claude);
    }
    sel.value = cur;
  }

  sel.addEventListener('change', () => {
    const i = sel.value.indexOf(':');
    if (i < 0) return;
    const provider = sel.value.slice(0, i);
    const modelId = sel.value.slice(i + 1);
    if (provider === 'claude') saveSettings({ provider: 'claude', claudeModel: modelId });
    else saveSettings({ provider: 'corp', chatModel: modelId });
    sync();
  });

  sync();
  return sel;
}

function clampW(w: number): number {
  return Math.max(SIDEBAR_MIN, Math.min(SIDEBAR_MAX, Math.round(w)));
}

function attachDividerDrag(divider: HTMLElement, sidebar: HTMLElement): void {
  divider.addEventListener('pointerdown', (e: PointerEvent) => {
    e.preventDefault();
    const startX = e.clientX;
    const startW = sidebar.getBoundingClientRect().width;
    divider.classList.add('is-dragging');
    divider.setPointerCapture(e.pointerId);

    const onMove = (ev: PointerEvent) => {
      sidebar.style.width = `${clampW(startW + (ev.clientX - startX))}px`;
    };
    const onUp = () => {
      divider.classList.remove('is-dragging');
      divider.removeEventListener('pointermove', onMove);
      divider.removeEventListener('pointerup', onUp);
      try { localStorage.setItem(SIDEBAR_W_KEY, String(Math.round(sidebar.getBoundingClientRect().width))); } catch { /* quota */ }
    };
    divider.addEventListener('pointermove', onMove);
    divider.addEventListener('pointerup', onUp);
  });
}
