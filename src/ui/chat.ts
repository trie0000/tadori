// RAG チャットパネル。左ペイン (セッション一覧) + スレッド表示 + コンポーザ。
// 送信 → searchVectors → generateAnswer (streaming) → 出典カード。
// セッションは localStorage に保存し、左ペインから再表示できる (src/ui/sessions.ts)。

import { el } from '../lib/dom';
import { icons } from './icons';
import { toast } from './toast';
import { searchVectors } from '../search/vectorSearch';
import { htmlToText, renderMailBody } from '../lib/mailhtml';
import { generateAnswer, type RagSource } from '../rag/client';
import { loadSettings, saveSettings, CORP_AI_MODELS, CLAUDE_MODELS } from '../api/aiSettings';
import { isDeveloperMode } from '../utils/devMode';
import { renderMarkdown } from '../lib/markdown';
import { openMailInOutlook } from '../outlook/import';
import { confirmModal } from './modal';
import {
  listSessions, getSession, appendTurn, setTitle, deleteSession, newSessionId,
  type ChatSession, type SavedHit,
} from './sessions';

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
  const sendBtn = el('button', { class: 'tdr-note-submit', 'aria-label': '送信', html: icons.send(16) });

  const emptyState = el('div', { class: 'tdr-empty' }, [
    el('div', { class: 'big' }, ['辿り']),
    el('p', {}, ['社内メーリングリストを自然言語で検索できます。']),
    el('p', { style: 'font-size:var(--fs-sm);color:var(--ink-4);margin-top:var(--s-3)' }, [
      '例: 「先月の懇親会の日程はいつですか?」',
    ]),
  ]);

  let abort: AbortController | null = null;
  let currentId = newSessionId();   // 現在のセッション (最初の送信まで未保存)
  let hasTurns = false;

  // ── 左ペイン (セッション一覧) ──
  const sessionList = el('div', { class: 'tdr-session-list' });
  const newBtn = el('button', { class: 'tdr-new-session' }, [
    el('span', { html: icons.plus(15) }), 'New session',
  ]);
  newBtn.addEventListener('click', startNewSession);

  function refreshList(): void {
    sessionList.replaceChildren();
    const sessions = listSessions();
    if (sessions.length === 0) {
      sessionList.appendChild(el('div', { class: 'tdr-session-empty' }, ['履歴はまだありません']));
      return;
    }
    for (const s of sessions) {
      const item = el('div', { class: 'tdr-session' + (s.id === currentId ? ' is-active' : '') }, [
        el('span', { class: 'tdr-session-ic', html: icons.message(14) }),
        el('span', { class: 'tdr-session-title', title: s.title }, [s.title]),
      ]);
      item.addEventListener('click', () => openSession(s.id));

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

  function finalizeTurn(refs: TurnRefs, fullMarkdown: string, hits: SavedHit[], ms: number, relayBaseUrl: string): void {
    refs.answerText.innerHTML = renderMarkdown(fullMarkdown).replace(
      /\[(\d+)\]/g,
      (_, n) => `<span class="cite" data-n="${n}">[${n}]</span>`,
    );
    refs.metaEl.replaceChildren(
      el('span', {}, [`${hits.length} 件参照`]),
      el('span', { class: 'mono' }, [`${ms} ms`]),
    );
    refs.aBody.appendChild(makeCopyBtn(fullMarkdown));
    if (hits.length) appendSources(refs.aBody, hits, relayBaseUrl);
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
      finalizeTurn(refs, t.answer, t.hits, t.ms, relayBaseUrl);
    }
    scrollBottom();
  }

  async function submit(): Promise<void> {
    const q = input.value.trim();
    if (!q || sendBtn.disabled) return;
    input.value = '';
    autosize();
    sendBtn.disabled = true;

    if (!hasTurns) { emptyState.remove(); hasTurns = true; }

    const refs = buildTurn(q);
    scrollBottom();
    refs.answerText.replaceChildren(thinkingEl('メールを検索中'));

    abort?.abort();
    abort = new AbortController();
    const s = loadSettings();

    try {
      const raw = await searchVectors(q, s, siteUrl, s.ragTopK);
      // スコアしきい値で足切り。全部下回ったら最良 1 件だけ残す (取りこぼし防止)。
      let hits = raw.filter(h => h.score >= s.ragMinScore);
      if (hits.length === 0 && raw.length > 0) hits = [raw[0]];
      if (hits.length === 0) {
        refs.answerText.textContent = '該当するメールが見つかりませんでした。';
        return;
      }
      refs.answerText.replaceChildren(thinkingEl('回答を生成中'));

      // RAG には本文をプレーンテキストで渡す (HTML はタグを除去)。
      const sources: RagSource[] = hits.map((h, i) => ({
        n: i + 1, subject: h.subject, from: h.from, date: h.date,
        body: h.isHtml ? htmlToText(h.body) : h.body,
      }));

      const t0 = performance.now();

      let full = '';
      let firstDelta = true;
      let aiTitle = '';
      await generateAnswer(q, sources, s, delta => {
        full += delta;
        if (firstDelta) { refs.answerText.textContent = ''; firstDelta = false; } // 「生成中」を消す
        refs.answerText.textContent = full; // ストリーム中はプレーン (タイプ感)
        scrollBottom();
      }, abort.signal, title => { aiTitle = title; });

      const ms = Math.round(performance.now() - t0);
      finalizeTurn(refs, full, hits, ms, s.relayBaseUrl);

      // 履歴へ保存 (MailHit は SavedHit と同形)。
      const saved = appendTurn(currentId, { q, answer: full, hits: hits as SavedHit[], ms });
      // 新規セッションの最初の質問は、回答と同時に得たタイトル (TITLE 行) を採用。
      if (saved.turns.length === 1 && aiTitle) setTitle(currentId, aiTitle);
      refreshList();

    } catch (err: unknown) {
      if (err instanceof Error && err.name === 'AbortError') return;
      refs.answerText.textContent = 'エラーが発生しました。';
      toast(root, `エラー: ${err instanceof Error ? err.message : String(err)}`, 'error');
    } finally {
      sendBtn.disabled = false;
      abort = null;
    }
  }

  function appendSources(container: HTMLElement, hits: SavedHit[], relayBaseUrl: string): void {
    // 既定は折りたたみ (collapsed)。ヘッダクリックで開閉。
    const hdrEl  = el('div', { class: 'tdr-sources-h collapsed' }, [
      el('span', { html: icons.chevron(14) }),
      el('span', {}, [`参照メール (${hits.length})`]),
    ]);
    const listEl = el('div', { class: 'tdr-sources collapsed' });

    hdrEl.addEventListener('click', () => {
      hdrEl.classList.toggle('collapsed');
      listEl.classList.toggle('collapsed');
    });

    hits.forEach((h, i) => {
      const plain = h.isHtml ? htmlToText(h.body) : h.body;
      const hitEl = el('div', { class: 'tdr-hit' });
      const head = el('div', { class: 'tdr-hit-head' }, [
        el('span', { class: 'tdr-hit-num' }, [String(i + 1)]),
        el('span', { class: 'tdr-hit-subject' }, [h.subject]),
        el('span', { class: 'tdr-hit-score' }, [h.score.toFixed(3)]),
      ]);
      if (h.internetMessageId) {
        const openBtn = el('button', {
          class: 'tdr-hit-open', 'aria-label': 'Outlook で開く', title: 'Outlook で開く',
          html: icons.external(14),
        });
        openBtn.addEventListener('click', async (e) => {
          e.stopPropagation();
          openBtn.disabled = true;
          try {
            await openMailInOutlook(relayBaseUrl, h.internetMessageId);
          } catch (err: unknown) {
            toast(root, `Outlook 表示に失敗: ${err instanceof Error ? err.message : String(err)}`, 'error');
          } finally {
            openBtn.disabled = false;
          }
        });
        head.appendChild(openBtn);
      }
      hitEl.appendChild(head);
      hitEl.appendChild(el('div', { class: 'tdr-hit-from' }, [
        `${h.from}  ${h.date.slice(0, 10)}`,
      ]));
      hitEl.appendChild(el('div', { class: 'tdr-hit-snippet' }, [
        plain.slice(0, 140) + (plain.length > 140 ? '…' : ''),
      ]));

      let detail: HTMLElement | null = null;
      hitEl.addEventListener('click', () => {
        hitEl.classList.toggle('is-open');
        if (!detail) {
          detail = el('div', { class: 'tdr-hit-detail' });
          renderMailBody(detail, h.body, h.isHtml); // HTML はサニタイズ描画、プレーンは pre-wrap
          hitEl.appendChild(detail);
        } else {
          detail.hidden = !detail.hidden;
        }
      });
      listEl.appendChild(hitEl);
    });

    container.append(hdrEl, listEl);
  }

  function scrollBottom(): void { thread.scrollTop = thread.scrollHeight; }

  function autosize(): void {
    input.style.height = 'auto';
    input.style.height = Math.min(input.scrollHeight, window.innerHeight * 0.4) + 'px';
  }

  input.addEventListener('input', autosize);
  input.addEventListener('keydown', (e: KeyboardEvent) => {
    if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
      e.preventDefault();
      void submit();
    }
  });
  sendBtn.addEventListener('click', () => void submit());

  // ── レイアウト: 左ペイン | ドラッグ可能な仕切り | チャット ──
  thread.appendChild(emptyState);
  refreshList();

  const sidebar = el('aside', { class: 'tdr-sidebar' }, [newBtn, sessionList]);
  const storedW = Number(localStorage.getItem(SIDEBAR_W_KEY) || '');
  sidebar.style.width = `${clampW(storedW || 260)}px`;

  const divider = el('div', { class: 'tdr-divider', 'aria-label': '幅を調整', title: 'ドラッグで幅を調整' });
  attachDividerDrag(divider, sidebar);

  const chatCol = el('div', { class: 'tdr-chat' }, [
    thread,
    el('div', { class: 'tdr-composer' }, [
      el('div', { class: 'tdr-composer-inner' }, [
        buildModelPicker(),
        el('div', { class: 'tdr-note-form' }, [input, sendBtn]),
        el('div', { class: 'tdr-note-hint' }, ['⌘+Enter または Ctrl+Enter で送信']),
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
