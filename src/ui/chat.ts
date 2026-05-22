// RAG チャットパネル。スレッド表示 + コンポーザ。
// 送信 → searchMails → generateAnswer (streaming) → 出典カード。

import { el } from '../lib/dom';
import { icons } from './icons';
import { toast } from './toast';
import { searchVectors } from '../search/vectorSearch';
import { generateAnswer, type RagSource } from '../rag/client';
import { loadSettings } from '../api/aiSettings';

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
  thread.appendChild(emptyState);

  let abort: AbortController | null = null;
  let emptyRemoved = false;

  async function submit(): Promise<void> {
    const q = input.value.trim();
    if (!q || sendBtn.disabled) return;
    input.value = '';
    autosize();
    sendBtn.disabled = true;

    if (!emptyRemoved) { emptyState.remove(); emptyRemoved = true; }

    // ユーザー吹き出し
    const turnEl = el('div', { class: 'tdr-turn' });
    turnEl.appendChild(el('div', { class: 'tdr-q' }, [q]));
    thread.appendChild(turnEl);
    scrollBottom();

    // AI 吹き出し (ローディング)
    const answerText = el('div', { class: 'tdr-answer' });
    const metaEl    = el('div', { class: 'tdr-a-meta' });
    const aBody     = el('div', { class: 'tdr-a-body' }, [metaEl, answerText]);
    turnEl.appendChild(el('div', { class: 'tdr-a' }, [
      el('div', { class: 'tdr-a-avatar' }, ['T']),
      aBody,
    ]));
    answerText.textContent = '検索中…';

    abort?.abort();
    abort = new AbortController();
    const s = loadSettings();

    try {
      const hits = await searchVectors(q, s, siteUrl, 5);
      if (hits.length === 0) {
        answerText.textContent = '該当するメールが見つかりませんでした。';
        return;
      }

      const sources: RagSource[] = hits.map((h, i) => ({
        n: i + 1, subject: h.subject, from: h.from, date: h.date, body: h.body,
      }));

      answerText.textContent = '';
      const t0 = performance.now();

      await generateAnswer(q, sources, s, delta => {
        answerText.textContent += delta;
        scrollBottom();
      }, abort.signal);

      // [n] を引用チップに変換
      answerText.innerHTML = (answerText.textContent ?? '').replace(
        /\[(\d+)\]/g,
        (_, n) => `<span class="cite" data-n="${n}">[${n}]</span>`,
      );

      const ms = Math.round(performance.now() - t0);
      metaEl.append(
        el('span', {}, [`${hits.length} 件参照`]),
        el('span', { class: 'mono' }, [`${ms} ms`]),
      );

      appendSources(aBody, sources, hits.map(h => h.score));

    } catch (err: unknown) {
      if (err instanceof Error && err.name === 'AbortError') return;
      answerText.textContent = 'エラーが発生しました。';
      toast(root, `エラー: ${err instanceof Error ? err.message : String(err)}`, 'error');
    } finally {
      sendBtn.disabled = false;
      abort = null;
    }
  }

  function appendSources(container: HTMLElement, sources: RagSource[], scores: number[]): void {
    const hdrEl  = el('div', { class: 'tdr-sources-h' }, [
      el('span', { html: icons.chevron(14) }),
      el('span', {}, [`参照メール (${sources.length})`]),
    ]);
    const listEl = el('div', { class: 'tdr-sources' });

    hdrEl.addEventListener('click', () => {
      hdrEl.classList.toggle('collapsed');
      listEl.classList.toggle('collapsed');
    });

    sources.forEach((src, i) => {
      const hitEl = el('div', { class: 'tdr-hit' });
      hitEl.appendChild(el('div', { class: 'tdr-hit-head' }, [
        el('span', { class: 'tdr-hit-num' }, [String(src.n)]),
        el('span', { class: 'tdr-hit-subject' }, [src.subject]),
        el('span', { class: 'tdr-hit-score' }, [scores[i].toFixed(3)]),
      ]));
      hitEl.appendChild(el('div', { class: 'tdr-hit-from' }, [
        `${src.from}  ${src.date.slice(0, 10)}`,
      ]));
      hitEl.appendChild(el('div', { class: 'tdr-hit-snippet' }, [
        src.body.slice(0, 140) + (src.body.length > 140 ? '…' : ''),
      ]));

      let detail: HTMLElement | null = null;
      hitEl.addEventListener('click', () => {
        hitEl.classList.toggle('is-open');
        if (!detail) {
          detail = el('div', { class: 'tdr-hit-detail' }, [src.body]);
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

  return el('div', { style: 'display:flex;flex-direction:column;flex:1;min-height:0;' }, [
    thread,
    el('div', { class: 'tdr-composer' }, [
      el('div', { class: 'tdr-note-form' }, [input, sendBtn]),
      el('div', { class: 'tdr-note-hint' }, ['⌘+Enter または Ctrl+Enter で送信']),
    ]),
  ]);
}
