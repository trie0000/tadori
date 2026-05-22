// 設定ハブ — Spira と同じ左ナビ + 右ペイン構成。
// AI 接続設定は spira:ai:* キーで Spira と共有。
// 開発者モード時のみ Claude API / Voyage 埋め込みの設定とテストデータ投入を表示。

import { el } from '../lib/dom';
import { icons } from './icons';
import { openModal } from './modal';
import { toast } from './toast';
import {
  loadSettings, saveSettings, parseAddressList,
  DEFAULT_VOYAGE_MODEL,
  CORP_AI_MODELS, CLAUDE_MODELS, EMBEDDING_MODELS,
  type RuntimeSettings, type Provider,
} from '../api/aiSettings';
import { isDeveloperMode, setDeveloperMode } from '../utils/devMode';
import { embedQueryFor } from '../embeddings/router';
import { seedTestData, SAMPLE_MAILS } from '../dev/seed';

type SectionId = 'ai' | 'ingest' | 'display' | 'diag' | 'dev' | 'about';

export function openSettingsHub(root: HTMLElement, siteUrl: string): void {
  const draft: RuntimeSettings = { ...loadSettings() };

  const nav  = el('div', { class: 'tdr-hub-nav' });
  const pane = el('div', { class: 'tdr-hub-pane' });

  const navItems: { id: SectionId; label: string; icon: string }[] = [
    { id: 'ai',      label: 'AI 接続',  icon: icons.settings() },
    { id: 'ingest',  label: '取り込み', icon: icons.activity() },
    { id: 'display', label: '表示',     icon: icons.moon()     },
    { id: 'diag',    label: '診断',     icon: icons.search()   },
    { id: 'dev',     label: '開発者',   icon: icons.activity() },
    { id: 'about',   label: 'About',    icon: icons.chevron()  },
  ];

  const navBtns = new Map<SectionId, HTMLElement>();
  for (const item of navItems) {
    const btn = el('div', { class: 'tdr-hub-navitem' }, [
      el('span', { html: item.icon }),
      el('span', {}, [item.label]),
    ]);
    btn.addEventListener('click', () => activate(item.id));
    navBtns.set(item.id, btn);
    nav.appendChild(btn);
  }

  function activate(id: SectionId): void {
    for (const [sid, btn] of navBtns) btn.classList.toggle('is-active', sid === id);
    pane.textContent = '';
    switch (id) {
      case 'ai':      buildAiPane(pane, draft); break;
      case 'ingest':  buildIngestPane(pane, draft); break;
      case 'display': buildDisplayPane(pane, root); break;
      case 'diag':    buildDiagPane(pane, draft, root); break;
      case 'dev':     buildDevPane(pane, draft, root, siteUrl); break;
      case 'about':   buildAboutPane(pane); break;
    }
  }

  const saveBtn = el('button', { class: 'tdr-btn tdr-btn--primary' }, ['保存']);
  saveBtn.addEventListener('click', () => {
    saveSettings(draft);
    toast(root, '設定を保存しました', 'info');
  });

  openModal({
    root,
    title: '設定',
    body:   el('div', { class: 'tdr-hub' }, [nav, pane]),
    footer: el('div', { class: 'tdr-modal-footer' }, [saveBtn]),
  });

  activate('ai');
}

// ─── 共通ヘルパ ───────────────────────────────────────────────────────────────

function mkInput(value: string, onchange: (v: string) => void, type = 'text'): HTMLInputElement {
  const inp = el('input', { class: 'tdr-input', type, value });
  inp.addEventListener('change', () => onchange(inp.value));
  return inp;
}

function mkRow(label: string, ctrl: HTMLElement, hint?: string): HTMLElement[] {
  const nodes: HTMLElement[] = [el('label', {}, [label]), ctrl];
  if (hint) nodes.push(el('p', { class: 'tdr-hint' }, [hint]));
  return nodes;
}

function mkSelect(options: { value: string; label: string }[], current: string, onchange: (v: string) => void): HTMLSelectElement {
  const sel = el('select', { class: 'tdr-input' }, options.map(o => {
    const opt = el('option', { value: o.value }, [o.label]);
    if (o.value === current) opt.setAttribute('selected', 'selected');
    return opt;
  }));
  sel.addEventListener('change', () => onchange(sel.value));
  return sel;
}

// ─── AI 接続 ──────────────────────────────────────────────────────────────────

function buildAiPane(pane: HTMLElement, draft: RuntimeSettings): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title' }, ['AI 接続']));
  const dev = isDeveloperMode();

  // dev OFF なのに claude が選ばれていたら corp に丸める
  if (!dev && draft.provider === 'claude') draft.provider = 'corp';

  // ── provider セレクタ ──
  const select = el('select', { class: 'tdr-input' });
  const optCorp = el('option', { value: 'corp' }, ['社内 AI (Azure OpenAI 互換)']);
  if (draft.provider === 'corp') optCorp.setAttribute('selected', 'selected');
  select.appendChild(optCorp);
  if (dev) {
    const optClaude = el('option', { value: 'claude' }, ['Claude (Anthropic) — 開発者モード']);
    if (draft.provider === 'claude') optClaude.setAttribute('selected', 'selected');
    select.appendChild(optClaude);
  }

  const provGrid = el('div', { class: 'tdr-fieldgrid' });
  provGrid.append(...mkRow('プロバイダ', select, dev ? '開発者モードでは Claude を選べます' : undefined));
  pane.appendChild(provGrid);

  pane.appendChild(el('p', { class: 'tdr-shared-note', style: 'margin:var(--s-5) 0' }, [
    '★ Spira と共有される設定です。どちらで変更しても両方のツールに反映されます。',
  ]));

  // ── corp ブロック (Spira と同じ deploy-prefix 方式) ──
  const corpModelOpts = CORP_AI_MODELS.map(m => ({ value: m.id, label: m.id }));
  const embModelOpts = EMBEDDING_MODELS.map(m => ({ value: m, label: m }));
  const overrideTa = el('textarea', { class: 'tdr-input', rows: '3', placeholder: '{"gpt-5":{"apiVersion":"2025-01-01-preview"}}' });
  overrideTa.value = draft.corpOverridesRaw;
  overrideTa.addEventListener('change', () => { draft.corpOverridesRaw = overrideTa.value; });

  const corpGrid = el('div', { class: 'tdr-fieldgrid' });
  corpGrid.append(
    ...mkRow('API キー', mkInput(draft.apiKey, v => { draft.apiKey = v; }, 'password'), 'Azure OpenAI 互換 API キー'),
    ...mkRow('ベース URL', mkInput(draft.relayBaseUrl, v => { draft.relayBaseUrl = v; }), '中継サーバ / ゲートウェイ (例: http://localhost:18080)'),
    ...mkRow('デプロイ prefix', mkInput(draft.corpDeployPrefix, v => { draft.corpDeployPrefix = v; }), 'デプロイ名 = <prefix><モデル名(.除去)>'),
    ...mkRow('チャットモデル', mkSelect(corpModelOpts, draft.chatModel, v => { draft.chatModel = v; }), 'RAG 回答生成に使うモデル'),
    ...mkRow('埋め込みモデル', mkSelect(embModelOpts, draft.embeddingModel, v => { draft.embeddingModel = v; }), '検索ベクトル生成 (Tadori 固有)'),
    ...mkRow('埋め込み API バージョン', mkInput(draft.apiVersion, v => { draft.apiVersion = v; }), '例: 2024-02-01'),
    ...mkRow('次元数', mkInput(String(draft.dimensions), v => { draft.dimensions = Number(v) || 256; }), 'Matryoshka 短縮次元数 (256)'),
    el('label', { class: 'top' }, ['オーバーライド']),
    overrideTa,
    el('p', { class: 'tdr-hint' }, ['モデル毎に baseUrl/apiVersion/deploymentId を上書き (JSON, 任意)']),
  );
  const corpBlock = el('div', {}, [corpGrid]);

  // ── claude / voyage ブロック (dev のみ) ──
  const claudeBlock = el('div', {});
  if (dev) {
    const cGrid = el('div', { class: 'tdr-fieldgrid' });
    cGrid.append(
      el('p', { class: 'tdr-hint', style: 'grid-column:1/-1;margin-bottom:var(--s-2)' }, ['── Claude (回答生成) ──']),
      ...mkRow('Claude API キー', mkInput(draft.claudeApiKey, v => { draft.claudeApiKey = v; }, 'password'), 'sk-ant-... (この端末のみ保存)'),
      ...mkRow('Claude モデル', mkSelect(CLAUDE_MODELS.map(m => ({ value: m.id, label: m.label })), draft.claudeModel, v => { draft.claudeModel = v; })),
      el('p', { class: 'tdr-hint', style: 'grid-column:1/-1;margin:var(--s-3) 0 var(--s-2)' }, ['── Voyage (検索埋め込み) ──']),
      ...mkRow('Voyage API キー', mkInput(draft.voyageApiKey, v => { draft.voyageApiKey = v; }, 'password'), 'pa-... Claude に埋め込み API が無いため'),
      ...mkRow('Voyage モデル', mkInput(draft.voyageModel, v => { draft.voyageModel = v; }), `既定: ${DEFAULT_VOYAGE_MODEL}`),
      ...mkRow('次元数', mkInput(String(draft.dimensions), v => { draft.dimensions = Number(v) || 256; }), 'output_dimension (256/512/1024)'),
    );
    claudeBlock.appendChild(cGrid);
  }

  pane.append(corpBlock, claudeBlock);

  function sync(): void {
    const p = select.value as Provider;
    corpBlock.style.display   = p === 'corp' ? '' : 'none';
    claudeBlock.style.display = p === 'claude' ? '' : 'none';
  }
  select.addEventListener('change', () => { draft.provider = select.value as Provider; sync(); });
  sync();
}

// ─── 取り込み ─────────────────────────────────────────────────────────────────

function buildIngestPane(pane: HTMLElement, draft: RuntimeSettings): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title' }, ['取り込み']));

  const grid = el('div', { class: 'tdr-fieldgrid' });
  const addrArea = el('textarea', { class: 'tdr-input', rows: '4' });
  addrArea.value = draft.mlAddresses.join('\n');
  addrArea.addEventListener('change', () => { draft.mlAddresses = parseAddressList(addrArea.value); });

  grid.append(
    ...mkRow('List 表示名', mkInput(draft.listTitle, v => { draft.listTitle = v; }), '例: 受信メールリスト'),
    el('label', { class: 'top' }, ['ML アドレス']),
    addrArea,
    el('p', { class: 'tdr-hint' }, ['取り込み対象アドレス。1 行に 1 件。']),
    ...mkRow('取り込み間隔 (秒)', mkInput(String(draft.ingestIntervalSec), v => { draft.ingestIntervalSec = Number(v) || 30; }), 'デフォルト 30 秒'),
  );
  pane.appendChild(grid);
}

// ─── 表示 ─────────────────────────────────────────────────────────────────────

function buildDisplayPane(pane: HTMLElement, root: HTMLElement): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title' }, ['表示']));

  const toggleBtn = el('button', { class: 'tdr-btn' }, [
    el('span', { html: icons.moon() }),
    el('span', {}, [root.dataset.theme === 'dark' ? 'ダークモード: ON' : 'ダークモード: OFF']),
  ]);
  toggleBtn.addEventListener('click', () => {
    const isDark = root.dataset.theme === 'dark';
    root.dataset.theme = isDark ? '' : 'dark';
    localStorage.setItem('tadori:theme', isDark ? '' : 'dark');
    const lbl = toggleBtn.querySelector('span:last-child');
    if (lbl) lbl.textContent = isDark ? 'ダークモード: OFF' : 'ダークモード: ON';
  });
  pane.appendChild(el('div', { style: 'margin-top:var(--s-3)' }, [toggleBtn]));
}

// ─── 診断 ─────────────────────────────────────────────────────────────────────

function buildDiagPane(pane: HTMLElement, draft: RuntimeSettings, root: HTMLElement): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title' }, ['診断']));

  function mkDiagRow(label: string): { row: HTMLElement; set: (ok: boolean, text: string) => void } {
    const stat = el('span', { class: 'stat' }, ['—']);
    const row  = el('div', { class: 'tdr-diag' }, [el('span', {}, [label]), stat]);
    return {
      row,
      set(ok: boolean, text: string): void {
        stat.textContent = text;
        stat.className   = `stat ${ok ? 'ok' : 'ng'}`;
      },
    };
  }

  const embed = mkDiagRow(draft.provider === 'claude' ? '埋め込み (Voyage)' : '埋め込み (中継)');
  const runBtn = el('button', { class: 'tdr-btn tdr-btn--primary', style: 'margin-top:var(--s-6)' }, ['埋め込みテスト']);

  runBtn.addEventListener('click', () => {
    runBtn.disabled = true;
    embed.set(false, '確認中…');
    void (async () => {
      try {
        const vec = await embedQueryFor('テスト', draft);
        embed.set(true, `OK — dim: ${vec.length}`);
      } catch (e) {
        embed.set(false, e instanceof Error ? e.message.slice(0, 80) : 'failed');
        toast(root, `埋め込みテスト失敗: ${e instanceof Error ? e.message : ''}`, 'error');
      }
      runBtn.disabled = false;
    })();
  });

  pane.append(embed.row, runBtn);
}

// ─── 開発者 ─────────────────────────────────────────────────────────────────────

function buildDevPane(
  pane: HTMLElement,
  draft: RuntimeSettings,
  root: HTMLElement,
  siteUrl: string,
): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title' }, ['開発者']));

  // 開発者モードトグル (即時保存)
  const checkbox = el('input', { type: 'checkbox' });
  if (isDeveloperMode()) checkbox.checked = true;
  checkbox.addEventListener('change', () => {
    setDeveloperMode(checkbox.checked);
    toast(root, checkbox.checked ? '開発者モード ON' : '開発者モード OFF', 'info');
  });
  pane.appendChild(el('label', {
    style: 'display:inline-flex;align-items:center;gap:var(--s-3);cursor:pointer;padding:var(--s-3);background:var(--paper-2);border-radius:var(--r-2)',
  }, [checkbox, el('span', { style: 'font-size:var(--fs-md)' }, ['開発者モードを有効にする (Claude API / テスト投入)'])]));

  pane.appendChild(el('p', { class: 'tdr-hint', style: 'margin-top:var(--s-3)' }, [
    '※ 端末ローカル (localStorage) に保存。AI 接続で Claude を選べるようになります。',
  ]));

  // テストデータ投入
  pane.appendChild(el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-7)' }, [`テストデータ投入 (${SAMPLE_MAILS.length} 件)`]));
  pane.appendChild(el('p', { class: 'tdr-hint', style: 'margin-bottom:var(--s-4)' }, [
    `現在の provider で本文を埋め込み、List「${draft.listTitle}」へサンプルメールを作成します。`,
  ]));

  const status = el('div', { style: 'font-size:var(--fs-sm);color:var(--ink-3);margin-top:var(--s-3)' }, ['']);
  const seedBtn = el('button', { class: 'tdr-btn tdr-btn--primary' }, ['サンプルメールを投入']);

  seedBtn.addEventListener('click', () => {
    if (!confirm(`List「${draft.listTitle}」に ${SAMPLE_MAILS.length} 件のテストメールを作成します。よろしいですか?`)) return;
    seedBtn.disabled = true;
    status.textContent = '埋め込み中…';
    void (async () => {
      try {
        const r = await seedTestData(draft, siteUrl, (done, total) => {
          status.textContent = `投入中… ${done}/${total}`;
        });
        status.textContent = `完了: ${r.created} 件作成` + (r.errors.length ? ` / ${r.errors.length} 件失敗` : '');
        toast(root, `テストデータ ${r.created} 件を投入しました`, r.errors.length ? 'error' : 'info');
        if (r.errors.length) console.warn('[tadori/seed] errors:', r.errors);
      } catch (e) {
        status.textContent = '失敗';
        toast(root, `投入失敗: ${e instanceof Error ? e.message : String(e)}`, 'error');
      }
      seedBtn.disabled = false;
    })();
  });

  pane.append(seedBtn, status);
}

// ─── About ────────────────────────────────────────────────────────────────────

function buildAboutPane(pane: HTMLElement): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title' }, ['About']));
  const grid = el('div', { class: 'tdr-fieldgrid' });
  grid.append(
    el('label', {}, ['バージョン']),
    el('span', { style: 'font-family:var(--font-mono);font-size:var(--fs-sm)' }, [__TADORI_VERSION__]),
    el('label', {}, ['ビルド']),
    el('span', { style: 'font-family:var(--font-mono);font-size:var(--fs-xs);color:var(--ink-3)' }, [__TADORI_BUILD_ID__]),
  );
  pane.appendChild(grid);
}
