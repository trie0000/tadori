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
import {
  getBundleSource, setBundleSource, getLocalBase, setLocalBase,
  getRelayBundleDir, setRelayBundleDir, DEFAULT_LOCAL_BASE, type BundleSource,
} from '../utils/bundleSource';
import { embedQueryFor } from '../embeddings/router';
import { seedTestData, SAMPLE_MAILS } from '../dev/seed';
import { fetchOutlookMails, toIngestMails } from '../outlook/import';
import { ingestToSegments } from '../db/writer';
import { getEngine } from '../db/engine';
import { fetchMonthlyTotals, currentUser } from '../usage/tracker';

type SectionId = 'ai' | 'search' | 'ingest' | 'diag' | 'usage' | 'display' | 'dev';

// メニュー構成は固定 (Spira と同じグループ流儀)。むやみに名前を変えないこと。
const NAV_GROUPS: { title: string; items: { id: SectionId; label: string }[] }[] = [
  { title: '表示', items: [
    { id: 'display', label: '表示' },
  ] },
  { title: 'AI / 自動化', items: [
    { id: 'ai',     label: 'AI 設定' },
    { id: 'search', label: '検索' },
    { id: 'ingest', label: '取り込み' },
    { id: 'diag',   label: '診断' },
    { id: 'usage',  label: '利用料' },
  ] },
  { title: '運用', items: [
    { id: 'dev', label: '開発者モード' },
  ] },
];

export function openSettingsHub(root: HTMLElement, siteUrl: string): void {
  const draft: RuntimeSettings = { ...loadSettings() };

  const nav  = el('div', { class: 'tdr-hub-nav' });
  const pane = el('div', { class: 'tdr-hub-pane' });

  const navBtns = new Map<SectionId, HTMLElement>();
  for (const g of NAV_GROUPS) {
    nav.appendChild(el('div', { class: 'tdr-hub-group' }, [g.title]));
    for (const item of g.items) {
      const btn = el('div', { class: 'tdr-hub-navitem' }, [item.label]);
      btn.addEventListener('click', () => activate(item.id));
      navBtns.set(item.id, btn);
      nav.appendChild(btn);
    }
  }

  function activate(id: SectionId): void {
    for (const [sid, btn] of navBtns) btn.classList.toggle('is-active', sid === id);
    pane.textContent = '';
    switch (id) {
      case 'ai':      buildAiPane(pane, draft); break;
      case 'search':  buildSearchPane(pane, draft); break;
      case 'ingest':  buildIngestPane(pane, draft, root, siteUrl); break;
      case 'display': buildDisplayPane(pane, root); break;
      case 'diag':    buildDiagPane(pane, draft, root, siteUrl); break;
      case 'usage':   buildUsagePane(pane, root); break;
      case 'dev':     buildDevPane(pane, draft, root, siteUrl); break;
    }
  }

  const saveBtn = el('button', { class: 'tdr-btn tdr-btn--primary' }, ['保存']);
  saveBtn.addEventListener('click', () => {
    saveSettings(draft);
    toast(root, '設定を保存しました', 'ok');
  });

  openModal({
    root,
    title: '設定',
    large: true,
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

/** ペイン共通の見出し (タイトル + プレーンな説明文。枠は付けない)。 */
function paneHead(pane: HTMLElement, title: string, desc?: string): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title' }, [title]));
  if (desc) pane.appendChild(el('p', { class: 'tdr-hint', style: 'margin:0 0 var(--s-5)' }, [desc]));
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
  paneHead(pane, 'AI 設定', '★ Spira と共有される設定です。どちらで変更しても両方のツールに反映されます。');
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

  const provGrid = el('div', { class: 'tdr-fieldgrid', style: 'margin-bottom:var(--s-5)' });
  provGrid.append(...mkRow('プロバイダ', select, dev ? '開発者モードでは Claude を選べます' : undefined));
  pane.appendChild(provGrid);

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

// ─── 検索 ─────────────────────────────────────────────────────────────────────

function buildSearchPane(pane: HTMLElement, draft: RuntimeSettings): void {
  paneHead(pane, '検索', '回答の根拠にする参照メールの取り方を調整します。多いほど網羅的ですが、プロンプトが長くなり利用料も増えます。');

  const grid = el('div', { class: 'tdr-fieldgrid' });
  grid.append(
    ...mkRow('参照件数 (1〜20)', mkInput(String(draft.ragTopK), v => {
      draft.ragTopK = Math.min(20, Math.max(1, Number(v) || 8));
    }), '上位何件のメールを根拠にするか。デフォルト 8'),
    ...mkRow('最小スコア (0〜1)', mkInput(String(draft.ragMinScore), v => {
      const n = Number(v); draft.ragMinScore = isNaN(n) ? 0.3 : Math.min(1, Math.max(0, n));
    }), 'cosine 類似度がこの値未満のメールは除外。デフォルト 0.3 (0 で無効)'),
  );
  pane.appendChild(grid);
}

// ─── 取り込み ─────────────────────────────────────────────────────────────────

function buildIngestPane(pane: HTMLElement, draft: RuntimeSettings, root: HTMLElement, siteUrl: string): void {
  paneHead(pane, '取り込み', '取り込み対象のメーリングリストと、ベクトルを格納する SharePoint List を設定します。');

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
    ...mkRow('埋め込み並列数 (1〜10)', mkInput(String(draft.embedConcurrency), v => { draft.embedConcurrency = Math.min(10, Math.max(1, Number(v) || 3)); }), '大きいほど取り込みが速いが API 負荷増。デフォルト 3'),
  );
  pane.appendChild(grid);

  // ── Outlook からの既存メールインポート ──
  buildOutlookImport(pane, draft, root, siteUrl);
}

function buildOutlookImport(pane: HTMLElement, draft: RuntimeSettings, root: HTMLElement, siteUrl: string): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-8)' }, ['Outlook からインポート']));
  pane.appendChild(el('p', { class: 'tdr-hint', style: 'margin:0 0 var(--s-4)' }, [
    'ローカル中継サーバ経由で Outlook の既存メールを読み込み、To/Cc 条件と受信期間で絞って取り込みます (中継サーバの起動が必要)。',
  ]));

  const today = new Date();
  const yearAgo = new Date(today.getTime() - 365 * 86400000);
  const iso = (d: Date) => d.toISOString().slice(0, 10);

  const toArea = el('textarea', { class: 'tdr-input', rows: '2' });
  toArea.value = draft.mlAddresses.join('\n');
  const ccArea = el('textarea', { class: 'tdr-input', rows: '2' });
  const sinceInp = el('input', { class: 'tdr-input', type: 'date', value: iso(yearAgo) });
  const untilInp = el('input', { class: 'tdr-input', type: 'date', value: iso(today) });
  const maxInp = el('input', { class: 'tdr-input', type: 'number', value: '1000' });

  const grid = el('div', { class: 'tdr-fieldgrid' });
  grid.append(
    el('label', { class: 'top' }, ['To (宛先)']),
    toArea,
    el('p', { class: 'tdr-hint' }, ['この宛先のメールを取り込む。1 行に 1 件 (既定: ML アドレス)。']),
    el('label', { class: 'top' }, ['Cc']),
    ccArea,
    el('p', { class: 'tdr-hint' }, ['Cc にこのアドレスが入るメールも対象。任意。']),
    ...mkRow('期間 (開始)', sinceInp),
    ...mkRow('期間 (終了)', untilInp),
    ...mkRow('最大件数', maxInp),
  );
  pane.appendChild(grid);

  const status = el('div', { style: 'font-size:var(--fs-sm);color:var(--ink-3);margin-top:var(--s-3)' }, ['']);
  const barFill = el('div', { class: 'tdr-progress-fill' });
  const bar = el('div', { class: 'tdr-progress', style: 'display:none' }, [barFill]);
  const btn = el('button', { class: 'tdr-btn tdr-btn--primary', style: 'margin-top:var(--s-4)' }, ['Outlook から取り込む']);

  function showBar(pct: number): void { bar.style.display = ''; barFill.style.width = `${pct}%`; }
  function hideBar(): void { bar.style.display = 'none'; barFill.style.width = '0%'; }

  const RUN_LABEL = 'Outlook から取り込む';
  let ac: AbortController | null = null;

  btn.addEventListener('click', () => {
    // 実行中にもう一度押す = 停止
    if (ac) { ac.abort(); btn.textContent = '停止中…'; return; }

    const filter = {
      to: parseAddressList(toArea.value),
      cc: parseAddressList(ccArea.value),
      since: sinceInp.value || undefined,
      until: untilInp.value || undefined,
      max: Number(maxInp.value) || 1000,
    };
    ac = new AbortController();
    const signal = ac.signal;
    btn.textContent = '停止';
    status.textContent = 'Outlook を検索中…';
    hideBar();
    void (async () => {
      try {
        // ① まず対象件数を取得して提示
        const mails = await fetchOutlookMails(draft.relayBaseUrl, filter, signal);
        if (mails.length === 0) {
          status.textContent = '該当するメールがありませんでした (条件・期間を確認)';
          toast(root, '該当メール 0 件', 'warn');
          return;
        }
        status.textContent = `対象 ${mails.length} 件が見つかりました`;
        if (!confirm(`Outlook から ${mails.length} 件取り込みます。よろしいですか? (実行中は「停止」で中断可)`)) {
          status.textContent = `キャンセル (対象 ${mails.length} 件)`;
          return;
        }
        // ② 埋め込み → 投入 (進捗バー + 件数。埋め込みと保存を 1 本のバーで単調増加)。
        let embedded = 0, saved = 0;
        const r = await ingestToSegments(toIngestMails(mails), draft, siteUrl, (phase, done, total) => {
          if (phase === 'sync') { status.textContent = '準備中…'; return; }
          if (phase === 'embed') embedded = done;
          if (phase === 'upload') saved = done;
          const units = (total || mails.length) * 2 || 1;
          const pct = Math.min(100, Math.round((embedded + saved) / units * 100));
          showBar(pct);
          status.textContent = `埋め込み ${embedded}/${total} ・ 保存 ${saved}/${total} 件 (${pct}%)`;
        }, signal);
        hideBar();
        const dup = r.skipped ? ` / 重複スキップ ${r.skipped} 件` : '';
        if (r.cancelled) {
          status.textContent = `停止しました: 保存済み 新規 ${r.added} 件 (セグメント ${r.segments})${dup}。再実行で続きから取り込めます。`;
          toast(root, `停止 (新規 ${r.added} 件は保存済み)`, 'warn');
        } else if (r.added === 0) {
          // 新規 0 件 = 取得分はすべて登録済み。失敗ではない旨を明示。
          status.textContent = `すべて登録済みでした (取得 ${mails.length} 件 / 新規 0 件)。SharePoint への追記はありません。`;
          toast(root, `新規なし (${mails.length} 件はすべて登録済み)`, 'warn');
        } else {
          status.textContent = `完了: 取得 ${mails.length} 件 / 新規 ${r.added} 件 (セグメント ${r.segments})${dup}`;
          toast(root, `Outlook から ${r.added} 件取り込みました`, 'ok');
        }
      } catch (e) {
        if (signal.aborted || (e instanceof Error && e.name === 'AbortError')) {
          status.textContent = '停止しました';
        } else {
          status.textContent = '失敗';
          toast(root, `インポート失敗: ${e instanceof Error ? e.message : String(e)}`, 'error');
        }
      } finally {
        ac = null;
        btn.textContent = RUN_LABEL;
        hideBar();
      }
    })();
  });

  pane.append(btn, bar, status);
}

// ─── 表示 ─────────────────────────────────────────────────────────────────────

function buildDisplayPane(pane: HTMLElement, root: HTMLElement): void {
  paneHead(pane, '表示', '外観の設定。テーマはこの端末にのみ保存されます。');

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

function buildDiagPane(pane: HTMLElement, draft: RuntimeSettings, root: HTMLElement, siteUrl: string): void {
  paneHead(pane, '診断', '現在の provider で埋め込み API に接続できるか確認します。');

  // 登録済みベクトル件数 (SharePoint から同期したローカルDBの実件数)。
  const countRow = el('div', { class: 'tdr-diag' }, [
    el('span', {}, ['登録済みベクトル件数']),
    el('span', { class: 'stat' }, ['確認中…']),
  ]);
  pane.appendChild(countRow);
  void (async () => {
    const stat = countRow.querySelector('.stat') as HTMLElement;
    try {
      const eng = await getEngine(siteUrl);
      await eng.sync.sync();
      stat.textContent = `${eng.db.size} 件`;
      stat.className = 'stat ok';
    } catch (e) {
      stat.textContent = e instanceof Error ? e.message.slice(0, 60) : '取得失敗';
      stat.className = 'stat ng';
    }
  })();

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

// ─── 利用料 ─────────────────────────────────────────────────────────────────────

function fmtYen(n: number): string {
  return (Math.round(n * 100) / 100).toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function buildUsagePane(pane: HTMLElement, root: HTMLElement): void {
  paneHead(pane, '利用料', 'AI (チャット・埋め込み) 利用料金の目安です。料金表から概算し、毎月 1 日にリセットされます。実請求は為替・実トークン数で変動します。');

  pane.appendChild(el('p', { class: 'tdr-hint', style: 'margin-top:var(--s-5)' }, ['今月の利用料 (利用者全員の合計)']));
  const totalEl = el('div', { style: 'font-size:var(--fs-xl);font-weight:700;color:var(--ink);margin-top:var(--s-2)' }, ['—']);
  const subEl   = el('div', { class: 'tdr-hint', style: 'margin-top:var(--s-2)' }, ['']);
  const breakdown = el('div', { style: 'margin-top:var(--s-6)' });
  const refresh = el('button', { class: 'tdr-btn', style: 'margin-top:var(--s-6)' }, [el('span', { html: icons.activity(14) }), '更新']);

  async function load(): Promise<void> {
    totalEl.textContent = '集計中…';
    subEl.textContent = '';
    breakdown.replaceChildren();
    try {
      const t = await fetchMonthlyTotals();
      totalEl.textContent = `¥${fmtYen(t.total)}`;
      subEl.textContent = `${t.month} ・ あなたの利用分 ¥${fmtYen(t.ownYen)} (${t.ownTokens.toLocaleString()} トークン)`;
      if (isDeveloperMode()) {
        breakdown.appendChild(el('p', { class: 'tdr-pane-title' }, ['ユーザー別内訳 (開発者モード)']));
        if (t.byUser.length === 0) {
          breakdown.appendChild(el('p', { class: 'tdr-hint' }, ['記録がありません']));
        } else {
          const me = currentUser();
          for (const u of t.byUser) {
            breakdown.appendChild(el('div', { class: 'tdr-diag' }, [
              el('span', {}, [u.user === me ? `${u.user} (あなた)` : u.user]),
              el('span', { class: 'stat ok' }, [`¥${fmtYen(u.yen)} ・ ${u.tokens.toLocaleString()} tok`]),
            ]));
          }
        }
      }
    } catch (e) {
      totalEl.textContent = '取得失敗';
      toast(root, `利用料の取得に失敗: ${e instanceof Error ? e.message : ''}`, 'error');
    }
  }
  refresh.addEventListener('click', () => void load());

  pane.append(totalEl, subEl, breakdown, refresh);
  void load();
}

// ─── 開発者 ─────────────────────────────────────────────────────────────────────

function buildDevPane(
  pane: HTMLElement,
  draft: RuntimeSettings,
  root: HTMLElement,
  siteUrl: string,
): void {
  paneHead(pane, '開発者', '実験的機能 (Claude API 直接利用・テストデータ投入) を有効化します。通常運用では OFF のまま。');

  // 開発者モードトグル (即時保存)
  const checkbox = el('input', { type: 'checkbox' });
  if (isDeveloperMode()) checkbox.checked = true;
  checkbox.addEventListener('change', () => {
    setDeveloperMode(checkbox.checked);
    toast(root, checkbox.checked ? '開発者モード ON' : '開発者モード OFF', 'ok');
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
    '現在の provider で本文を埋め込み、ベクトルDB (SharePoint のセグメント) へサンプルメールを投入します。',
  ]));

  const status = el('div', { style: 'font-size:var(--fs-sm);color:var(--ink-3);margin-top:var(--s-3)' }, ['']);
  const seedBtn = el('button', { class: 'tdr-btn tdr-btn--primary' }, ['サンプルメールを投入']);

  seedBtn.addEventListener('click', () => {
    if (!confirm(`ベクトルDBに ${SAMPLE_MAILS.length} 件のサンプルメールを投入します。よろしいですか?`)) return;
    seedBtn.disabled = true;
    status.textContent = '投入中…';
    void (async () => {
      try {
        const r = await seedTestData(draft, siteUrl, (done, total) => {
          status.textContent = total ? `投入中… ${done}/${total}` : '投入中…';
        });
        const dup = r.skipped ? ` / 重複スキップ ${r.skipped} 件` : '';
        status.textContent = `完了: 新規 ${r.added} 件 (セグメント ${r.segments})${dup}`;
        toast(root, `サンプル ${r.added} 件を投入しました`, 'ok');
      } catch (e) {
        status.textContent = '失敗';
        toast(root, `投入失敗: ${e instanceof Error ? e.message : String(e)}`, 'error');
      }
      seedBtn.disabled = false;
    })();
  });

  pane.append(seedBtn, status);

  buildBundleSourcePane(pane, root);
}

// 開発: バンドル読み込み元 (SharePoint / ローカル relay) の切替。変更は即保存・
// 次回リロードで反映。
function buildBundleSourcePane(pane: HTMLElement, root: HTMLElement): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-8)' }, ['バンドル読み込み元']));
  pane.appendChild(el('p', { class: 'tdr-hint', style: 'margin:0 0 var(--s-4)' }, [
    'Tadori 本体をどこから読むか。テスト時にローカル relay の dist を読ませる用途。次回リロードで反映。',
  ]));

  const sel = el('select', { class: 'tdr-input' });
  const optSP = el('option', { value: 'sharepoint' }, ['SharePoint (本番)']);
  const optLocal = el('option', { value: 'local' }, ['ローカル relay (開発)']);
  if (getBundleSource() === 'local') optLocal.setAttribute('selected', 'selected');
  else optSP.setAttribute('selected', 'selected');
  sel.append(optSP, optLocal);
  sel.addEventListener('change', () => {
    setBundleSource(sel.value as BundleSource);
    toast(root, `読み込み元を ${sel.value === 'local' ? 'ローカル relay' : 'SharePoint'} に。リロードで反映`, 'ok');
  });

  const baseInp = mkInput(getLocalBase(), v => setLocalBase(v));
  baseInp.placeholder = DEFAULT_LOCAL_BASE;

  const dirInp = mkInput('', v => {
    const dir = v.trim();
    if (!dir) return;
    void setRelayBundleDir(dir).then(r => {
      if (r) dirStatus.textContent = `現在: ${r.dir}  ${r.hasBundle ? '✅ tadori.bundle.js あり' : '⚠ 無い'}`;
      else toast(root, 'relay へのフォルダ設定に失敗 (relay 未起動?)', 'warn');
    });
  });
  dirInp.placeholder = 'C:\\tools\\tadori\\dist';
  const dirStatus = el('p', { class: 'tdr-hint' }, ['relay に照会中…']);
  void getRelayBundleDir().then(r => {
    if (!r) { dirStatus.textContent = '⚠ relay 未起動 / 応答なし'; return; }
    if (!dirInp.value) dirInp.value = r.dir;
    dirStatus.textContent = `現在: ${r.dir}  ${r.hasBundle ? '✅ tadori.bundle.js あり' : '⚠ tadori.bundle.js が無い'}`;
  });

  const grid = el('div', { class: 'tdr-fieldgrid' });
  grid.append(
    ...mkRow('読み込み元', sel),
    ...mkRow('ローカル base', baseInp, '例: http://127.0.0.1:18080/tadori'),
    el('label', { class: 'top' }, ['relay 配信フォルダ']),
    dirInp,
    dirStatus,
  );
  pane.appendChild(grid);
}

