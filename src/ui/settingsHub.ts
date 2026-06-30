// 設定ハブ — Spira と同じ左ナビ + 右ペイン構成 (UI の見た目だけ流用)。
// AI 接続設定は tadori:ai:* キーで Tadori 専用に独立管理 (Spira とは別)。
// 開発者モード時のみ Claude API / Voyage 埋め込みの設定とテストデータ投入を表示。

import { el } from '../lib/dom';
import { icons } from './icons';
import { openModal, confirmModal } from './modal';
import { toast } from './toast';
import {
  loadSettings, saveSettings, parseAddressList,
  DEFAULT_VOYAGE_MODEL,
  CORP_AI_MODELS, CLAUDE_MODELS, EMBEDDING_MODELS,
  type RuntimeSettings, type Provider,
} from '../api/aiSettings';
import { isDeveloperMode, setDeveloperMode, isAutoIngestFlagOn, setAutoIngestFlag } from '../utils/devMode';
import { type GlossaryEntry, loadGlossary, fetchGlossary, persistGlossary, parseGlossaryTable } from '../search/glossary';
import {
  getBundleSource, setBundleSource, getLocalBase, setLocalBase,
  getRelayBundleDir, setRelayBundleDir, DEFAULT_LOCAL_BASE, type BundleSource,
} from '../utils/bundleSource';
import { embedQueryFor } from '../embeddings/router';
import { seedTestData, SAMPLE_MAILS } from '../dev/seed';
import { fetchOutlookMails, toIngestMails } from '../outlook/import';
import { fetchOneNoteHierarchy, fetchOneNotePages, pagesToIngestMails, type OneNoteNotebook } from '../onenote/import';
import { recordOneNoteBatch, removeOneNoteBatch, renameOneNoteBatch, listOneNoteBatches, setOneNoteBatchPageIds, type OneNoteBatch } from '../sync/onenoteSources';
import { getExcludedOneNotePageIds, setExcludedOneNotePageIds } from '../onenote/exclude';
import { ingestToSegments } from '../db/writer';
import { type PptxFolderConfig } from '../sync/pptxFolders';
import { syncPptxFolder } from '../sync/pptxIngest';
import {
  listTranscriptFolders, addTranscriptFolder, removeTranscriptFolder,
  deriveLabel as deriveTranscriptLabel, type TranscriptFolderConfig,
} from '../sync/transcriptFolders';
import { syncTranscriptFolder, type TranscriptIngestProgress } from '../sync/transcriptIngest';
import {
  listDocFolders, addDocFolder, removeDocFolder, updateDocFolderPptxSync,
  deriveLabel as deriveDocLabel, type DocFolderConfig,
} from '../sync/docFolders';
import { syncDocFolder, type DocIngestProgress } from '../sync/docIngest';
import { getEngine, wipeImportedMails } from '../db/engine';
import { getFontSize, setFontSize } from '../utils/fontSize';
import {
  loadRules as loadExclusionRules, addRule as addExclusionRule, deleteRule as deleteExclusionRule,
  updateRule as updateExclusionRule, FIELD_LABELS as EXCLUDE_FIELD_LABELS,
  type ExclusionField, type ExclusionRule,
} from '../search/exclusionRules';
import { fetchMonthlyTotals, currentUser } from '../usage/tracker';

type SectionId = 'ai' | 'search' | 'glossary' | 'exclude' | 'ingest' | 'diag' | 'usage' | 'display' | 'dev' | 'paSetup' | 'about' | 'resetMail' | 'resetAll';

// メニュー構成は固定 (Spira と同じグループ流儀)。むやみに名前を変えないこと。
const NAV_GROUPS: { title: string; items: { id: SectionId; label: string }[] }[] = [
  { title: '表示', items: [
    { id: 'display', label: '表示' },
  ] },
  { title: 'AI / 自動化', items: [
    { id: 'ai',     label: 'AI 設定' },
    { id: 'search', label: '検索' },
    { id: 'glossary', label: '用語辞書' },
    { id: 'exclude',label: '除外ルール' },
    { id: 'ingest', label: '取り込み' },
    { id: 'paSetup',label: 'PA セットアップ' },
    { id: 'diag',   label: '診断' },
    { id: 'usage',  label: '利用料' },
  ] },
  { title: '運用', items: [
    { id: 'dev',   label: '開発者モード' },
    { id: 'about', label: 'Tadori について' },
  ] },
  { title: '危険ゾーン', items: [
    { id: 'resetMail', label: '取り込みメールを全削除' },
    { id: 'resetAll',  label: 'ツール全体をリセット' },
  ] },
];

export function openSettingsHub(root: HTMLElement, siteUrl: string): void {
  const draft: RuntimeSettings = { ...loadSettings() };

  const nav  = el('div', { class: 'tdr-hub-nav' });
  const pane = el('div', { class: 'tdr-hub-pane' });

  const navBtns = new Map<SectionId, HTMLElement>();
  for (const g of NAV_GROUPS) {
    const danger = g.title === '危険ゾーン';
    nav.appendChild(el('div', { class: 'tdr-hub-group' + (danger ? ' is-danger' : '') }, [g.title]));
    for (const item of g.items) {
      const btn = el('div', { class: 'tdr-hub-navitem' + (danger ? ' is-danger' : '') }, [item.label]);
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
      case 'glossary': buildGlossaryPane(pane, root, siteUrl); break;
      case 'exclude': buildExcludePane(pane, root); break;
      case 'ingest':  buildIngestPane(pane, draft, root, siteUrl); break;
      case 'paSetup': buildPaSetupPane(pane, draft, root, siteUrl); break;
      case 'about':   buildAboutPane(pane, root); break;
      case 'display': buildDisplayPane(pane, root); break;
      case 'diag':    buildDiagPane(pane, draft, root, siteUrl); break;
      case 'usage':   buildUsagePane(pane, root); break;
      case 'dev':     buildDevPane(pane, draft, root, siteUrl); break;
      case 'resetMail': buildResetMailPane(pane, root, siteUrl); break;
      case 'resetAll':  buildResetAllPane(pane, root, siteUrl); break;
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
  paneHead(pane, 'AI 設定', '★ この設定は Tadori 専用です (Spira とは独立管理)。同じ値を設定してもよいですが、変更は伝播しません。');
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
    ...mkRow('チャットモデル', mkSelect(corpModelOpts, draft.chatModel, v => { draft.chatModel = v; }), 'RAG 回答生成に使うモデル (チャット欄のピッカーでも変更可)'),
    ...mkRow('Vision モデル', mkSelect(corpModelOpts, draft.visionModel, v => { draft.visionModel = v; }), 'PPTX 取り込みの画像解析専用。チャットとは独立 (高精度モデル固定推奨)'),
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
    }), 'スコアがこの値未満のメールは除外。デフォルト 0.3 (0 で無効)'),
  );
  pane.appendChild(grid);

  // ハイブリッド検索の重み (キーワード ⇄ 意味)。
  const wLabel = el('span', { class: 'mono' }, [draft.ragKeywordWeight.toFixed(2)]);
  const slider = el('input', {
    type: 'range', min: '0', max: '1', step: '0.05', value: String(draft.ragKeywordWeight),
    style: 'flex:1',
  }) as HTMLInputElement;
  slider.addEventListener('input', () => {
    draft.ragKeywordWeight = Math.min(1, Math.max(0, Number(slider.value)));
    wLabel.textContent = draft.ragKeywordWeight.toFixed(2);
  });
  pane.appendChild(el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-7)' }, ['ハイブリッド検索']));
  pane.appendChild(el('p', { class: 'tdr-hint' }, [
    '意味(ベクトル)とキーワード(文字bigram)を合成して並べ替えます。固有名詞・型番が多いならキーワード寄りに。',
  ]));
  pane.appendChild(el('div', { style: 'display:flex;align-items:center;gap:var(--s-4);margin-top:var(--s-3)' }, [
    el('span', { class: 'tdr-hint' }, ['意味']),
    slider,
    el('span', { class: 'tdr-hint' }, ['キーワード']),
    wLabel,
  ]));

  // 再ランカー (LLM で候補を並べ替え)。
  pane.appendChild(el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-7)' }, ['再ランカー (精度向上)']));
  pane.appendChild(el('p', { class: 'tdr-hint' }, [
    '検索で取った候補を AI で関連度順に並べ替えてから回答を生成します。精度は上がりますが、1 質問あたり AI コールが 1 回追加されるため利用料が増えます。',
  ]));
  const rerankCb = el('input', { type: 'checkbox' }) as HTMLInputElement;
  rerankCb.checked = !!draft.rerankEnabled;
  rerankCb.addEventListener('change', () => { draft.rerankEnabled = rerankCb.checked; });
  pane.appendChild(el('label', { style: 'display:inline-flex;align-items:center;gap:var(--s-3);margin-top:var(--s-3);cursor:pointer' }, [
    rerankCb, el('span', {}, ['再ランカーを有効にする']),
  ]));
  const grid2 = el('div', { class: 'tdr-fieldgrid', style: 'margin-top:var(--s-3)' });
  grid2.append(
    ...mkRow('候補数 (5〜30)', mkInput(String(draft.rerankCandidates), v => {
      draft.rerankCandidates = Math.min(30, Math.max(5, Number(v) || 15));
    }), '再ランカーへ渡す候補件数。多いほど精度↑だが入力トークン↑。デフォルト 15'),
  );
  pane.appendChild(grid2);
}

// ─── 除外ルール ─────────────────────────────────────────────────────────────

function buildExcludePane(pane: HTMLElement, root: HTMLElement): void {
  paneHead(pane, '除外ルール', '検索結果から除外したいメールの条件を指定します (件名 / 送信者 / To / Cc / 本文 の部分一致、大文字小文字無視)。Outlook の振り分けルールに近い動作です。1 つでも一致すると除外されます。');

  const list = el('div', { style: 'display:flex;flex-direction:column;gap:var(--s-3);margin-top:var(--s-3)' });

  function render(): void {
    list.replaceChildren();
    const rules = loadExclusionRules();
    if (rules.length === 0) {
      list.appendChild(el('div', { class: 'tdr-hint' }, ['ルールはまだありません。下の「+ ルール追加」から登録してください。']));
    } else {
      for (const r of rules) list.appendChild(renderRow(r));
    }
  }

  function renderRow(r: ExclusionRule): HTMLElement {
    const enabled = r.enabled !== false;
    const cb = el('input', { type: 'checkbox' }) as HTMLInputElement;
    cb.checked = enabled;
    cb.addEventListener('change', () => { updateExclusionRule(r.id, { enabled: cb.checked }); });

    const fieldSel = el('select', { class: 'tdr-input', style: 'max-width:120px' }) as HTMLSelectElement;
    for (const f of Object.keys(EXCLUDE_FIELD_LABELS) as ExclusionField[]) {
      const o = document.createElement('option');
      o.value = f; o.textContent = EXCLUDE_FIELD_LABELS[f];
      if (f === r.field) o.selected = true;
      fieldSel.appendChild(o);
    }
    fieldSel.addEventListener('change', () => { updateExclusionRule(r.id, { field: fieldSel.value as ExclusionField }); });

    const valueInp = el('input', { class: 'tdr-input', value: r.value, placeholder: '部分一致する文字列', style: 'flex:1' }) as HTMLInputElement;
    valueInp.addEventListener('change', () => { updateExclusionRule(r.id, { value: valueInp.value }); });

    const del = el('button', { class: 'tdr-iconbtn', 'aria-label': 'ルールを削除', title: '削除', html: icons.trash(14) });
    del.addEventListener('click', () => {
      confirmModal({
        root, title: 'ルールを削除', primaryLabel: '削除', primaryVariant: 'danger',
        message: `${EXCLUDE_FIELD_LABELS[r.field]} に「${r.value}」を含むメールを除外、というルールを削除しますか?`,
        onConfirm: () => { deleteExclusionRule(r.id); render(); },
      });
    });

    return el('div', { style: 'display:flex;align-items:center;gap:var(--s-3)' }, [cb, fieldSel, valueInp, del]);
  }

  const addBtn = el('button', { class: 'tdr-btn', style: 'margin-top:var(--s-5)' }, [
    el('span', { html: icons.plus(14) }), 'ルール追加',
  ]);
  addBtn.addEventListener('click', () => {
    addExclusionRule({ field: 'subject', value: '', enabled: true });
    render();
  });

  pane.append(list, addBtn);
  render();
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
    ...mkRow('PPTX解析の並列数 (1〜16)', mkInput(String(draft.visionConcurrency), v => { draft.visionConcurrency = Math.min(16, Math.max(1, Number(v) || 3)); }), 'PPTX のスライド解析 (Vision/テキスト) を同時に何枚処理するか。大きいほど速いが API 負荷増。デフォルト 3'),
  );
  pane.appendChild(grid);

  // ── Outlook からの既存メールインポート ──
  buildOutlookImport(pane, draft, root, siteUrl);
  buildOneNoteImport(pane, draft, root, siteUrl);
  buildDocImport(pane, draft, root, siteUrl);  // フォルダ取り込み (pptx/pdf/docx/xlsx/md/txt 統合)
  buildTranscriptImport(pane, draft, root, siteUrl);  // 会議録 (.vtt) は別系統
}

/** "YYYY-MM-DD" の since〜until 期間を月単位の半開区間に分割。
 *  各チャンクの since はその月初、until はその月末 (relay 仕様で until は「その日いっぱい含む」)。
 *  全期間が 1 ヶ月以内なら [since,until] のまま 1 チャンク。 */
function splitPeriodMonthly(sinceStr: string, untilStr: string): Array<{ since: string; until: string }> {
  const sinceD = new Date(sinceStr + 'T00:00:00');
  const untilD = new Date(untilStr + 'T00:00:00');
  if (isNaN(sinceD.getTime()) || isNaN(untilD.getTime()) || sinceD >= untilD) {
    return [{ since: sinceStr, until: untilStr }];
  }
  const iso = (d: Date): string => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const out: Array<{ since: string; until: string }> = [];
  let cur = new Date(sinceD);
  while (cur <= untilD) {
    // 当月の最終日 = 翌月 1 日の前日
    const monthEnd = new Date(cur.getFullYear(), cur.getMonth() + 1, 0);
    const chunkUntil = monthEnd < untilD ? monthEnd : untilD;
    out.push({ since: iso(cur), until: iso(chunkUntil) });
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  return out;
}

function buildOutlookImport(pane: HTMLElement, draft: RuntimeSettings, root: HTMLElement, siteUrl: string): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-8)' }, ['Outlook からインポート']));
  pane.appendChild(el('p', { class: 'tdr-hint', style: 'margin:0 0 var(--s-4)' }, [
    'ローカル中継サーバ経由で Outlook の既存メールを読み込み、To/Cc 条件と受信期間で絞って取り込みます (中継サーバの起動が必要)。',
    '期間が 1 ヶ月を超える場合は自動で月単位に分割して順次取り込みます (relay の HTTP レスポンス肥大化と途中失敗のリスクを避けるため)。',
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
    ...mkRow('1 期間あたりの最大件数', maxInp),
    el('div', {}, []), // placeholder
    el('p', { class: 'tdr-hint' }, ['月単位の各バッチで取得する上限。普通の ML なら 1000 で十分。多すぎると 1 リクエストが重くなる。']),
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
  let inSearch = false; // Outlook 検索フェーズ中か (relay 側の検索は停止できない)

  btn.addEventListener('click', () => {
    // 実行中にもう一度押す = 停止
    if (ac) {
      ac.abort();
      if (inSearch) {
        btn.textContent = 'キャンセル待ち';
        status.textContent = '※ Outlook の検索処理は relay 側で完了するまで止められません。検索完了後、取り込みを行わずに終了します。';
      } else {
        btn.textContent = '停止中…';
      }
      return;
    }

    const baseFilter = {
      to: parseAddressList(toArea.value),
      cc: parseAddressList(ccArea.value),
      max: Number(maxInp.value) || 1000,
    };
    ac = new AbortController();
    const signal = ac.signal;
    inSearch = true;
    btn.textContent = '停止';
    hideBar();

    // 期間を月単位の半開区間 [since, until] に分割 (until は「その日いっぱい含む」relay 仕様)。
    // 1 ヶ月以下なら 1 チャンクのまま。
    const sinceStr = sinceInp.value || iso(yearAgo);
    const untilStr = untilInp.value || iso(today);
    const chunks = splitPeriodMonthly(sinceStr, untilStr);

    if (!confirm(`Outlook から取り込みます (${sinceStr} 〜 ${untilStr}、${chunks.length} 期間に分割)。期間ごとに「最大 ${baseFilter.max} 件」まで取得して順次投入します。よろしいですか?`)) {
      ac = null; inSearch = false; btn.textContent = RUN_LABEL;
      status.textContent = 'キャンセル';
      return;
    }

    void (async () => {
      let totalFetched = 0, totalAdded = 0, totalSkipped = 0, totalSegments = 0;
      try {
        for (let i = 0; i < chunks.length; i++) {
          if (signal.aborted) break;
          const c = chunks[i];
          const label = `期間 ${i + 1}/${chunks.length} (${c.since} 〜 ${c.until})`;
          inSearch = true;
          status.textContent = `${label}: Outlook を検索中…`;

          const mails = await fetchOutlookMails(draft.relayBaseUrl, { ...baseFilter, since: c.since, until: c.until }, signal);
          inSearch = false;
          if (signal.aborted) break;
          totalFetched += mails.length;
          if (mails.length === 0) {
            status.textContent = `${label}: 該当 0 件 (累計 取得 ${totalFetched} / 新規 ${totalAdded})`;
            continue;
          }

          // 埋め込み + 保存。
          // 注: コールバック内では r を参照しない (r はまだ初期化中なので TDZ で死ぬ)。
          // 「現チャンクで保存した件数」は upload フェーズの done で進行値が取れる。
          let embedded = 0, saved = 0;
          const r = await ingestToSegments(toIngestMails(mails), draft, siteUrl, (phase, done, total) => {
            if (phase === 'sync') { status.textContent = `${label}: 準備中…`; return; }
            if (phase === 'embed') embedded = done;
            if (phase === 'upload') saved = done;
            const units = (total || mails.length) * 2 || 1;
            const pct = Math.min(100, Math.round((embedded + saved) / units * 100));
            // 全体進捗 = 完了済チャンク + 現チャンク内の比率
            const overall = Math.round(((i + pct / 100) / chunks.length) * 100);
            showBar(overall);
            status.textContent = `${label}: 埋め込み ${embedded}/${total} ・ 保存 ${saved}/${total} 件 (累計 取得 ${totalFetched} / 新規 ${totalAdded + saved})`;
          }, signal);
          totalAdded += r.added;
          totalSkipped += r.skipped;
          totalSegments += r.segments;
          if (r.cancelled) break;
        }
        hideBar();
        const dup = totalSkipped ? ` / 重複スキップ ${totalSkipped} 件` : '';
        if (signal.aborted) {
          status.textContent = `停止しました: ${chunks.length} 期間中 ${totalAdded} 件保存済み (セグメント ${totalSegments})${dup}`;
          toast(root, `停止 (新規 ${totalAdded} 件は保存済み)`, 'warn');
        } else if (totalAdded === 0 && totalFetched > 0) {
          status.textContent = `すべて登録済みでした (取得 ${totalFetched} 件 / 新規 0 件)`;
          toast(root, `新規なし (${totalFetched} 件はすべて登録済み)`, 'warn');
        } else if (totalFetched === 0) {
          status.textContent = `該当メールがありませんでした (${chunks.length} 期間検索済み)`;
          toast(root, '該当メール 0 件', 'warn');
        } else {
          status.textContent = `完了: ${chunks.length} 期間 / 取得 ${totalFetched} 件 / 新規 ${totalAdded} 件 (セグメント ${totalSegments})${dup}`;
          toast(root, `Outlook から ${totalAdded} 件取り込みました (${chunks.length} 期間)`, 'ok');
        }
      } catch (e) {
        if (signal.aborted || (e instanceof Error && e.name === 'AbortError')) {
          status.textContent = inSearch
            ? `Outlook の検索処理は relay 側で完了するまで継続します (結果は破棄、累計 新規 ${totalAdded} 件は保存済み)。`
            : `停止しました (累計 新規 ${totalAdded} 件は保存済み)`;
        } else {
          status.textContent = `失敗 (累計 新規 ${totalAdded} 件は保存済み)`;
          toast(root, `インポート失敗: ${e instanceof Error ? e.message : String(e)}`, 'error');
        }
      } finally {
        ac = null;
        inSearch = false;
        btn.textContent = RUN_LABEL;
        hideBar();
      }
    })();
  });

  pane.append(btn, bar, status);
}

/** OneNote 取り込み: 階層ツリーから選択 → ページ群を chunk 化して投入。 */
function buildOneNoteImport(pane: HTMLElement, draft: RuntimeSettings, root: HTMLElement, siteUrl: string): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-8)' }, ['OneNote から取り込み']));
  pane.appendChild(el('p', { class: 'tdr-hint', style: 'margin:0 0 var(--s-4)' }, [
    'OneNote のページをベクトル DB へ取り込みます (relay 経由で COM 抽出 → 自動チャンク分割)。',
    'Windows + OneNote 起動が必要。既に取り込み済みのページは初期チェック ON で表示され、チェックを外すと検索/更新チェック対象から除外されます。',
  ]));

  const loadBtn  = el('button', { class: 'tdr-btn' }, [el('span', { html: icons.notebook(14) }), 'ノートブック一覧を取得']);
  const treeEl   = el('div', { class: 'tdr-onenote-tree', style: 'max-height:280px;overflow:auto;border:1px solid var(--line);border-radius:var(--r-2);padding:var(--s-4);margin-top:var(--s-3);display:none;' });
  // ラベル: 選択ページをこのラベルの取り込みバッチとして束ねる (チャットのソース絞り込み単位)。
  const labelInput = el('input', { type: 'text', class: 'tdr-input', placeholder: 'ラベル (例: 開発マニュアル) — チャットの検索対象選択で使う', style: 'margin-top:var(--s-4);display:none;width:100%' }) as HTMLInputElement;
  const labelList  = el('datalist', { id: 'tdr-onenote-labels' });
  labelInput.setAttribute('list', 'tdr-onenote-labels');
  const batchesEl  = el('div', { style: 'margin-top:var(--s-3)' });
  const runBtn   = el('button', { class: 'tdr-btn tdr-btn--primary', style: 'margin-top:var(--s-4);display:none' }, ['選択内容を適用']);
  const stopBtn  = el('button', { class: 'tdr-btn', style: 'margin-top:var(--s-4);display:none' }, ['停止']);
  const status   = el('div', { style: 'font-size:var(--fs-sm);color:var(--ink-3);margin-top:var(--s-3)' }, ['']);
  const barFill  = el('div', { class: 'tdr-progress-fill' });
  const bar      = el('div', { class: 'tdr-progress', style: 'display:none' }, [barFill]);

  const showBar = (pct: number): void => { bar.style.display = ''; barFill.style.width = `${pct}%`; };
  const hideBar = (): void => { bar.style.display = 'none'; barFill.style.width = '0%'; };

  let notebooks: OneNoteNotebook[] = [];
  let importedIds = new Set<string>(); // DB に取り込み済みの page-id
  let excludedIds = new Set<string>(); // 検索対象から除外している page-id
  let ac: AbortController | null = null;

  function selectedPageIds(): string[] {
    const ids: string[] = [];
    treeEl.querySelectorAll<HTMLInputElement>('input[data-pid]:checked').forEach(cb => ids.push(cb.dataset.pid!));
    return ids;
  }

  /** 全ページがチェックされているノート/セクションを「コンテナ選択」として抽出。
   *  これをバッチに記録すれば、再同期でそのノート/セクション配下の新規ページも拾える。 */
  function selectedContainers(checked: Set<string>): { notebookIds: string[]; sectionIds: string[] } {
    const notebookIds: string[] = [], sectionIds: string[] = [];
    for (const nb of notebooks) {
      const allPages = nb.sections.flatMap(s => s.pages.map(p => p.id));
      if (allPages.length && allPages.every(id => checked.has(id))) { notebookIds.push(nb.id); continue; }
      for (const sec of nb.sections) {
        const sp = sec.pages.map(p => p.id);
        if (sp.length && sp.every(id => checked.has(id))) sectionIds.push(sec.id);
      }
    }
    return { notebookIds, sectionIds };
  }

  function renderTree(): void {
    treeEl.replaceChildren();
    if (notebooks.length === 0) { treeEl.appendChild(el('div', { class: 'tdr-hint' }, ['ノートブックが見つかりませんでした。'])); return; }
    for (const nb of notebooks) {
      const secList = el('div', { style: 'margin-left:var(--s-5)' });
      const nbCb = el('input', { type: 'checkbox' }) as HTMLInputElement;
      nbCb.addEventListener('change', () => {
        secList.querySelectorAll<HTMLInputElement>('input[type=checkbox]').forEach(cb => { cb.checked = nbCb.checked; });
      });
      const nbRow = el('label', { style: 'display:flex;align-items:center;gap:var(--s-2);cursor:pointer;font-weight:600;margin-top:var(--s-2)' }, [
        nbCb, el('span', { html: icons.notebook(14), style: 'display:inline-flex;color:var(--ink-3)' }), el('span', {}, [nb.name]),
      ]);
      treeEl.appendChild(nbRow);
      let nbAllChecked = true, nbAnyPage = false;
      for (const sec of nb.sections) {
        const pageList = el('div', { style: 'margin-left:var(--s-5)' });
        const secCb = el('input', { type: 'checkbox' }) as HTMLInputElement;
        secCb.addEventListener('change', () => {
          pageList.querySelectorAll<HTMLInputElement>('input[type=checkbox]').forEach(cb => { cb.checked = secCb.checked; });
        });
        const secRow = el('label', { style: 'display:flex;align-items:center;gap:var(--s-2);cursor:pointer;margin-top:var(--s-1)' }, [
          secCb, el('span', { html: icons.folder(14), style: 'display:inline-flex;color:var(--ink-3)' }), el('span', {}, [sec.name + ` (${sec.pages.length})`]),
        ]);
        secList.appendChild(secRow);
        let secAllChecked = true, secAnyPage = false;
        for (const pg of sec.pages) {
          const isImported = importedIds.has(pg.id);
          const isExcluded = excludedIds.has(pg.id);
          // 初期状態: 取り込み済み AND 除外指定なし → ON。それ以外 → OFF。
          const initialChecked = isImported && !isExcluded;
          const pgCb = el('input', { type: 'checkbox', 'data-pid': pg.id }) as HTMLInputElement;
          pgCb.checked = initialChecked;
          const tag = isImported
            ? el('span', { class: 'tdr-pill', style: 'background:var(--accent-soft);color:var(--accent-strong);font-size:var(--fs-xs);padding:1px 6px;border-radius:var(--r-1);margin-left:var(--s-2)' }, ['取り込み済み'])
            : null;
          const pgRow = el('label', { style: 'display:flex;align-items:center;gap:var(--s-2);cursor:pointer;font-size:var(--fs-sm);color:var(--ink-3)' }, [
            pgCb, el('span', {}, [pg.name]), ...(tag ? [tag] : []),
          ]);
          pageList.appendChild(pgRow);
          secAnyPage = true; nbAnyPage = true;
          if (!initialChecked) { secAllChecked = false; nbAllChecked = false; }
        }
        secList.appendChild(pageList);
        secCb.checked = secAnyPage && secAllChecked;
      }
      treeEl.appendChild(secList);
      nbCb.checked = nbAnyPage && nbAllChecked;
    }
  }

  function renderBatches(): void {
    const batches = listOneNoteBatches(siteUrl);
    labelList.replaceChildren(...batches.map(b => el('option', { value: b.label })));
    batchesEl.replaceChildren();
    if (batches.length === 0) return;
    batchesEl.appendChild(el('div', { class: 'tdr-hint', style: 'margin-bottom:var(--s-2)' }, ['取り込み済みラベル (チャットの検索対象選択に表示):']));
    for (const b of batches) {
      const del = el('button', { class: 'tdr-btn tdr-btn--sm' }, ['削除']);
      del.addEventListener('click', () => {
        removeOneNoteBatch(siteUrl, b.label);
        renderBatches();
        toast(root, `ラベル「${b.label}」を削除 (ベクトルは残ります)`, 'ok');
      });
      const cont = b.notebookIds.length + b.sectionIds.length;
      const resyncBtn = el('button', {
        class: 'tdr-btn tdr-btn--sm', title: '選択ノート/セクション配下の新規ページを取り込み、対象を最新化',
      }, ['再同期']);
      if (cont === 0) (resyncBtn as HTMLButtonElement).disabled = true; // コンテナ未指定 (個別ページのみ) は再同期対象なし
      resyncBtn.addEventListener('click', () => { void resyncBatch(b, resyncBtn); });
      const renameBtn = el('button', { class: 'tdr-btn tdr-btn--sm', title: 'ラベル名を変更 (再取り込み不要)' }, ['✏ ラベル']);
      renameBtn.addEventListener('click', () => {
        const v = prompt('ラベル名', b.label);
        if (v == null || !v.trim() || v.trim() === b.label) return;
        renameOneNoteBatch(siteUrl, b.label, v.trim());
        renderBatches();
        toast(root, `ラベルを「${v.trim()}」に変更しました`, 'ok');
      });
      batchesEl.appendChild(el('div', {
        style: 'display:flex;align-items:center;gap:var(--s-3);padding:var(--s-2) 0;border-bottom:1px solid var(--line)',
      }, [
        el('span', { class: 'tdr-pill', style: 'background:var(--accent-soft);color:var(--accent-strong);padding:1px 8px;border-radius:var(--r-1)' }, [b.label]),
        el('span', { class: 'tdr-hint', style: 'flex:1' }, [`ノート${b.notebookIds.length} / セクション${b.sectionIds.length} / ${b.pageIds.length}ページ`]),
        renameBtn, resyncBtn, del,
      ]));
    }
    batchesEl.appendChild(el('p', { class: 'tdr-hint', style: 'margin-top:var(--s-2);font-size:var(--fs-xs)' }, [
      '別ラベルを追加するには: ツリーで対象を選び直し → ラベル欄に新しい名前を入力 → 「選択内容を適用」。',
      '同じラベル名で適用すると、そのラベルにノート/セクションを追記 (和集合) します。',
    ]));
  }

  /** バッチの選択ノート/セクション配下の現行ページを取り直し、新規ページを取り込んで対象を最新化。 */
  async function resyncBatch(b: OneNoteBatch, btn: HTMLElement): Promise<void> {
    (btn as HTMLButtonElement).disabled = true; status.textContent = `「${b.label}」を再同期中…`;
    try {
      const [hier, eng] = await Promise.all([fetchOneNoteHierarchy(draft.relayBaseUrl), getEngine(siteUrl)]);
      const nbSet = new Set(b.notebookIds), secSet = new Set(b.sectionIds);
      // コンテナ配下の現行ページ ID を解決。
      const target = new Set<string>(b.pageIds); // 既存個別ページは維持
      for (const nb of hier) {
        const nbAll = nbSet.has(nb.id);
        for (const sec of nb.sections) {
          if (nbAll || secSet.has(sec.id)) for (const pg of sec.pages) target.add(pg.id);
        }
      }
      const already = eng.db.importedOneNotePageIds();
      const toIngest = [...target].filter(id => !already.has(id));
      let added = 0;
      if (toIngest.length > 0) {
        const pages = await fetchOneNotePages(draft.relayBaseUrl, { ids: toIngest, max: 2000, batchSize: 20 });
        if (pages.length > 0) {
          const r = await ingestToSegments(pagesToIngestMails(pages, b.label), draft, siteUrl, () => { /* silent */ });
          added = r.added;
        }
      }
      setOneNoteBatchPageIds(siteUrl, b.label, [...target]);
      renderBatches();
      status.textContent = `「${b.label}」再同期完了: 対象 ${target.size} ページ / 新規取り込み ${added} チャンク`;
      toast(root, `「${b.label}」再同期: 新規 ${added} チャンク`, 'ok');
    } catch (e) {
      status.textContent = '再同期失敗';
      toast(root, `再同期失敗: ${e instanceof Error ? e.message : String(e)}`, 'error');
    } finally { (btn as HTMLButtonElement).disabled = false; }
  }

  loadBtn.addEventListener('click', () => {
    void (async () => {
      loadBtn.disabled = true; status.textContent = 'OneNote の階層を取得中…';
      try {
        const [hier, eng] = await Promise.all([
          fetchOneNoteHierarchy(draft.relayBaseUrl),
          getEngine(siteUrl),
        ]);
        notebooks = hier;
        importedIds = eng.db.importedOneNotePageIds();
        excludedIds = getExcludedOneNotePageIds();
        const total = notebooks.reduce((n, nb) => n + nb.sections.reduce((m, s) => m + s.pages.length, 0), 0);
        const checkedNow = total === 0 ? 0 : (importedIds.size - [...importedIds].filter(id => excludedIds.has(id)).length);
        status.textContent = `${notebooks.length} ノートブック / ${total} ページ。取り込み済み ${importedIds.size} 件 (うち除外中 ${excludedIds.size} 件 / 検索対象 ${checkedNow} 件)。チェックの増減で「取り込み/除外」をまとめて適用します。`;
        treeEl.style.display = ''; runBtn.style.display = ''; labelInput.style.display = '';
        renderTree();
        renderBatches();
      } catch (e) {
        status.textContent = `失敗: ${e instanceof Error ? e.message : String(e)}`;
      } finally { loadBtn.disabled = false; }
    })();
  });

  runBtn.addEventListener('click', () => {
    const checked = new Set(selectedPageIds());
    const checkedAll = [...checked];
    const label = labelInput.value.trim();
    const newImports: string[] = [];
    const newlyExcluded: string[] = [];
    const reincluded: string[] = [];
    // ツリー全ページを走査して 3 つに振り分け。
    for (const nb of notebooks) for (const sec of nb.sections) for (const pg of sec.pages) {
      const isImported = importedIds.has(pg.id);
      const isChecked = checked.has(pg.id);
      if (isChecked && !isImported) newImports.push(pg.id);
      else if (!isChecked && isImported) newlyExcluded.push(pg.id);
      else if (isChecked && isImported && excludedIds.has(pg.id)) reincluded.push(pg.id);
    }
    if (newImports.length === 0 && newlyExcluded.length === 0 && reincluded.length === 0) {
      // 取り込み/除外の変更は無いが、ラベルが指定されていれば「選択ページをそのラベルに束ねる」
      // だけは実行する (既存取り込み済みページへのラベル付けはこの経路)。
      if (label && checkedAll.length) {
        const cont = selectedContainers(checked);
        recordOneNoteBatch(siteUrl, label, { pageIds: checkedAll, ...cont });
        renderBatches();
        toast(root, `ラベル「${label}」に登録 (ノート${cont.notebookIds.length}/セクション${cont.sectionIds.length}/ページ${checkedAll.length})`, 'ok');
      } else {
        toast(root, '変更なし', 'warn');
      }
      return;
    }
    const lines: string[] = [];
    if (newImports.length)   lines.push(`新規取り込み: ${newImports.length} ページ`);
    if (newlyExcluded.length) lines.push(`検索対象から除外: ${newlyExcluded.length} ページ`);
    if (reincluded.length)   lines.push(`検索対象に再有効化: ${reincluded.length} ページ`);
    if (label) lines.push(`ラベル「${label}」`);
    confirmModal({
      root, title: 'OneNote 取り込み適用', primaryLabel: '適用',
      message: lines.join(' / ') + '。長文ページは自動でチャンク分割されます。',
      onConfirm: () => { void doApply(newImports, newlyExcluded, reincluded, label, checkedAll); },
    });
  });

  stopBtn.addEventListener('click', () => { ac?.abort(); stopBtn.textContent = '停止中…'; });

  async function doApply(newImports: string[], newlyExcluded: string[], reincluded: string[], label = '', checkedAll: string[] = []): Promise<void> {
    // ラベル指定があれば、選択ページ全体 (既存取り込み済み含む) + 全選択ノート/セクションを
    // そのラベルに束ねる。検索は conversationId(=pageId) で照合するので、既存ページも
    // 再取り込み不要でラベル絞り込みできる。ノート/セクションは再同期で新規ページを拾う用。
    if (label && checkedAll.length) {
      const cont = selectedContainers(new Set(checkedAll));
      recordOneNoteBatch(siteUrl, label, { pageIds: checkedAll, ...cont });
    }
    ac = new AbortController();
    const signal = ac.signal;
    runBtn.disabled = true;
    hideBar();

    // 1) 除外指定の更新 (即時 localStorage に反映)。
    const nextExcluded = new Set(excludedIds);
    for (const id of newlyExcluded) nextExcluded.add(id);
    for (const id of reincluded)    nextExcluded.delete(id);
    setExcludedOneNotePageIds(nextExcluded);
    excludedIds = nextExcluded;
    if (newlyExcluded.length || reincluded.length) {
      status.textContent = `除外指定を更新: +${newlyExcluded.length} / −${reincluded.length}`;
    }

    // 2) 新規取り込み (任意)。
    if (newImports.length > 0) {
      stopBtn.style.display = ''; stopBtn.textContent = '停止';
      status.textContent = 'ページ本文を取得中…';
      try {
        const pages = await fetchOneNotePages(
          draft.relayBaseUrl,
          { ids: newImports, max: 1000, batchSize: 20 },
          signal,
          (done, total) => { status.textContent = `ページ本文を取得中… ${done}/${total}`; },
        );
        if (pages.length === 0) {
          status.textContent = '取得できたページがありませんでした';
        } else {
          const mails = pagesToIngestMails(pages, label || undefined);
          const skipped = newImports.length - pages.length;
          const skipNote = skipped > 0 ? ` (取得失敗 ${skipped} ページはスキップ)` : '';
          status.textContent = `${pages.length} ページ → ${mails.length} チャンク${skipNote}。埋め込みを開始…`;
          let embedded = 0, saved = 0;
          const r = await ingestToSegments(mails, draft, siteUrl, (phase, done, total) => {
            if (phase === 'sync') { status.textContent = '準備中…'; return; }
            if (phase === 'embed') embedded = done;
            if (phase === 'upload') saved = done;
            const units = (total || mails.length) * 2 || 1;
            const pct = Math.min(100, Math.round((embedded + saved) / units * 100));
            showBar(pct);
            status.textContent = `埋め込み ${embedded}/${total} ・ 保存 ${saved}/${total} チャンク (${pct}%)`;
          }, signal);
          hideBar();
          // 新規取り込んだ pageId は importedIds に追加。
          for (const p of pages) importedIds.add(p.pageId);
          const dup = r.skipped ? ` / 重複スキップ ${r.skipped}` : '';
          if (r.cancelled) {
            status.textContent = `停止しました: 新規 ${r.added} チャンク (セグメント ${r.segments})${dup}`;
            toast(root, `停止 (新規 ${r.added} チャンクは保存済み)`, 'warn');
          } else {
            status.textContent = `完了: 新規 ${pages.length} ページ / ${r.added} チャンク${dup}。除外 +${newlyExcluded.length} / 再有効化 −${reincluded.length}`;
            toast(root, `OneNote: 新規 ${r.added} チャンク / 除外 ${newlyExcluded.length} / 再有効化 ${reincluded.length}`, 'ok');
          }
        }
      } catch (e) {
        if (signal.aborted || (e instanceof Error && e.name === 'AbortError')) {
          status.textContent = '停止しました';
        } else {
          status.textContent = '失敗';
          toast(root, `OneNote 取り込み失敗: ${e instanceof Error ? e.message : String(e)}`, 'error');
        }
      } finally {
        stopBtn.style.display = 'none'; hideBar();
      }
    } else {
      // 取り込みが無く除外設定だけ変更したケース。
      status.textContent = `完了: 除外 +${newlyExcluded.length} / 再有効化 −${reincluded.length}`;
      toast(root, `OneNote: 検索対象を更新 (除外 ${newlyExcluded.length} / 再有効化 ${reincluded.length})`, 'ok');
    }
    ac = null; runBtn.disabled = false;
    renderTree(); // チェック状態を新しい importedIds / excludedIds で再描画
    renderBatches();
    if (label && checkedAll.length) toast(root, `ラベル「${label}」に ${checkedAll.length} ページを登録`, 'ok');
  }

  pane.append(loadBtn, treeEl, labelInput, labelList, el('div', { style: 'display:flex;gap:var(--s-3);align-items:center' }, [runBtn, stopBtn]), bar, status, batchesEl);
}


/** Teams 会議文字起こし (.vtt) 取り込み。SP の会議録置き場フォルダを指定して同期。
 *  設計参照: docs/teams-transcript-design.md */
function buildTranscriptImport(pane: HTMLElement, draft: RuntimeSettings, root: HTMLElement, siteUrl: string): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-8)' }, ['Teams 会議録 取り込み']));
  pane.appendChild(el('p', { class: 'tdr-hint', style: 'margin:0 0 var(--s-4)' }, [
    'Teams 会議の文字起こし (.vtt) を SharePoint のフォルダから取り込みます。',
    '運用: SP に「会議録置き場」フォルダを 1 つ作り、対象の .vtt を手動アップロード → このフォルダを登録して同期。',
    '発言は話者・時刻付きでチャンク化してそのまま埋め込みます (LLM 要約なし)。Vision/relay 不要、ブラウザ完結。',
    '※ .vtt 内に会議名・日時は無いため、ファイル名から推定します。',
  ]));

  const urlInput = el('input', { type: 'text', class: 'tdr-input', placeholder: 'https://contoso.sharepoint.com/sites/foo/Shared Documents/会議録' }) as HTMLInputElement;
  urlInput.style.flex = '1';
  const labelInput = el('input', { type: 'text', class: 'tdr-input', placeholder: 'ラベル (任意)' }) as HTMLInputElement;
  labelInput.style.width = '160px';
  const recursiveCb = el('input', { type: 'checkbox' }) as HTMLInputElement;
  const recursiveLabel = el('label', { style: 'display:flex;align-items:center;gap:var(--s-2);font-size:var(--fs-sm);color:var(--ink-3);white-space:nowrap' }, [recursiveCb, '再帰']);
  const addBtn = el('button', { class: 'tdr-btn' }, ['追加']);
  const addRow = el('div', { style: 'display:flex;gap:var(--s-2);align-items:center;margin-top:var(--s-2)' }, [urlInput, labelInput, recursiveLabel, addBtn]);

  const listEl = el('div', { style: 'margin-top:var(--s-4);display:flex;flex-direction:column;gap:var(--s-3)' });
  const syncAllBtn = el('button', { class: 'tdr-btn tdr-btn--primary' }, [el('span', { html: icons.mic(14) }), 'すべて同期']);
  const stopBtn = el('button', { class: 'tdr-btn', style: 'display:none' }, ['停止']);
  const status = el('div', { style: 'font-size:var(--fs-sm);color:var(--ink-3);margin-top:var(--s-3)' }, ['']);
  const barFill = el('div', { class: 'tdr-progress-fill' });
  const bar = el('div', { class: 'tdr-progress', style: 'display:none' }, [barFill]);
  const showBar = (pct: number): void => { bar.style.display = ''; barFill.style.width = `${pct}%`; };
  const hideBar = (): void => { bar.style.display = 'none'; barFill.style.width = '0%'; };

  let ac: AbortController | null = null;

  function renderList(): void {
    listEl.replaceChildren();
    const folders = listTranscriptFolders(siteUrl);
    if (folders.length === 0) {
      listEl.appendChild(el('div', { class: 'tdr-hint' }, ['まだ取り込みフォルダが登録されていません。']));
      syncAllBtn.disabled = true;
      return;
    }
    syncAllBtn.disabled = false;
    for (const f of folders) {
      const lastSync = f.lastSyncAt ? new Date(f.lastSyncAt).toLocaleString() : '未同期';
      const fileCount = Object.keys(f.perFile).length;
      const renameBtn = el('button', { class: 'tdr-btn tdr-btn--sm', title: 'ラベルを変更 (再取り込み不要)' }, ['✏ ラベル']);
      renameBtn.addEventListener('click', () => {
        const v = prompt('表示ラベル', f.label || deriveTranscriptLabel(f.url));
        if (v == null) return;
        addTranscriptFolder(siteUrl, { url: f.url, label: v.trim(), recursive: f.recursive });
        renderList();
        toast(root, 'ラベルを変更しました', 'ok');
      });
      const head = el('div', { style: 'display:flex;align-items:center;gap:var(--s-2);font-weight:600' }, [
        el('span', { html: icons.mic(14), style: 'display:inline-flex;color:var(--ink-3)' }),
        el('span', { class: 'mono', style: 'font-size:var(--fs-sm)' }, [f.label || deriveTranscriptLabel(f.url)]),
        renameBtn,
      ]);
      const meta = el('div', { class: 'tdr-hint', style: 'margin-top:var(--s-1);font-size:var(--fs-xs)' }, [
        `URL: ${f.url}`, el('br'),
        `最終同期: ${lastSync}　/　会議 (.vtt): ${fileCount} 件　/　${f.recursive ? '再帰あり' : '直下のみ'}`,
      ]);
      const syncBtn = el('button', { class: 'tdr-btn', style: 'font-size:var(--fs-sm)' }, ['同期']);
      const delBtn = el('button', { class: 'tdr-btn', style: 'font-size:var(--fs-sm)' }, ['削除']);
      const actions = el('div', { style: 'display:flex;gap:var(--s-2);margin-top:var(--s-2)' }, [syncBtn, delBtn]);
      const card = el('div', { style: 'border:1px solid var(--line);border-radius:var(--r-2);padding:var(--s-3)' }, [head, meta, actions]);
      syncBtn.addEventListener('click', () => { void runSync([f]); });
      delBtn.addEventListener('click', () => {
        confirmModal({
          root, title: '会議録フォルダ設定を削除',
          message: `「${f.label || f.url}」の設定を削除します。\n(取り込み済みのチャンクはベクトル DB に残ります)`,
          primaryLabel: '削除', primaryVariant: 'danger',
          onConfirm: () => { removeTranscriptFolder(siteUrl, f.url); renderList(); toast(root, 'フォルダ設定を削除しました', 'ok'); },
        });
      });
      listEl.appendChild(card);
    }
  }

  addBtn.addEventListener('click', () => {
    const url = urlInput.value.trim();
    if (!url) { toast(root, 'フォルダ URL を入力してください', 'warn'); return; }
    if (!/^https?:\/\//i.test(url) && !url.startsWith('/')) {
      toast(root, 'URL は https://... か /sites/... の形式で入力してください', 'warn'); return;
    }
    addTranscriptFolder(siteUrl, { url, label: labelInput.value.trim() || undefined, recursive: recursiveCb.checked });
    urlInput.value = ''; labelInput.value = '';
    renderList();
    toast(root, 'フォルダを追加しました。「同期」で取り込みを開始してください', 'ok');
  });

  async function runSync(folders: TranscriptFolderConfig[]): Promise<void> {
    if (ac) return;
    ac = new AbortController();
    syncAllBtn.style.display = 'none'; stopBtn.style.display = '';
    showBar(0);
    let totalChunks = 0, totalSkipped = 0, totalDeleted = 0, totalFailed = 0;
    try {
      for (let i = 0; i < folders.length; i++) {
        if (ac.signal.aborted) break;
        const f = folders[i];
        status.textContent = `[${i + 1}/${folders.length}] ${f.label || deriveTranscriptLabel(f.url)} を同期中…`;
        const r = await syncTranscriptFolder(
          f, draft, siteUrl,
          (p: TranscriptIngestProgress) => {
            const fileLabel = p.file ? `${p.file} (${p.fileIdx}/${p.fileTotal})` : '一覧取得中';
            status.textContent = `${fileLabel} — ${p.message ?? p.phase}`;
            if (p.fileTotal > 0) {
              const pct = Math.round((p.fileIdx - 1 + (p.chunkTotal > 0 ? p.chunkIdx / p.chunkTotal : 0)) / p.fileTotal * 100);
              showBar(Math.min(99, Math.max(0, pct)));
            }
          },
          ac.signal,
        );
        totalChunks += r.ingestedChunks;
        totalSkipped += r.skippedFiles;
        totalDeleted += r.deletedFiles;
        totalFailed += r.failedFiles;
      }
      showBar(100);
      const msg = `完了: ${totalChunks} チャンク取込 / スキップ ${totalSkipped} 件 / 削除 ${totalDeleted} 件${totalFailed ? ` / 失敗 ${totalFailed} 件` : ''}`;
      status.textContent = msg;
      toast(root, msg, totalFailed ? 'warn' : 'ok');
      renderList();
    } catch (e) {
      if ((e as Error).name === 'AbortError') {
        status.textContent = '停止しました (取り込み済みは保存済み)';
        toast(root, '取り込みを停止しました', 'warn');
      } else {
        status.textContent = `失敗: ${(e as Error).message}`;
        toast(root, `取り込み失敗: ${(e as Error).message}`, 'error');
      }
    } finally {
      ac = null;
      syncAllBtn.style.display = ''; stopBtn.style.display = 'none';
      setTimeout(hideBar, 1500);
    }
  }

  syncAllBtn.addEventListener('click', () => { void runSync(listTranscriptFolders(siteUrl)); });
  stopBtn.addEventListener('click', () => { ac?.abort(); });

  pane.append(addRow, listEl, el('div', { style: 'display:flex;gap:var(--s-3);align-items:center;margin-top:var(--s-3)' }, [syncAllBtn, stopBtn]), bar, status);
  renderList();
}

/** ドキュメント (docx/doc/pdf/md/txt) 取り込み。設定側でフォルダを登録・同期する。
 *  検索対象に「含めるか」はチャット画面の「文書フォルダ」スコープ選択で切替。 */
function buildDocImport(pane: HTMLElement, draft: RuntimeSettings, root: HTMLElement, siteUrl: string): void {
  pane.appendChild(el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-8)' }, ['フォルダ取り込み']));
  pane.appendChild(el('p', { class: 'tdr-hint', style: 'margin:0 0 var(--s-4)' }, [
    'SharePoint のフォルダを指定して、配下の pptx / pdf / docx / xlsx / md / txt を種別問わずまとめて取り込みます。',
    'pdf/docx は relay で本文抽出 (relay 起動が必要)。md/txt は relay 不要。',
    'pptx は「Vision 解析」ONで画像も含めて解説化、OFFでテキスト/表/ノートのみ取り込み (ラベル=フォルダ単位で指定)。',
    '※ 検索対象はチャットの「＋ 検索対象」で文書/PPTX のフォルダ(ラベル)から選べます。',
  ]));

  const urlInput = el('input', { type: 'text', class: 'tdr-input', placeholder: 'https://contoso.sharepoint.com/sites/foo/Shared Documents/資料' }) as HTMLInputElement;
  urlInput.style.flex = '1';
  const labelInput = el('input', { type: 'text', class: 'tdr-input', placeholder: 'ラベル (任意)' }) as HTMLInputElement;
  labelInput.style.width = '160px';
  const recursiveCb = el('input', { type: 'checkbox' }) as HTMLInputElement;
  recursiveCb.checked = true;
  const recursiveLabel = el('label', { style: 'display:flex;align-items:center;gap:var(--s-2);font-size:var(--fs-sm);color:var(--ink-3);white-space:nowrap' }, [recursiveCb, '再帰']);
  const visionCb = el('input', { type: 'checkbox' }) as HTMLInputElement;
  const visionLabel = el('label', { style: 'display:flex;align-items:center;gap:var(--s-2);font-size:var(--fs-sm);color:var(--ink-3);white-space:nowrap', title: 'pptx を画像込みで Vision 解析 (OFF=テキストのみ)' }, [visionCb, 'pptx Vision']);
  const addBtn = el('button', { class: 'tdr-btn' }, ['追加']);
  const addRow = el('div', { style: 'display:flex;gap:var(--s-2);align-items:center;margin-top:var(--s-2);flex-wrap:wrap' }, [urlInput, labelInput, recursiveLabel, visionLabel, addBtn]);

  const listEl = el('div', { style: 'margin-top:var(--s-4);display:flex;flex-direction:column;gap:var(--s-3)' });
  const syncAllBtn = el('button', { class: 'tdr-btn tdr-btn--primary' }, [el('span', { html: icons.fileText(14) }), 'すべて同期']);
  const stopBtn = el('button', { class: 'tdr-btn', style: 'display:none' }, ['停止']);
  const status = el('div', { style: 'font-size:var(--fs-sm);color:var(--ink-3);margin-top:var(--s-3)' }, ['']);
  const barFill = el('div', { class: 'tdr-progress-fill' });
  const bar = el('div', { class: 'tdr-progress', style: 'display:none' }, [barFill]);
  const showBar = (pct: number): void => { bar.style.display = ''; barFill.style.width = `${pct}%`; };
  const hideBar = (): void => { bar.style.display = 'none'; barFill.style.width = '0%'; };

  let ac: AbortController | null = null;

  function renderList(): void {
    listEl.replaceChildren();
    const folders = listDocFolders(siteUrl);
    if (folders.length === 0) {
      listEl.appendChild(el('div', { class: 'tdr-hint' }, ['まだ取り込みフォルダが登録されていません。']));
      syncAllBtn.disabled = true;
      return;
    }
    syncAllBtn.disabled = false;
    for (const f of folders) {
      const lastSync = f.lastSyncAt ? new Date(f.lastSyncAt).toLocaleString() : '未同期';
      const fileCount = Object.keys(f.perFile).length;
      const renameBtn = el('button', { class: 'tdr-btn tdr-btn--sm', title: 'ラベルを変更 (再取り込み不要)' }, ['✏ ラベル']);
      renameBtn.addEventListener('click', () => {
        const v = prompt('表示ラベル', f.label || deriveDocLabel(f.url));
        if (v == null) return;
        addDocFolder(siteUrl, { url: f.url, label: v.trim(), recursive: f.recursive, visionForPptx: f.visionForPptx });
        renderList();
        toast(root, 'ラベルを変更しました', 'ok');
      });
      const visionToggle = el('input', { type: 'checkbox' }) as HTMLInputElement;
      visionToggle.checked = f.visionForPptx === true;
      visionToggle.addEventListener('change', () => {
        addDocFolder(siteUrl, { url: f.url, label: f.label, recursive: f.recursive, visionForPptx: visionToggle.checked });
        toast(root, `pptx Vision を ${visionToggle.checked ? 'ON' : 'OFF'} (変更後は「同期」、既存分は強制再取り込みが必要)`, 'ok');
      });
      const head = el('div', { style: 'display:flex;align-items:center;gap:var(--s-2);font-weight:600;flex-wrap:wrap' }, [
        el('span', { html: icons.fileText(14), style: 'display:inline-flex;color:var(--ink-3)' }),
        el('span', { class: 'mono', style: 'font-size:var(--fs-sm)' }, [f.label || deriveDocLabel(f.url)]),
        renameBtn,
        el('label', { style: 'display:flex;align-items:center;gap:4px;font-weight:400;font-size:var(--fs-xs);color:var(--ink-3)', title: 'pptx を Vision 解析 (OFF=テキストのみ)' }, [visionToggle, 'pptx Vision']),
      ]);
      const meta = el('div', { class: 'tdr-hint', style: 'margin-top:var(--s-1);font-size:var(--fs-xs)' }, [
        `URL: ${f.url}`, el('br'),
        `最終同期: ${lastSync}　/　文書: ${fileCount} 件　/　${f.recursive ? '再帰あり' : '直下のみ'}`,
      ]);
      const syncBtn = el('button', { class: 'tdr-btn', style: 'font-size:var(--fs-sm)' }, ['同期']);
      const forceBtn = el('button', { class: 'tdr-btn', style: 'font-size:var(--fs-sm)', title: '更新時刻に関係なく全ファイルを再解析 (pptx Vision 切替後・別サイトへ再投入時など)' }, ['強制再取り込み']);
      const delBtn = el('button', { class: 'tdr-btn', style: 'font-size:var(--fs-sm)' }, ['削除']);
      const actions = el('div', { style: 'display:flex;gap:var(--s-2);margin-top:var(--s-2)' }, [syncBtn, forceBtn, delBtn]);
      const card = el('div', { style: 'border:1px solid var(--line);border-radius:var(--r-2);padding:var(--s-3)' }, [head, meta, actions]);
      syncBtn.addEventListener('click', () => { void runSync([f]); });
      forceBtn.addEventListener('click', () => { void runSync([f], { force: true }); });
      delBtn.addEventListener('click', () => {
        confirmModal({
          root, title: 'ドキュメントフォルダ設定を削除',
          message: `「${f.label || f.url}」の設定を削除します。\n(取り込み済みのチャンクはベクトル DB に残ります)`,
          primaryLabel: '削除', primaryVariant: 'danger',
          onConfirm: () => { removeDocFolder(siteUrl, f.url); renderList(); toast(root, 'フォルダ設定を削除しました', 'ok'); },
        });
      });
      listEl.appendChild(card);
    }
  }

  addBtn.addEventListener('click', () => {
    const url = urlInput.value.trim();
    if (!url) { toast(root, 'フォルダ URL を入力してください', 'warn'); return; }
    if (!/^https?:\/\//i.test(url) && !url.startsWith('/')) { toast(root, 'URL は https://... か /sites/... の形式で', 'warn'); return; }
    const label = labelInput.value.trim() || undefined;
    addDocFolder(siteUrl, { url, label, recursive: recursiveCb.checked, visionForPptx: visionCb.checked });
    urlInput.value = ''; labelInput.value = ''; visionCb.checked = false;
    renderList();
    toast(root, 'フォルダを追加しました。「同期」で取り込みを開始してください', 'ok');
  });

  async function runSync(folders: DocFolderConfig[], runOpts: { force?: boolean } = {}): Promise<void> {
    if (ac) return;
    ac = new AbortController();
    syncAllBtn.style.display = 'none'; stopBtn.style.display = '';
    showBar(0);
    let totalChunks = 0, totalSkipped = 0, totalDeleted = 0, totalFailed = 0;
    try {
      for (let i = 0; i < folders.length; i++) {
        if (ac.signal.aborted) break;
        const f = folders[i];
        status.textContent = `[${i + 1}/${folders.length}] ${f.label || deriveDocLabel(f.url)} を同期中…`;
        const r = await syncDocFolder(
          f, draft, siteUrl,
          (p: DocIngestProgress) => {
            const fileLabel = p.file ? `${p.file} (${p.fileIdx}/${p.fileTotal})` : '一覧取得中';
            status.textContent = `${fileLabel} — ${p.message ?? p.phase}`;
            if (p.fileTotal > 0) showBar(Math.min(99, Math.max(0, Math.round((p.fileIdx - 1) / p.fileTotal * 100))));
          },
          ac.signal,
          { force: runOpts.force },
        );
        totalChunks += r.ingestedChunks; totalSkipped += r.skippedFiles; totalDeleted += r.deletedFiles; totalFailed += r.failedFiles;

        // 同フォルダの pptx を pptx パイプラインで取り込む (Vision はフォルダ設定で ON/OFF)。
        // pptxFolders は使わず、perFile は docFolder.pptxPerFile に持たせる (専用登録を作らない)。
        if (ac.signal.aborted) break;
        const pptxCfg: PptxFolderConfig = {
          url: f.url, label: f.label, recursive: f.recursive,
          perFile: f.pptxPerFile || {}, lastSyncAt: f.lastSyncAt,
        };
        const pr = await syncPptxFolder(
          pptxCfg, draft, siteUrl,
          (p) => { status.textContent = `pptx: ${p.file || ''} ${p.slideIdx}/${p.slideTotal} — ${p.message ?? p.phase}`; },
          ac.signal,
          { vision: f.visionForPptx === true, force: runOpts.force, persist: (pf) => updateDocFolderPptxSync(siteUrl, f.url, pf) },
        );
        totalChunks += pr.ingestedSlides; totalDeleted += pr.deletedFiles; totalFailed += pr.failedSlides;
      }
      showBar(100);
      const msg = `完了: ${totalChunks} チャンク取込 / スキップ ${totalSkipped} 件 / 削除 ${totalDeleted} 件${totalFailed ? ` / 失敗 ${totalFailed} 件` : ''}`;
      status.textContent = msg;
      toast(root, msg, totalFailed ? 'warn' : 'ok');
      renderList();
    } catch (e) {
      if ((e as Error).name === 'AbortError') { status.textContent = '停止しました (取り込み済みは保存済み)'; toast(root, '取り込みを停止しました', 'warn'); }
      else { status.textContent = `失敗: ${(e as Error).message}`; toast(root, `取り込み失敗: ${(e as Error).message}`, 'error'); }
    } finally {
      ac = null; syncAllBtn.style.display = ''; stopBtn.style.display = 'none'; setTimeout(hideBar, 1500);
    }
  }

  syncAllBtn.addEventListener('click', () => { void runSync(listDocFolders(siteUrl)); });
  stopBtn.addEventListener('click', () => { ac?.abort(); });

  pane.append(addRow, listEl, el('div', { style: 'display:flex;gap:var(--s-3);align-items:center;margin-top:var(--s-3)' }, [syncAllBtn, stopBtn]), bar, status);
  renderList();
}

// ─── PA セットアップ ─────────────────────────────────────────────────────────
// Power Automate で新着メールを Tadori 受信メール List に投入する手順を案内。
// Spira の同名ペインと同じ作法 (順序付きステップ + コピー可能フィールド)。

function buildPaSetupPane(pane: HTMLElement, draft: RuntimeSettings, _root: HTMLElement, siteUrl: string): void {
  paneHead(pane, 'PA セットアップ', '新着メールを Tadori が自動で拾えるようにするための Power Automate の設定手順です。手動の Outlook 取り込みボタンを毎回押さなくても、PA がメールを List に書き込み → Tadori 自動取り込みが拾う流れになります。');

  const listName = draft.listTitle || 'Tadori 受信メール';

  // コピー可能なテキストフィールドのヘルパ (Spira と同じ作法)
  const copyable = (label: string, value: string, hint?: string): HTMLElement => {
    const inp = el('input', { class: 'tdr-input', readonly: 'readonly', value, style: 'font-family:var(--font-mono);font-size:var(--fs-sm)' }) as HTMLInputElement;
    const btn = el('button', { class: 'tdr-btn', style: 'flex-shrink:0' }, [el('span', { html: icons.copy(14) }), 'コピー']);
    btn.addEventListener('click', () => {
      void navigator.clipboard?.writeText(value).then(() => {
        btn.replaceChildren(el('span', { html: icons.check(14) }), 'コピー済');
        setTimeout(() => btn.replaceChildren(el('span', { html: icons.copy(14) }), 'コピー'), 1500);
      });
    });
    return el('div', { style: 'margin-bottom:var(--s-3)' }, [
      el('label', { class: 'tdr-label' }, [label]),
      el('div', { style: 'display:flex;gap:var(--s-2)' }, [inp, btn]),
      ...(hint ? [el('p', { class: 'tdr-hint', style: 'margin-top:var(--s-1)' }, [hint])] : []),
    ]);
  };

  const step = (n: number, title: string, body: HTMLElement[]): HTMLElement => {
    return el('div', { style: 'display:flex;gap:var(--s-4);margin-bottom:var(--s-5)' }, [
      el('div', { style: 'flex-shrink:0;width:32px;height:32px;border-radius:50%;background:var(--accent-soft);color:var(--accent-strong);display:flex;align-items:center;justify-content:center;font-weight:700' }, [String(n)]),
      el('div', { style: 'flex:1;min-width:0' }, [
        el('div', { style: 'font-weight:600;margin-bottom:var(--s-2)' }, [title]),
        ...body,
      ]),
    ]);
  };

  // 各種フィールド名 (PA から書き込む列)
  const colMap: Array<{ name: string; from: string }> = [
    { name: 'Title',        from: '件名 (Subject)' },
    { name: 'From',         from: '送信者アドレス' },
    { name: 'ToAddrs',      from: '宛先一覧 (改行区切り)' },
    { name: 'CcAddrs',      from: 'Cc 一覧 (改行区切り)' },
    { name: 'ReceivedTime', from: '受信日時 (ISO 8601)' },
    { name: 'MessageId',    from: 'Internet-Message-Id (重複排除キー)' },
    { name: 'Body',         from: '本文 (テキストまたは HTML)' },
    { name: 'IsHtml',       from: 'HTML 形式なら true' },
  ];
  const mapTable = el('table', { style: 'width:100%;border-collapse:collapse;font-size:var(--fs-sm);margin-top:var(--s-2)' }, [
    el('thead', {}, [
      el('tr', { style: 'background:var(--paper-2);border-bottom:1px solid var(--line)' }, [
        el('th', { style: 'text-align:left;padding:6px 10px;font-weight:600' }, ['列名 (List 側)']),
        el('th', { style: 'text-align:left;padding:6px 10px;font-weight:600' }, ['PA で渡す内容']),
      ]),
    ]),
    el('tbody', {},
      colMap.map(c => el('tr', { style: 'border-bottom:1px solid var(--line)' }, [
        el('td', { style: 'padding:6px 10px;font-family:var(--font-mono);color:var(--accent-strong)' }, [c.name]),
        el('td', { style: 'padding:6px 10px;color:var(--ink-3)' }, [c.from]),
      ])),
    ),
  ]);

  pane.append(
    step(1, 'Power Automate を開く', [
      el('p', { class: 'tdr-hint', style: 'margin:0' }, ['ブラウザで Power Automate (Microsoft 365) を開いて「マイ フロー」→「新しいフロー」→「自動化したクラウド フロー」を選択します。']),
    ]),
    step(2, 'トリガを設定', [
      el('p', { class: 'tdr-hint', style: 'margin:0 0 var(--s-2)' }, ['「Outlook.com / Office 365 Outlook」 → 「新しいメールが届いたとき (V3)」を選択。受信フォルダ・差出人・件名フィルタなどを指定 (例: ML アドレス宛のみ)。']),
    ]),
    step(3, 'アクション: SharePoint「項目の作成」', [
      el('p', { class: 'tdr-hint', style: 'margin:0 0 var(--s-3)' }, ['アクション「項目の作成 (Create item)」を追加し、以下を指定します:']),
      copyable('サイト URL', siteUrl, 'PA の「サイトのアドレス」にこの URL をそのまま貼ります。'),
      copyable('List 名', listName, 'PA の「リスト名」のドロップダウンに表示されない時はこの文字列を直接入力してください (初回起動時に自動作成されます)。'),
    ]),
    step(4, 'フィールドマッピング', [
      el('p', { class: 'tdr-hint', style: 'margin:0 0 var(--s-2)' }, ['PA のフィールドに以下を割り当てます:']),
      mapTable,
      el('p', { class: 'tdr-hint', style: 'margin-top:var(--s-3)' }, ['※ MessageId はメールの一意 ID (Internet-Message-Id ヘッダ)。同じメールを 2 度取り込まないための重複排除キーなので必須。']),
    ]),
    step(5, 'フローを保存して有効化', [
      el('p', { class: 'tdr-hint', style: 'margin:0' }, ['右上の「保存」→「テスト」で新着メールを 1 件テスト → SharePoint で List 行が作られていれば成功。以降、新着が来るたびに自動で投入され、Tadori 側の自動取り込みが拾います。']),
    ]),
    step(6, 'Tadori 側の確認', [
      el('p', { class: 'tdr-hint', style: 'margin:0' }, ['Tadori のトップバー右側の在席チップで「書込: ◯◯」担当が決まっており、relay (PowerShell) が起動していれば、List に入った新着メールは自動で embed → SharePoint セグメントへ書き込まれます。明示的にボタンを押す必要はありません。']),
    ]),
  );

  pane.appendChild(el('div', { class: 'tdr-hint', style: 'padding:var(--s-4);background:var(--accent-soft);border-radius:var(--r-2);margin-top:var(--s-3);color:var(--accent-strong)' }, [
    'ℹ️ List がまだ存在しない場合は、Tadori を一度起動するだけで自動作成されます。PA 側で「リスト名のドロップダウンが空」になっていたら、まず Tadori を再起動してから PA を設定してください。',
  ]));
}

// ─── Tadori について (技術仕様) ───────────────────────────────────────────
// Spira の「Spira について」と同じ作法で、アーキテクチャ概要と主要 ADR を提示。

function buildAboutPane(pane: HTMLElement, _root: HTMLElement): void {
  paneHead(pane, 'Tadori について', 'Tadori (辿り) はメール / OneNote / PPTX / Teams 会議録を意味検索 + RAG で辿るブックマークレット型ツールです。');

  const buildId = (window as unknown as { __TADORI_BUILD_ID__?: string }).__TADORI_BUILD_ID__ || '(不明)';

  const section = (title: string, body: HTMLElement | HTMLElement[]): HTMLElement => {
    return el('div', { style: 'margin-bottom:var(--s-6)' }, [
      el('p', { class: 'tdr-pane-title', style: 'margin-bottom:var(--s-2)' }, [title]),
      ...(Array.isArray(body) ? body : [body]),
    ]);
  };

  pane.appendChild(section('概要',
    el('p', { class: 'tdr-hint' }, [
      'メーリングリスト等の過去メール / OneNote ページ / PowerPoint マニュアル / Teams 会議の文字起こしをまとめて意味検索 (ベクトル類似度 + キーワード一致) し、AI が出典付きで回答するツール。回答内容は OneNote の任意ページに「Tadori 追記」として書き戻せます。委任先 M365 制約下で動作 (独立サーバ不要)。',
    ]),
  ));

  pane.appendChild(section('検索対象ソース (kind)',
    el('ul', { class: 'tdr-hint', style: 'margin:0;padding-left:var(--s-5)' }, [
      el('li', {}, ['メール — Outlook / Power Automate 経由で取り込み']),
      el('li', {}, ['OneNote — relay (COM) でページ抽出']),
      el('li', {}, ['PPTX — SP フォルダの .pptx を relay + Vision LLM で markdown 化']),
      el('li', {}, ['会議 (Teams) — SP フォルダの .vtt を話者・時刻付きで取り込み (relay/Vision 不要)']),
      el('li', { style: 'margin-top:var(--s-2);color:var(--ink-4)' }, ['※ チャット入力欄上の「検索対象」チップで種別ごとに ON/OFF 可能']),
    ]),
  ));

  pane.appendChild(section('アーキテクチャ', [
    el('pre', { style: 'background:var(--paper-2);border:1px solid var(--line);border-radius:var(--r-2);padding:var(--s-4);font-size:var(--fs-xs);overflow-x:auto;line-height:1.5' }, [
`[Outlook] (各メンバの業務 PC)
   │ COM (relay 経由)
   ▼
[ローカル relay (PowerShell)]
   ├ Outlook COM → メール取り込み
   ├ OneNote COM → ページ取り込み / 追記 / 更新
   └ Azure OpenAI / Anthropic ゲートウェイ (チャット + 埋め込み)

[ブラウザ (bookmarklet)]
   ├ ベクトル DB (in-memory + IndexedDB cache)
   ├ ハイブリッド検索 (cosine + 文字 bigram)
   ├ LLM クエリルータ → 意味 + 完全一致を自動分担
   ├ RAG 回答生成 (SSE ストリーミング + Markdown レンダ)
   └ Sticky モード (ハートビート + 自動取り込み)
   ↑↓ Cookie 認証 SP REST
[SharePoint ドキュメントライブラリ / Shared Documents/Tadori/]
   ├ manifest.json (世代管理)
   └ seg-NNNNN.json (追記専用セグメント、≤100 件)
[SharePoint List]
   ├ Tadori 受信メール (PA からの投入バッファ)
   ├ Tadori Sync (在席ハートビート + 書き込みリース)
   └ Tadori 利用料 (LLM コスト集計)`,
    ]),
  ]));

  pane.appendChild(section('データフロー', el('ul', { style: 'margin:0;padding-left:var(--s-6)' }, [
    el('li', {}, ['取り込み: relay COM か PA → List → 自動取り込み (writer 担当のみ) → 埋め込み (件名 + 本文) → SharePoint セグメントへ書き込み → 全員へ配布']),
    el('li', {}, ['検索: 質問 → LLM クエリルータ (keyword + 意味分担) → ベクトル DB ハイブリッド検索 → 同一 OneNote ページのチャンク重複排除 → リランカ (任意)']),
    el('li', {}, ['回答: Top-K の参照を context に LLM が出典 [n] 付き Markdown 生成 → 質問末尾でフォローアップ案も生成']),
    el('li', {}, ['追記: AI が OneNote ページ末尾に「Tadori 追記」Outline を新規挿入 (バナー + 見出し + 本文 + 出典フッター)。バナーが識別子になり既存追記の更新も可能']),
  ])));

  pane.appendChild(section('SharePoint List 構成', el('table', { style: 'width:100%;border-collapse:collapse;font-size:var(--fs-sm)' }, [
    el('thead', {}, [
      el('tr', { style: 'background:var(--paper-2);border-bottom:1px solid var(--line)' }, [
        el('th', { style: 'text-align:left;padding:6px 10px' }, ['リスト名']),
        el('th', { style: 'text-align:left;padding:6px 10px' }, ['用途']),
        el('th', { style: 'text-align:left;padding:6px 10px' }, ['作成主体']),
      ]),
    ]),
    el('tbody', {}, [
      el('tr', { style: 'border-bottom:1px solid var(--line)' }, [
        el('td', { style: 'padding:6px 10px;font-family:var(--font-mono);color:var(--accent-strong)' }, ['Tadori 受信メール']),
        el('td', { style: 'padding:6px 10px' }, ['PA からの新着メール投入先 (Tadori が自動取り込み)']),
        el('td', { style: 'padding:6px 10px;color:var(--ink-3)' }, ['Tadori 起動時に自動']),
      ]),
      el('tr', { style: 'border-bottom:1px solid var(--line)' }, [
        el('td', { style: 'padding:6px 10px;font-family:var(--font-mono);color:var(--accent-strong)' }, ['Tadori Sync']),
        el('td', { style: 'padding:6px 10px' }, ['在席ハートビート + 書き込みリース (writer 単一化)']),
        el('td', { style: 'padding:6px 10px;color:var(--ink-3)' }, ['Tadori 起動時に自動']),
      ]),
      el('tr', { style: 'border-bottom:1px solid var(--line)' }, [
        el('td', { style: 'padding:6px 10px;font-family:var(--font-mono);color:var(--accent-strong)' }, ['Tadori 利用料']),
        el('td', { style: 'padding:6px 10px' }, ['月次 LLM コスト集計 (ユーザ別)']),
        el('td', { style: 'padding:6px 10px;color:var(--ink-3)' }, ['Tadori 起動時に自動']),
      ]),
    ]),
  ])));

  pane.appendChild(section('主要 ADR (Notion 設計ドキュメント参照)', el('ul', { style: 'margin:0;padding-left:var(--s-6)' }, [
    el('li', {}, ['ADR-001: 独立ベクトル DB サーバを立てない → クライアントサイドベクトル検索']),
    el('li', {}, ['ADR-007: 認証は Cookie 再利用 (Azure AD アプリ登録不可のため)']),
    el('li', {}, ['ADR-010 / 011: ベクトル DB 本体は relay の SQLite → SharePoint セグメントで配布']),
    el('li', {}, ['ADR-012: 書き込みは List リースで単一化 (同時 1 人、輪番)']),
    el('li', {}, ['ADR-013: OneNote を取り込み + 追記対象に拡張 (relay COM 経由)']),
    el('li', {}, ['ADR-014: AI 追記候補は B 方式 (確認 1 クリック) で書き込み']),
  ])));

  pane.appendChild(section('制約', el('ul', { style: 'margin:0;padding-left:var(--s-6)' }, [
    el('li', {}, ['Azure AD アプリ登録不可 → Cookie / MSAL 公開クライアントのみ']),
    el('li', {}, ['外部 SaaS への送信不可 → 社内 Azure OpenAI 経由 (開発者モードのみ Claude/Voyage 直叩き可)']),
    el('li', {}, ['Outlook の「オフラインに保持するメール」設定外のメールは取得不可 (キャッシュ依存)']),
    el('li', {}, ['npm 不可 / Python 不可 → JS/TS のみ、dist は git にコミット']),
  ])));

  pane.appendChild(section('ビルド情報',
    el('p', { class: 'tdr-hint', style: 'font-family:var(--font-mono);font-size:var(--fs-xs)' }, [buildId]),
  ));

  pane.appendChild(el('div', { class: 'tdr-hint', style: 'padding:var(--s-4);background:var(--paper-2);border-radius:var(--r-2);margin-top:var(--s-3)' }, [
    'リポジトリ: ',
    (() => {
      const a = el('a', { href: 'https://github.com/trie0000/tadori', target: '_blank', style: 'color:var(--accent-strong)' }, ['github.com/trie0000/tadori']);
      return a;
    })(),
    ' (社内 private)',
  ]));
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

  // 文字サイズ (小 / 中 / 大) — 即時反映。Spira と同じカード型ラジオ。
  pane.appendChild(el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-7)' }, ['文字サイズ']));
  pane.appendChild(el('p', { class: 'tdr-hint' }, ['ラジオを選んだ瞬間に反映され、この端末にだけ保存されます。']));
  const curSize = getFontSize();
  for (const opt of [
    { v: 'sm', label: '小', desc: '一覧で多くの行を見たい場合' },
    { v: 'md', label: '中 (既定)', desc: '標準サイズ — バランス重視' },
    { v: 'lg', label: '大', desc: '視認性重視・長時間作業向け' },
  ] as const) {
    pane.appendChild(radioCard({
      name: 'tdr-font-size', value: opt.v, checked: opt.v === curSize,
      title: opt.label, desc: opt.desc,
      onSelect: () => setFontSize(opt.v),
    }));
  }

  // チャット送信キー — 即時保存。
  pane.appendChild(el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-7)' }, ['チャット送信キー']));
  pane.appendChild(el('p', { class: 'tdr-hint' }, ['入力欄で Enter キーを押した時の挙動を選びます。']));
  const enterSends = loadSettings().enterSends;
  pane.appendChild(radioCard({
    name: 'tdr-enter-send', value: 'ctrl', checked: !enterSends,
    title: '⌘ / Ctrl + Enter で送信',
    desc: 'Enter は改行。長文を書きやすい (既定)',
    onSelect: () => { saveSettings({ enterSends: false }); toast(root, '⌘/Ctrl+Enter で送信に切替', 'ok'); },
  }));
  pane.appendChild(radioCard({
    name: 'tdr-enter-send', value: 'enter', checked: enterSends,
    title: 'Enter で送信',
    desc: 'Shift + Enter で改行。テンポよく投げたい人向け',
    onSelect: () => { saveSettings({ enterSends: true }); toast(root, 'Enter で送信に切替', 'ok'); },
  }));
}

/** Spira 流のカード型ラジオ (選択時に薄い背景強調 + タイトル + 説明)。 */
function radioCard(opts: {
  name: string; value: string; checked: boolean;
  title: string; desc: string;
  onSelect: () => void;
}): HTMLElement {
  const radio = el('input', { type: 'radio', name: opts.name, value: opts.value, style: 'margin:0;flex-shrink:0' }) as HTMLInputElement;
  if (opts.checked) radio.checked = true;
  const label = el('label', { class: 'tdr-radio-card' + (opts.checked ? ' is-checked' : '') }, [
    radio,
    el('div', { style: 'flex:1;min-width:0' }, [
      el('div', { class: 'tdr-radio-card-title' }, [opts.title]),
      el('div', { class: 'tdr-radio-card-desc' }, [opts.desc]),
    ]),
  ]);
  radio.addEventListener('change', () => {
    if (!radio.checked) return;
    // 同名グループの他カードから is-checked を外し、自分に付ける
    const root = label.closest('.tdr-hub-pane') ?? label.parentElement;
    if (root) for (const c of root.querySelectorAll<HTMLElement>(`.tdr-radio-card`)) {
      const r = c.querySelector<HTMLInputElement>(`input[name="${opts.name}"]`);
      if (r) c.classList.toggle('is-checked', r.checked);
    }
    opts.onSelect();
  });
  return label;
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

/** 用語辞書: 同義語グループを編集 (正式名 | 別名 | 意味)。SP の glossary.json に共有保存。 */
function buildGlossaryPane(pane: HTMLElement, root: HTMLElement, siteUrl: string): void {
  paneHead(pane, '用語辞書',
    '社内用語・略語の同義語を登録すると、検索が表記違いも拾います (例: 「P-WF」で「ワークフロー」もヒット)。意味は任意。SharePoint に保存しチームで共有します。');

  let entries: GlossaryEntry[] = loadGlossary(siteUrl);
  const tableWrap = el('div', { style: 'margin-top:var(--s-3)' });
  const status = el('div', { class: 'tdr-hint', style: 'margin-top:var(--s-3)' }, ['']);

  const render = (): void => {
    tableWrap.replaceChildren();
    const head = el('div', { style: 'display:grid;grid-template-columns:1fr 1.4fr 1.4fr 32px;gap:var(--s-2);font-size:var(--fs-xs);color:var(--ink-4);padding:0 2px var(--s-1)' }, [
      el('span', {}, ['正式名']), el('span', {}, ['別名 (カンマ区切り)']), el('span', {}, ['意味 (任意)']), el('span', {}, ['']),
    ]);
    tableWrap.appendChild(head);
    if (entries.length === 0) tableWrap.appendChild(el('div', { class: 'tdr-hint', style: 'padding:var(--s-3) 2px' }, ['まだ登録がありません。「行を追加」か Excel から貼り付けで登録してください。']));
    entries.forEach((e, i) => {
      const cIn = el('input', { class: 'tdr-input', value: e.canonical }) as HTMLInputElement;
      const aIn = el('input', { class: 'tdr-input', value: e.aliases.join(', ') }) as HTMLInputElement;
      const dIn = el('input', { class: 'tdr-input', value: e.def || '' }) as HTMLInputElement;
      cIn.addEventListener('input', () => { entries[i].canonical = cIn.value.trim(); });
      aIn.addEventListener('input', () => { entries[i].aliases = aIn.value.split(/[,;、；]/).map(s => s.trim()).filter(Boolean); });
      dIn.addEventListener('input', () => { entries[i].def = dIn.value.trim() || undefined; });
      const del = el('button', { class: 'tdr-btn tdr-btn--sm', title: '削除' }, ['×']);
      del.addEventListener('click', () => { entries.splice(i, 1); render(); });
      tableWrap.appendChild(el('div', { style: 'display:grid;grid-template-columns:1fr 1.4fr 1.4fr 32px;gap:var(--s-2);margin-bottom:var(--s-2)' }, [cIn, aIn, dIn, del]));
    });
  };
  render();

  const addBtn = el('button', { class: 'tdr-btn' }, ['+ 行を追加']);
  addBtn.addEventListener('click', () => { entries.push({ canonical: '', aliases: [] }); render(); });

  const saveBtn = el('button', { class: 'tdr-btn tdr-btn--primary' }, ['保存 (SPに共有)']);
  saveBtn.addEventListener('click', () => {
    saveBtn.disabled = true; status.textContent = '保存中…';
    void (async () => {
      try {
        await persistGlossary(siteUrl, entries);
        entries = loadGlossary(siteUrl); render();
        status.textContent = `保存しました (${entries.length} 件)`;
        toast(root, `用語辞書を保存しました (${entries.length} 件)`, 'ok');
      } catch (e) {
        status.textContent = '保存失敗';
        toast(root, `保存失敗: ${e instanceof Error ? e.message : String(e)}`, 'error');
      } finally { saveBtn.disabled = false; }
    })();
  });

  const reloadBtn = el('button', { class: 'tdr-btn' }, ['SPから再読込']);
  reloadBtn.addEventListener('click', () => {
    status.textContent = '読込中…';
    void (async () => { entries = await fetchGlossary(siteUrl); render(); status.textContent = `読み込みました (${entries.length} 件)`; })();
  });

  // Excel からの貼り付け (TSV/CSV) 取込。
  const pasteArea = el('textarea', { class: 'tdr-input', rows: '4', placeholder: 'Excel から貼り付け: 1列目=正式名, 2列目=別名(カンマ区切り), 3列目=意味(任意)', style: 'width:100%;margin-top:var(--s-3)' }) as HTMLTextAreaElement;
  const importBtn = el('button', { class: 'tdr-btn' }, ['貼り付けから取込 (追記)']);
  importBtn.addEventListener('click', () => {
    const add = parseGlossaryTable(pasteArea.value);
    if (add.length === 0) { toast(root, '取り込める行がありません', 'warn'); return; }
    // 同じ正式名は上書きマージ
    for (const a of add) {
      const idx = entries.findIndex(e => e.canonical && e.canonical === a.canonical);
      if (idx >= 0) entries[idx] = { ...entries[idx], aliases: [...new Set([...entries[idx].aliases, ...a.aliases])], def: a.def || entries[idx].def };
      else entries.push(a);
    }
    pasteArea.value = ''; render();
    toast(root, `${add.length} 行を取り込みました (保存で確定)`, 'ok');
  });

  pane.append(
    tableWrap,
    el('div', { style: 'display:flex;gap:var(--s-3);align-items:center;margin-top:var(--s-3)' }, [addBtn, saveBtn, reloadBtn]),
    el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-6)' }, ['Excel から一括取込']),
    pasteArea,
    el('div', { style: 'margin-top:var(--s-2)' }, [importBtn]),
    status,
  );

  // 起動キャッシュが古い可能性があるので SP から最新を取り直す。
  void fetchGlossary(siteUrl).then(fresh => { entries = fresh; render(); }).catch(() => { /* keep cache */ });
}

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

  // ベータ: ブラウザ常駐の自動取り込み (既定 OFF)。開発者モードのこのセクションでのみ操作可。
  pane.appendChild(el('p', { class: 'tdr-pane-title', style: 'margin-top:var(--s-7)' }, ['ベータ機能']));
  const autoCb = el('input', { type: 'checkbox' }) as HTMLInputElement;
  autoCb.checked = isAutoIngestFlagOn();
  autoCb.addEventListener('change', () => {
    setAutoIngestFlag(autoCb.checked);
    toast(root, autoCb.checked ? '自動取り込み ON (次回リロードで有効)' : '自動取り込み OFF (次回リロードで停止)', 'ok');
  });
  pane.appendChild(el('label', {
    style: 'display:inline-flex;align-items:center;gap:var(--s-3);cursor:pointer;padding:var(--s-3);background:var(--paper-2);border-radius:var(--r-2)',
  }, [autoCb, el('span', { style: 'font-size:var(--fs-md)' }, ['自動取り込み（常駐・ベータ）を有効にする'])]));
  pane.appendChild(el('p', { class: 'tdr-hint', style: 'margin-top:var(--s-3)' }, [
    '既定は OFF。OFF のときは「設定 → 取り込み」での明示同期だけが取り込み手段です。',
    'ON にすると、ハートビートで書き込み担当になったブラウザが新着を自動取り込みします (実験的)。',
    '※ 開発者モードを切ると自動で無効になります。変更は次回リロードで反映。',
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

// ─── 危険ゾーン ─────────────────────────────────────────────────────────────────

/** ペイン共通のデンジャーカード (Spira と同じ作法: 説明 + danger ボタン)。 */
function dangerCard(opts: {
  pane: HTMLElement; title: string; warning: string; buttonLabel: string; onRun: () => void;
}): void {
  paneHead(opts.pane, opts.title, '');
  const card = el('div', {
    style: 'border:1px solid var(--danger);border-radius:var(--r-3);padding:var(--s-5) var(--s-7);background:var(--danger-soft);margin-top:var(--s-3);',
  }, [
    el('p', { style: 'margin:0 0 var(--s-4);color:var(--danger);font-weight:600;white-space:pre-line;' }, ['⚠ ' + opts.warning]),
  ]);
  const btn = el('button', { class: 'tdr-btn tdr-btn--danger' }, [opts.buttonLabel]);
  btn.addEventListener('click', () => opts.onRun());
  card.appendChild(btn);
  opts.pane.appendChild(card);
}

function buildResetMailPane(pane: HTMLElement, root: HTMLElement, siteUrl: string): void {
  dangerCard({
    pane,
    title: '取り込みメールを全削除',
    warning:
      'ベクトルDB に登録された全メールを削除します (ブラウザのキャッシュと SharePoint 上の Tadori フォルダ配下の manifest.json / seg-*.json)。\n'
      + 'チャット履歴・設定は残ります。再取り込みすると最新の relay (conversationId 付き) で再構築されます。',
    buttonLabel: '取り込みメールを全削除する',
    onRun: () => {
      confirmModal({
        root, title: '取り込みメールを全削除', primaryLabel: '削除する', primaryVariant: 'danger',
        message: 'ベクトルDB の全メール (manifest + 全セグメント) を削除します。元に戻せません。よろしいですか?',
        onConfirm: async () => {
          try { await wipeImportedMails(siteUrl); toast(root, '取り込みメールを全削除しました。再取り込みしてください。', 'ok'); }
          catch (e) { toast(root, `削除失敗: ${e instanceof Error ? e.message : String(e)}`, 'error'); }
        },
      });
    },
  });
}

function buildResetAllPane(pane: HTMLElement, root: HTMLElement, siteUrl: string): void {
  dangerCard({
    pane,
    title: 'ツール全体をリセット',
    warning:
      '次のすべてを削除します:\n'
      + ' • 取り込みメール (ベクトルDB) — ローカル + SharePoint\n'
      + ' • Tadori と Spira の AI 設定 (localStorage)\n'
      + ' • チャット履歴・利用料の累計・サイドバー幅などのローカル状態\n'
      + 'リセット後はページが再読み込みされ、初期状態になります。元に戻せません。',
    buttonLabel: 'ツール全体をリセットする',
    onRun: () => {
      confirmModal({
        root, title: 'ツール全体をリセット', primaryLabel: 'リセットする', primaryVariant: 'danger',
        message: '本当に Tadori を完全リセットしますか? 設定・履歴・取り込みデータがすべて消えます。',
        onConfirm: async () => {
          try {
            await wipeImportedMails(siteUrl);
            // localStorage の Tadori 関連キーだけを掃除。
            // Tadori と Spira は AI 設定も独立管理なので、spira:* は触らない。
            try {
              const keys: string[] = [];
              for (let i = 0; i < localStorage.length; i++) {
                const k = localStorage.key(i); if (!k) continue;
                if (k.startsWith('tadori:') || k.startsWith('tadori.')) keys.push(k);
              }
              for (const k of keys) localStorage.removeItem(k);
            } catch { /* quota / noop */ }
            toast(root, 'リセットしました。再読み込みします…', 'ok');
            setTimeout(() => location.reload(), 800);
          } catch (e) { toast(root, `リセット失敗: ${e instanceof Error ? e.message : String(e)}`, 'error'); }
        },
      });
    },
  });
}

