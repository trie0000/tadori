// アプリシェル。#tadori-root を DOM に注入し、トップバー + チャットパネルをマウント。

import { el } from '../lib/dom';
import { icons } from './icons';
import { openSettingsHub } from './settingsHub';
import { createChatPanel } from './chat';
import { toast } from './toast';
import { resolveProvider, loadSettings } from '../api/aiSettings';
import { fetchLatestBuildId } from '../utils/bundleSource';
import { initUsage } from '../usage/tracker';
import { applyFontSize } from '../utils/fontSize';
import { getLease, type LeaseStatus } from '../sync/lease';
import { startAutoIngest } from '../sync/autoIngest';
import { isAutoIngestEnabled } from '../utils/devMode';
import { ensureTadoriInboxList } from '../sync/inboxList';
import {
  getSelectedSiteUrl, setSelectedSiteUrl, detectCurrentSiteUrl,
  fetchSiteTitle, refreshRecentSiteTitle,
} from '../sharepoint/spSites';
import { openSiteSelectionModal } from './siteSelectionModal';
import cssText from '../styles/app.css';

const LAST_BUILD_KEY = 'tadori:last-build';

export function boot(): void {
  // CSS 注入 (冪等)
  if (!document.getElementById('tdr-style')) {
    const style = document.createElement('style');
    style.id    = 'tdr-style';
    style.textContent = cssText;
    document.head.appendChild(style);
  }

  // 既存インスタンスをトグルオフ
  if (document.getElementById('tadori-root')) {
    document.getElementById('tadori-root')!.remove();
    return;
  }

  // サイト決定 (Spira 流):
  //   - 前回サイト (tadori:selected-site-url) があれば、そのまま使う (静かに高速起動)
  //   - 無ければ起動中ページから推定 (_spPageContextInfo) + バックグラウンドで一度だけ保存
  //   - 別サイトに切り替えたい時はトップバーの「サイト切替」ボタンから
  let siteUrl = getSelectedSiteUrl();
  if (!siteUrl) {
    siteUrl = detectCurrentSiteUrl();
    // 初回起動時は今のサイトを保存して、次回以降は静かに起動するようにする
    if (siteUrl) {
      setSelectedSiteUrl(siteUrl, siteUrl);
      // タイトル取れたら recent を上書き (非同期)
      void fetchSiteTitle(siteUrl).then(t => { if (t) refreshRecentSiteTitle(siteUrl!, t); });
    }
  }
  initUsage(siteUrl);
  const root    = el('div', { id: 'tadori-root' });

  // テーマ復元
  if (localStorage.getItem('tadori:theme') === 'dark') root.dataset.theme = 'dark';

  // トップバーボタン
  const moonBtn  = el('button', { class: 'tdr-iconbtn', 'aria-label': 'テーマ切替', html: icons.moon()     });
  const settBtn  = el('button', { class: 'tdr-iconbtn', 'aria-label': '設定',       html: icons.settings() });
  const closeBtn = el('button', { class: 'tdr-iconbtn', 'aria-label': 'アプリを閉じる', title: 'アプリを閉じる', html: icons.door() });
  // SP サイト切替: 起動時は前回サイトをそのまま使うので、別サイトに切り替えたい時はこれ。
  // 押すとアクセス可能サイト一覧モーダルを開き、選択後に shell を再マウントする。
  const siteSwitchBtn = el('button', {
    class: 'tdr-iconbtn', 'aria-label': 'サイト切替',
    title: 'SP サイトを切替 (Tadori をマウントする対象サイトを変更)',
    html: icons.folder(),
  });

  moonBtn.addEventListener('click', () => {
    const isDark = root.dataset.theme === 'dark';
    root.dataset.theme = isDark ? '' : 'dark';
    localStorage.setItem('tadori:theme', isDark ? '' : 'dark');
  });
  settBtn.addEventListener('click', e => { e.stopPropagation(); openSettingsMenu(root, settBtn, siteUrl); });
  closeBtn.addEventListener('click', () => root.remove());
  siteSwitchBtn.addEventListener('click', () => {
    void openSiteSelectionModal().then((sel) => {
      if (!sel) return;
      // ページはリロードしない。Tadori ルートを撤去してから boot() を再実行。
      // localStorage の selected-site-url が既に更新されているので新サイトで起動する。
      root.remove();
      // 念のため shell の常駐タイマー (lease.start 等) を止める手段は現状無いので、
      // 同一ページ内で 2 つ動くのを避けるため即座にリブート。
      boot();
    });
  });

  // サイト切替時の表示用に、現在サイトのタイトルを小さく出す。
  const siteChip = createSiteChip(siteUrl);

  const userChip = createUserChip(siteUrl);
  const presenceChip = createPresenceChip(siteUrl);

  const topbar = el('div', { class: 'tdr-topbar' }, [
    el('div', { class: 'tdr-brand' }, [
      el('span', { class: 'mark' }, ['辿']),
      el('span', { class: 'name' }, ['Tadori']),
      el('span', { class: 'sub' }, ['ML ナレッジサーチ']),
    ]),
    el('div', { class: 'tdr-spacer' }),
    siteChip,
    presenceChip,
    userChip,
    siteSwitchBtn,
    moonBtn,
    settBtn,
    closeBtn,
  ]);

  root.append(topbar, createChatPanel(root, siteUrl));
  document.body.appendChild(root);
  applyFontSize();

  // 自動取り込み (Sticky モード) はベータ機能。開発者モード かつ ベータフラグ ON のときだけ起動。
  // 既定 (OFF) では常駐の自動取り込みは一切走らず、取り込みは「設定→取り込み」の明示同期のみ。
  // - 有効時: lease.start() で 30 秒毎にハートビート / リース更新し、writer になったら新着を取り込み
  // - relay 未起動なら autoIngest は silently スキップ (実害なし)
  if (isAutoIngestEnabled()) {
    void getLease(siteUrl).start();
    startAutoIngest(siteUrl);
  }
  // Tadori 受信メール List (PA からの投入先) が無ければ作る。失敗しても致命にしない。
  void ensureTadoriInboxList(siteUrl, loadSettings().listTitle).catch(e => {
    console.warn('[tadori] Tadori 受信メール List 自動作成失敗:', (e as Error).message);
  });

  // 起動時の更新検知: 前回見たビルドと違えば「更新されました」トースト。
  try {
    const last = localStorage.getItem(LAST_BUILD_KEY);
    if (last && last !== __TADORI_BUILD_ID__) {
      toast(root, `Tadori が最新版に更新されました (${__TADORI_BUILD_ID__})`, 'ok');
    }
    localStorage.setItem(LAST_BUILD_KEY, __TADORI_BUILD_ID__);
  } catch { /* noop */ }

  // 起動時に中継サーバの死活確認。落ちていれば起動を促す。
  void checkRelayAlive(root);
}

async function checkRelayAlive(root: HTMLElement): Promise<void> {
  const s = loadSettings();
  if (!s.relayBaseUrl) return; // claude only 等、relay 不要の構成は対象外
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 2500);
  try {
    const res = await fetch(`${s.relayBaseUrl.replace(/\/+$/, '')}/tadori/health`, { method: 'GET', signal: ctrl.signal });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
  } catch {
    toast(root, '中継サーバが起動していません。デスクトップの tadori-start から起動するか、tadori-ai-relay.ps1 を実行してください。', 'warn');
  } finally {
    clearTimeout(timer);
  }
}

/** 現在の SP サイトを示す小チップ (タイトル + URL ホバー)。 */
function createSiteChip(siteUrl: string): HTMLElement {
  // 初期表示は URL の末尾セグメントで仮埋め (例: /sites/foo/... → "foo")。
  // 起動後に fetchSiteTitle で取れたら正式名で上書き。
  const fallbackLabel = ((): string => {
    try {
      const u = new URL(siteUrl);
      const m = u.pathname.match(/^\/(?:sites|teams)\/([^/]+)/i);
      return m ? decodeURIComponent(m[1]) : u.host;
    } catch { return siteUrl; }
  })();
  const labelEl = el('span', { class: 'name' }, [fallbackLabel]);
  const chip = el('div', {
    class: 'tdr-site-chip', title: siteUrl, 'aria-label': '現在の SP サイト',
  }, [
    el('span', { class: 'ic', html: icons.folder(12) }),
    labelEl,
  ]);
  // 非同期にサイトタイトル取得 → 取れたら表示更新 + recent も上書き
  void fetchSiteTitle(siteUrl).then(t => {
    if (t) {
      labelEl.textContent = t;
      chip.title = `${t}\n${siteUrl}`;
      refreshRecentSiteTitle(siteUrl, t);
    }
  });
  return chip;
}

// ログインユーザー表示 (Spira 同様)。アバター(イニシャル) + 名前のチップ。
// _spPageContextInfo を即時表示し、/_api/web/currentuser で名前を補正する。
/** トップバーの在席+書き込み担当インジケータ。lease.subscribe で更新される。 */
function createPresenceChip(siteUrl: string): HTMLElement {
  const chip = el('div', { class: 'tdr-presence', title: 'クリックで詳細', 'aria-label': '在席状況' }, [
    el('span', { class: 'ic' }, ['👥']),
    el('span', { class: 'cnt' }, ['—']),
    el('span', { class: 'sep' }, ['·']),
    el('span', { class: 'writer' }, ['…']),
  ]);
  const update = (st: LeaseStatus): void => {
    const cnt = chip.querySelector<HTMLElement>('.cnt');
    const writer = chip.querySelector<HTMLElement>('.writer');
    if (cnt) cnt.textContent = String(st.peers.length || (st.holderId ? 1 : 0) || 0);
    if (writer) {
      const who = st.holderId || '—';
      // client-id (c-xxxx) はそのままだと無意味なので、自分は「あなた」、他者は短縮表記
      const shortWho = who === st.myId ? 'あなた' : (who.length > 10 ? who.slice(0, 8) + '…' : who);
      writer.textContent = `書込: ${shortWho}`;
      if (st.isWriter) writer.classList.add('is-me'); else writer.classList.remove('is-me');
    }
    // ホバー用 title に詳細
    const peerList = st.peers.map(p => `${p.id === st.myId ? '◎ ' : '  '}${p.id} (${secsAgo(p.lastSeen)})`).join('\n');
    chip.title = `在席 ${st.peers.length} 人\n書き込み担当: ${st.holderId || '未確定'}\n\n${peerList}`;
  };
  getLease(siteUrl).subscribe(update);
  return chip;
}

function secsAgo(ts: number): string {
  const s = Math.max(0, Math.round((Date.now() - ts) / 1000));
  if (s < 60) return `${s}秒前`;
  const m = Math.round(s / 60);
  return m < 60 ? `${m}分前` : `${Math.round(m / 60)}時間前`;
}

function createUserChip(siteUrl: string): HTMLElement {
  const ctx = (window as unknown as {
    _spPageContextInfo?: { userDisplayName?: string; userEmail?: string; userLoginName?: string };
  })._spPageContextInfo;
  let name = ctx?.userDisplayName || ctx?.userEmail || ctx?.userLoginName || '';

  const avatar = el('span', { class: 'tdr-user-avatar' }, [(name || '?').slice(0, 1).toUpperCase()]);
  const nameEl = el('span', { class: 'tdr-user-name' }, [name || 'ログイン情報なし']);
  const chip = el('div', { class: 'tdr-user', title: name || 'ログイン情報なし' }, [avatar, nameEl]);

  // currentuser で表示名を補正 (取れなければ _spPageContextInfo のまま)。
  void (async () => {
    try {
      const res = await fetch(`${siteUrl}/_api/web/currentuser?$select=Title,Email,LoginName`, {
        headers: { Accept: 'application/json;odata=nometadata' }, credentials: 'include',
      });
      if (!res.ok) return;
      const u = await res.json() as { Title?: string; Email?: string; LoginName?: string };
      const refined = u.Title || u.Email || u.LoginName || '';
      if (!refined) return;
      name = refined;
      nameEl.textContent = refined;
      avatar.textContent = refined.slice(0, 1).toUpperCase();
      chip.title = u.Email && u.Title ? `${u.Title} <${u.Email}>` : refined;
    } catch { /* 取得不能はフォールバック表示のまま */ }
  })();

  return chip;
}

// 歯車のドロップダウンメニュー (Spira 同様)。build 表示 + 設定 + 更新確認。
function openSettingsMenu(root: HTMLElement, anchor: HTMLElement, siteUrl: string): void {
  const existing = root.querySelector('.tdr-settings-menu');
  if (existing) { existing.remove(); return; } // トグル

  const provLabel = el('div', {
    class: 'tdr-menu-item',
    style: 'cursor:default;color:var(--ink-3);font-size:var(--fs-xs);pointer-events:none',
  }, [`provider: ${resolveProvider()}`]);

  const buildLabel = el('div', {
    class: 'tdr-menu-item',
    style: 'cursor:pointer;color:var(--ink-3);font-size:var(--fs-xs);font-family:var(--font-mono);white-space:normal;word-break:break-all;line-height:1.4',
    title: 'クリックでコピー',
  }, [`build: ${__TADORI_BUILD_ID__}`]);
  buildLabel.addEventListener('click', () => {
    void navigator.clipboard?.writeText(__TADORI_BUILD_ID__).then(() => {
      buildLabel.textContent = '✓ コピーしました';
      setTimeout(() => { buildLabel.textContent = `build: ${__TADORI_BUILD_ID__}`; }, 1200);
    }).catch(() => { /* noop */ });
  });

  const settingsItem = el('div', { class: 'tdr-menu-item' }, [
    el('span', { html: icons.settings(14) }), '設定',
  ]);
  settingsItem.addEventListener('click', () => { menu.remove(); openSettingsHub(root, siteUrl); });

  const updateItem = el('div', { class: 'tdr-menu-item' }, [
    el('span', { html: icons.activity(14) }), '更新を確認',
  ]);
  updateItem.addEventListener('click', () => {
    menu.remove();
    void (async () => {
      const latest = await fetchLatestBuildId();
      if (!latest) { toast(root, '更新元に接続できませんでした', 'warn'); return; }
      if (latest === __TADORI_BUILD_ID__) { toast(root, '最新です', 'ok'); return; }
      if (confirm(`新しいバージョンがあります。\n\n新: ${latest}\n現在: ${__TADORI_BUILD_ID__}\n\n再読み込みして更新しますか?`)) {
        location.reload();
      }
    })();
  });

  const menu = el('div', {
    class: 'tdr-menu tdr-settings-menu',
    style: 'position:fixed;z-index:2147483700;min-width:240px',
  }, [provLabel, buildLabel, el('div', { class: 'tdr-menu-divider' }), settingsItem, updateItem]);

  const rect = anchor.getBoundingClientRect();
  menu.style.top = `${rect.bottom + 4}px`;
  menu.style.right = `${window.innerWidth - rect.right}px`;
  root.appendChild(menu);

  setTimeout(() => {
    const closer = (ev: Event) => {
      if (!menu.contains(ev.target as Node)) {
        menu.remove();
        document.removeEventListener('click', closer);
      }
    };
    document.addEventListener('click', closer);
  }, 0);
}
