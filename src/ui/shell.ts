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

  const siteUrl = window._spPageContextInfo?.webAbsoluteUrl ?? location.origin;
  initUsage(siteUrl);
  const root    = el('div', { id: 'tadori-root' });

  // テーマ復元
  if (localStorage.getItem('tadori:theme') === 'dark') root.dataset.theme = 'dark';

  // トップバーボタン
  const moonBtn  = el('button', { class: 'tdr-iconbtn', 'aria-label': 'テーマ切替', html: icons.moon()     });
  const settBtn  = el('button', { class: 'tdr-iconbtn', 'aria-label': '設定',       html: icons.settings() });
  const closeBtn = el('button', { class: 'tdr-iconbtn', 'aria-label': 'アプリを閉じる', title: 'アプリを閉じる', html: icons.door() });

  moonBtn.addEventListener('click', () => {
    const isDark = root.dataset.theme === 'dark';
    root.dataset.theme = isDark ? '' : 'dark';
    localStorage.setItem('tadori:theme', isDark ? '' : 'dark');
  });
  settBtn.addEventListener('click', e => { e.stopPropagation(); openSettingsMenu(root, settBtn, siteUrl); });
  closeBtn.addEventListener('click', () => root.remove());

  const userChip = createUserChip(siteUrl);

  const topbar = el('div', { class: 'tdr-topbar' }, [
    el('div', { class: 'tdr-brand' }, [
      el('span', { class: 'mark' }, ['辿']),
      el('span', { class: 'name' }, ['Tadori']),
      el('span', { class: 'sub' }, ['ML ナレッジサーチ']),
    ]),
    el('div', { class: 'tdr-spacer' }),
    userChip,
    moonBtn,
    settBtn,
    closeBtn,
  ]);

  root.append(topbar, createChatPanel(root, siteUrl));
  document.body.appendChild(root);
  applyFontSize();

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

// ログインユーザー表示 (Spira 同様)。アバター(イニシャル) + 名前のチップ。
// _spPageContextInfo を即時表示し、/_api/web/currentuser で名前を補正する。
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
