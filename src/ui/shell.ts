// アプリシェル。#tadori-root を DOM に注入し、トップバー + チャットパネルをマウント。

import { el } from '../lib/dom';
import { icons } from './icons';
import { openSettingsHub } from './settingsHub';
import { createChatPanel } from './chat';
import { toast } from './toast';
import { resolveProvider } from '../api/aiSettings';
import { fetchLatestBuildId } from '../utils/bundleSource';
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

  const topbar = el('div', { class: 'tdr-topbar' }, [
    el('div', { class: 'tdr-brand' }, [
      el('span', { class: 'mark' }, ['辿']),
      el('span', { class: 'name' }, ['Tadori']),
      el('span', { class: 'sub' }, ['ML ナレッジサーチ']),
    ]),
    el('div', { class: 'tdr-spacer' }),
    moonBtn,
    settBtn,
    closeBtn,
  ]);

  root.append(topbar, createChatPanel(root, siteUrl));
  document.body.appendChild(root);

  // 起動時の更新検知: 前回見たビルドと違えば「更新されました」トースト。
  try {
    const last = localStorage.getItem(LAST_BUILD_KEY);
    if (last && last !== __TADORI_BUILD_ID__) {
      toast(root, `Tadori が最新版に更新されました (${__TADORI_BUILD_ID__})`, 'ok');
    }
    localStorage.setItem(LAST_BUILD_KEY, __TADORI_BUILD_ID__);
  } catch { /* noop */ }
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
