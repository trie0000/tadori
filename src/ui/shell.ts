// アプリシェル。#tadori-root を DOM に注入し、トップバー + チャットパネルをマウント。

import { el } from '../lib/dom';
import { icons } from './icons';
import { openSettingsHub } from './settingsHub';
import { createChatPanel } from './chat';
import cssText from '../styles/app.css';

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
  const closeBtn = el('button', { class: 'tdr-iconbtn', 'aria-label': '閉じる',     html: icons.close()    });

  moonBtn.addEventListener('click', () => {
    const isDark = root.dataset.theme === 'dark';
    root.dataset.theme = isDark ? '' : 'dark';
    localStorage.setItem('tadori:theme', isDark ? '' : 'dark');
  });
  settBtn.addEventListener('click',  () => openSettingsHub(root, siteUrl));
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
}
