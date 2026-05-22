// モーダル基盤。backdrop クリック / Esc で閉じる。root 要素配下にマウント。
import { el } from '../lib/dom';
import { icons } from './icons';

export interface ModalHandle {
  el: HTMLElement;
  close: () => void;
}

export function openModal(opts: {
  root: HTMLElement;
  title: string;
  body: HTMLElement;
  footer?: HTMLElement;
  small?: boolean;
  large?: boolean; // 設定ハブ用: 項目で大きさが変わらない固定サイズ
  bodyless?: boolean; // body を直接 modal に入れる (hub のように内部スクロール自前管理)
}): ModalHandle {
  const closeBtn = el('button', { class: 'tdr-iconbtn', style: 'margin-left:auto', 'aria-label': '閉じる', html: icons.close() });
  const header = el('div', { class: 'tdr-modal-header' }, [
    el('h2', { class: 'tdr-modal-title' }, [opts.title]),
    closeBtn,
  ]);
  const modalChildren: HTMLElement[] = [header, opts.body];
  if (opts.footer) modalChildren.push(opts.footer);

  const sizeClass = opts.large ? ' tdr-modal--lg' : opts.small ? ' tdr-modal--sm' : '';
  const modal = el('div', { class: `tdr-modal${sizeClass}`, role: 'dialog', 'aria-modal': 'true' }, modalChildren);
  const backdrop = el('div', { class: 'tdr-backdrop' }, [modal]);

  function close() {
    backdrop.remove();
    document.removeEventListener('keydown', onKey);
  }
  function onKey(e: KeyboardEvent) { if (e.key === 'Escape') close(); }

  closeBtn.addEventListener('click', close);
  backdrop.addEventListener('click', e => { if (e.target === backdrop) close(); });
  document.addEventListener('keydown', onKey);

  opts.root.appendChild(backdrop);
  const firstInput = modal.querySelector<HTMLElement>('input, textarea, select');
  if (firstInput) firstInput.focus();

  return { el: modal, close };
}
