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
  xlarge?: boolean; // OneNote 追記用: lg より広く + body スクロール + footer 常時表示
  bodyless?: boolean; // body を直接 modal に入れる (hub のように内部スクロール自前管理)
  onClose?: () => void; // ✕ / Esc / backdrop で閉じた時にも呼ばれる
}): ModalHandle {
  const closeBtn = el('button', { class: 'tdr-iconbtn', style: 'margin-left:auto', 'aria-label': '閉じる', html: icons.close() });
  const header = el('div', { class: 'tdr-modal-header' }, [
    el('h2', { class: 'tdr-modal-title' }, [opts.title]),
    closeBtn,
  ]);
  const modalChildren: HTMLElement[] = [header, opts.body];
  if (opts.footer) modalChildren.push(opts.footer);

  const sizeClass = opts.xlarge ? ' tdr-modal--xl' : opts.large ? ' tdr-modal--lg' : opts.small ? ' tdr-modal--sm' : '';
  const modal = el('div', { class: `tdr-modal${sizeClass}`, role: 'dialog', 'aria-modal': 'true' }, modalChildren);
  const backdrop = el('div', { class: 'tdr-backdrop' }, [modal]);

  let closed = false;
  function close() {
    if (closed) return;
    closed = true;
    backdrop.remove();
    document.removeEventListener('keydown', onKey);
    opts.onClose?.();
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

/** 中央表示の確認モーダル (Spira の confirmModal と同じ作法)。
 *  primary 押下で onConfirm。✕/Esc/backdrop/キャンセルで onCancel。 */
export function confirmModal(opts: {
  root: HTMLElement;
  title: string;
  message: string;
  primaryLabel?: string;
  cancelLabel?: string;
  primaryVariant?: 'primary' | 'danger';
  onConfirm: () => void | Promise<void>;
  onCancel?: () => void;
}): void {
  const body = el('div', { class: 'tdr-modal-body', style: 'white-space:pre-line;line-height:1.7' }, [opts.message]);
  const cancelBtn = el('button', { class: 'tdr-btn' }, [opts.cancelLabel ?? 'キャンセル']);
  const primaryBtn = el('button', { class: `tdr-btn tdr-btn--${opts.primaryVariant ?? 'primary'}` }, [opts.primaryLabel ?? 'OK']);
  const footer = el('div', { class: 'tdr-modal-footer' }, [cancelBtn, primaryBtn]);

  let confirmed = false;
  const handle = openModal({
    root: opts.root, title: opts.title, body, footer, small: true,
    onClose: () => { if (!confirmed) opts.onCancel?.(); },
  });
  cancelBtn.addEventListener('click', () => handle.close());
  primaryBtn.addEventListener('click', async () => {
    confirmed = true;
    handle.close();
    await opts.onConfirm();
  });
}
