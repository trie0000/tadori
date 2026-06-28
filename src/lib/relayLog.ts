// 診断メッセージをローカル relay のコンソールに出す (ブラウザ Console が読みづらい時用)。
// loadSettings().relayBaseUrl 宛に fire-and-forget で POST。relay 未起動なら黙って無視。

import { loadSettings } from '../api/aiSettings';

export function relayLog(msg: string): void {
  let base = '';
  try { base = (loadSettings().relayBaseUrl || '').replace(/\/+$/, ''); } catch { /* noop */ }
  if (!base) return;
  try {
    void fetch(`${base}/tadori/log`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ msg }),
    }).catch(() => { /* relay 未起動等は無視 */ });
  } catch { /* noop */ }
}
