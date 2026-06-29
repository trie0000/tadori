// 開発者モード — localStorage の単純フラグ。
// 有効時のみ Claude API / Voyage 直接利用などの実験的設定を表示する。
// (Spira の src/utils/devMode.ts と同じ流儀。Tadori は独自キー。)

const KEY = 'tadori:developer-mode';

export function isDeveloperMode(): boolean {
  try {
    return localStorage.getItem(KEY) === '1';
  } catch {
    return false;
  }
}

export function setDeveloperMode(v: boolean): void {
  try {
    if (v) localStorage.setItem(KEY, '1');
    else localStorage.removeItem(KEY);
  } catch (e) {
    console.warn('[tadori/devMode] localStorage 書込失敗:', (e as Error).message);
  }
}

// ── ベータ機能: ブラウザ常駐の自動取り込み (ハートビートで writer になったら自動同期) ──
// 既定 OFF。開発者モード時のみ設定 UI に出す。通常運用は「設定→取り込み」での明示同期のみ。
const BETA_AUTO_INGEST_KEY = 'tadori:beta:auto-ingest';

/** ベータ自動取り込みフラグ単体の状態 (UI トグル用)。 */
export function isAutoIngestFlagOn(): boolean {
  try { return localStorage.getItem(BETA_AUTO_INGEST_KEY) === '1'; } catch { return false; }
}

export function setAutoIngestFlag(v: boolean): void {
  try {
    if (v) localStorage.setItem(BETA_AUTO_INGEST_KEY, '1');
    else localStorage.removeItem(BETA_AUTO_INGEST_KEY);
  } catch (e) {
    console.warn('[tadori/devMode] localStorage 書込失敗:', (e as Error).message);
  }
}

/** 自動取り込みを実際に動かしてよいか。開発者モード かつ ベータフラグ ON のときだけ true。 */
export function isAutoIngestEnabled(): boolean {
  return isDeveloperMode() && isAutoIngestFlagOn();
}
