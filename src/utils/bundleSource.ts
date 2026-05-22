// バンドル本体 (tadori.bundle.js) の読み込み元を開発者モードで切り替える設定。
//
// ローダー (build.js が生成する dist/tadori.loader.js / bookmarklet) が起動時に
// 下記 localStorage キーを読んで、本体を取りに行く base を決める:
//   - bundle-source = 'local'   → ローカル relay (local-base) から読む
//   - それ以外 / 未設定          → SharePoint の Tadori フォルダから読む (本番)
// 設定変更は「次回起動 / リロード」で反映される (ローダーは起動時に1度だけ走るため)。
//
// ※ キー名と既定 URL は build.js のローダー生成側と完全一致させること。

export const BUNDLE_SOURCE_KEY = 'tadori.dev.bundle-source';   // 'local' | (未設定=sharepoint)
export const BUNDLE_LOCAL_BASE_KEY = 'tadori.dev.local-base';  // 例: http://127.0.0.1:18080/tadori
export const DEFAULT_LOCAL_BASE = 'http://127.0.0.1:18080/tadori';

export type BundleSource = 'sharepoint' | 'local';

export function getBundleSource(): BundleSource {
  try {
    return localStorage.getItem(BUNDLE_SOURCE_KEY) === 'local' ? 'local' : 'sharepoint';
  } catch { return 'sharepoint'; }
}

export function setBundleSource(v: BundleSource): void {
  try {
    if (v === 'local') localStorage.setItem(BUNDLE_SOURCE_KEY, 'local');
    else localStorage.removeItem(BUNDLE_SOURCE_KEY);
  } catch (e) { console.warn('[tadori/bundleSource] localStorage 書込失敗:', (e as Error).message); }
}

export function getLocalBase(): string {
  try { return localStorage.getItem(BUNDLE_LOCAL_BASE_KEY) || DEFAULT_LOCAL_BASE; }
  catch { return DEFAULT_LOCAL_BASE; }
}

export function setLocalBase(url: string): void {
  try {
    const v = url.trim().replace(/\/+$/, '');
    if (v && v !== DEFAULT_LOCAL_BASE) localStorage.setItem(BUNDLE_LOCAL_BASE_KEY, v);
    else localStorage.removeItem(BUNDLE_LOCAL_BASE_KEY);
  } catch (e) { console.warn('[tadori/bundleSource] localStorage 書込失敗:', (e as Error).message); }
}

/** ローカル relay の origin (例 http://127.0.0.1:18080) を local-base から導出。 */
function localRelayOrigin(): string {
  try { return new URL(getLocalBase()).origin; } catch { return 'http://127.0.0.1:18080'; }
}

export interface RelayBundleDir { dir: string; exists: boolean; hasBundle: boolean }

/** relay が現在配信しているフォルダ (bundle dir) を取得。relay 不在時 null。 */
export async function getRelayBundleDir(): Promise<RelayBundleDir | null> {
  try {
    const res = await fetch(`${localRelayOrigin()}/tadori/bundle-dir`);
    if (!res.ok) return null;
    const j = await res.json() as Partial<RelayBundleDir>;
    return { dir: String(j.dir ?? ''), exists: !!j.exists, hasBundle: !!j.hasBundle };
  } catch { return null; }
}

/** relay の配信フォルダを変更する。成功時は変更後の状態を返す。 */
export async function setRelayBundleDir(dir: string): Promise<RelayBundleDir | null> {
  try {
    const res = await fetch(`${localRelayOrigin()}/tadori/bundle-dir`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ dir }),
    });
    if (!res.ok) return null;
    const j = await res.json() as Partial<RelayBundleDir>;
    return { dir: String(j.dir ?? ''), exists: !!j.exists, hasBundle: !!j.hasBundle };
  } catch { return null; }
}
