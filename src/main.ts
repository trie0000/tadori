// Tadori エントリポイント。SharePoint / OWA 上で bookmarklet から起動される。
// Phase 1 時点ではモジュール配線のみ。取り込みループ・検索 UI は後続フェーズ。

import { DEFAULT_CONFIG } from './config';
import { SharePointClient } from './sharepoint/client';

function detectSiteUrl(): string {
  const ctx = window._spPageContextInfo;
  return (ctx && ctx.webAbsoluteUrl) || location.origin;
}

export function boot(): void {
  const siteUrl = detectSiteUrl();
  const cfg = DEFAULT_CONFIG;
  const sp = new SharePointClient(siteUrl);

  // 現段階は配線確認のみ。pipeline / ui は後続で wire する。
  void sp;
  console.log(`[Tadori] build ${__TADORI_BUILD_ID__} booted on ${siteUrl}`, cfg);
}

boot();
