// Sticky モード: ハートビートで書き込み権を取れた時だけ Outlook の新着メールを
// 取り込むバックグラウンドジョブ。元の Phase 0 設計 (ADR-012 + Sticky モード) を
// 現アーキテクチャ (relay COM 取り込み) で復活させたもの。
//
// 流れ:
//   - app 起動時に startAutoIngest(siteUrl) を呼ぶ
//   - WriterLease.subscribe で writer 状態の変化を聞く
//   - writer = true になったら「最終取り込み完了日時以降」の新着を fetch → embed → segment
//   - 完了したら「最終取り込み完了日時」を更新 (Tadori Sync List に共有行で保存)
//   - writer = false に戻ったら進行中のジョブを abort
//   - 完了/失敗どちらでも、次の writer になるまでは何もしない
//
// relay が無い場合は fetch が失敗するので silently スキップする (エラー UI は出さない)。

import { getLease } from './lease';
import { SharePointClient } from '../sharepoint/client';
import { fetchOutlookMails, toIngestMails } from '../outlook/import';
import { ingestToSegments } from '../db/writer';
import { loadSettings, parseAddressList } from '../api/aiSettings';

const SYNC_LIST = 'Tadori Sync';
const LAST_INGEST_KEY = '__last_outlook_ingest__';

interface AutoIngestHandle {
  stop(): void;
}

/** Tadori Sync List の特別行に最終取り込み完了日時 (ISO) を保存/取得する。
 *  全メンバで共有 = 誰が次の writer になっても続きから取れる。 */
async function readLastIngestAt(sp: SharePointClient): Promise<string | null> {
  try {
    const rows = await sp.getItems(SYNC_LIST, `$select=Id,last_seen&$filter=Title eq '${LAST_INGEST_KEY}'&$top=1`);
    if (rows.length === 0) return null;
    const lastSeen = rows[0].last_seen as string | undefined;
    return lastSeen ?? null;
  } catch { return null; }
}

async function writeLastIngestAt(sp: SharePointClient, isoTs: string): Promise<void> {
  try {
    const rows = await sp.getItems(SYNC_LIST, `$select=Id&$filter=Title eq '${LAST_INGEST_KEY}'&$top=1`);
    if (rows.length === 0) {
      await sp.createItem(SYNC_LIST, { Title: LAST_INGEST_KEY, last_seen: isoTs });
    } else {
      await sp.updateItem(SYNC_LIST, Number(rows[0].Id), { last_seen: isoTs }, '*');
    }
  } catch (e) {
    console.warn('[autoIngest] last_ingest_at 書き込み失敗:', (e as Error).message);
  }
}

/** Sticky モードを開始。返り値の stop() で停止可能。 */
export function startAutoIngest(siteUrl: string): AutoIngestHandle {
  const lease = getLease(siteUrl);
  const sp = new SharePointClient(siteUrl);
  let running = false;
  let ac: AbortController | null = null;

  // writer になったら 1 度走らせる。失敗・終了したらフラグを戻す。
  const onStatus = (status: { isWriter: boolean }): void => {
    if (!status.isWriter) {
      // 書き込み権を失った → 進行中なら中断
      if (ac) ac.abort();
      return;
    }
    if (running) return; // 多重起動防止
    void runOnce();
  };

  async function runOnce(): Promise<void> {
    running = true;
    ac = new AbortController();
    try {
      const s = loadSettings();
      // 取り込み先 (relay) 未設定なら何もしない
      if (!s.relayBaseUrl) return;
      const to = s.mlAddresses ?? [];
      if (to.length === 0) return; // 取り込み対象アドレスが未指定

      // 最終取り込み日時を読み出し → 「それ以降〜現在」を取りに行く。
      // 初回 (last 無し) は安全側で「過去 7 日」だけ取る (大量取り込みは手動でやる前提)。
      const last = await readLastIngestAt(sp);
      const since = last ? last.slice(0, 10) : new Date(Date.now() - 7 * 86400_000).toISOString().slice(0, 10);
      const until = new Date().toISOString().slice(0, 10);

      const mails = await fetchOutlookMails(s.relayBaseUrl, {
        to, cc: parseAddressList(''),
        since, until,
        max: 500,
      }, ac.signal);

      if (ac.signal.aborted) return;
      if (mails.length === 0) {
        // 何もせずタイムスタンプだけ更新
        await writeLastIngestAt(sp, new Date().toISOString());
        return;
      }

      const r = await ingestToSegments(toIngestMails(mails), s, siteUrl, () => { /* silent */ }, ac.signal);
      if (!ac.signal.aborted && !r.cancelled) {
        await writeLastIngestAt(sp, new Date().toISOString());
        console.info(`[autoIngest] 新規 ${r.added} 件 / 重複 ${r.skipped} 件 (取得 ${mails.length} 件)`);
      }
    } catch (e) {
      if (ac?.signal.aborted) return; // 中断は正常
      console.warn('[autoIngest] 失敗:', (e as Error).message);
    } finally {
      running = false;
      ac = null;
    }
  }

  const unsubscribe = lease.subscribe(onStatus);
  return {
    stop(): void {
      unsubscribe();
      if (ac) ac.abort();
    },
  };
}
