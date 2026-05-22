// 新方式の検索エントリ。SharePoint のセグメントをブラウザ内ベクトルDBへ同期し、
// クエリを埋め込み → 総当たり cosine で Top-K。List 列方式 (旧 searchMails) の置換。

import { SharePointClient } from '../sharepoint/client';
import { SpVectorStore } from '../sync/spStore';
import { SegmentCache } from '../db/cache';
import { VectorDb } from '../db/store';
import { VectorSync } from '../sync/sync';
import { embedQueryFor } from '../embeddings/router';
import type { RuntimeSettings } from '../api/aiSettings';
import type { MailHit } from './index';

let db: VectorDb | null = null;
let sync: VectorSync | null = null;
let syncedSite = '';

/** 初回はセグメントを同期してから検索。siteUrl が変わったら作り直す。 */
async function ensureSynced(siteUrl: string): Promise<VectorDb> {
  if (db && syncedSite === siteUrl) return db;
  const fresh = new VectorDb();
  const store = new SpVectorStore(new SharePointClient(siteUrl));
  sync = new VectorSync(store, new SegmentCache(), fresh);
  await sync.sync();
  db = fresh;
  syncedSite = siteUrl;
  return db;
}

/** 手動再同期 (取り込み後など)。 */
export async function resyncVectors(siteUrl: string): Promise<void> {
  if (db && sync && syncedSite === siteUrl) { await sync.sync(); return; }
  await ensureSynced(siteUrl);
}

export async function searchVectors(
  question: string,
  s: RuntimeSettings,
  siteUrl: string,
  topK: number,
): Promise<MailHit[]> {
  const vdb = await ensureSynced(siteUrl);
  if (vdb.size === 0) return [];
  const qvec = await embedQueryFor(question, s);
  return vdb.search(qvec, topK).map(({ record, score }) => ({
    messageId: record.messageId,
    subject: record.subject,
    from: record.from,
    date: record.date,
    body: record.body,
    score,
  }));
}
