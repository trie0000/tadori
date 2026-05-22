// 新方式の検索エントリ。SharePoint のセグメントをブラウザ内ベクトルDBへ同期し、
// クエリを埋め込み → 総当たり cosine で Top-K。List 列方式 (旧 searchMails) の置換。

import { getEngine } from '../db/engine';
import { embedQueryFor } from '../embeddings/router';
import type { RuntimeSettings } from '../api/aiSettings';

export interface MailHit {
  messageId: string;
  subject: string;
  from: string;
  date: string;
  body: string;
  score: number;
}

/** 手動再同期 (取り込み後など)。 */
export async function resyncVectors(siteUrl: string): Promise<void> {
  const eng = await getEngine(siteUrl);
  await eng.sync.sync();
}

export async function searchVectors(
  question: string,
  s: RuntimeSettings,
  siteUrl: string,
  topK: number,
): Promise<MailHit[]> {
  const eng = await getEngine(siteUrl);
  if (eng.db.size === 0) return [];
  const qvec = await embedQueryFor(question, s);
  return eng.db.search(qvec, topK).map(({ record, score }) => ({
    messageId: record.messageId,
    subject: record.subject,
    from: record.from,
    date: record.date,
    body: record.body,
    score,
  }));
}
