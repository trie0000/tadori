// 新方式の検索エントリ。SharePoint のセグメントをブラウザ内ベクトルDBへ同期し、
// クエリを埋め込み → 総当たり cosine で Top-K。List 列方式 (旧 searchMails) の置換。

import { getEngine } from '../db/engine';
import { embedQueryFor } from '../embeddings/router';
import type { MailRecord } from '../db/store';
import type { RuntimeSettings } from '../api/aiSettings';

export interface MailHit {
  messageId: string;
  internetMessageId: string;
  conversationId: string;
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  date: string;
  body: string;
  isHtml: boolean;
  score: number;
}

function toHit(record: MailRecord, score: number): MailHit {
  return {
    messageId: record.messageId,
    internetMessageId: record.internetMessageId,
    conversationId: record.conversationId,
    subject: record.subject,
    from: record.from,
    to: record.to,
    cc: record.cc,
    date: record.date,
    body: record.body,
    isHtml: record.isHtml,
    score,
  };
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
  return eng.db.search(qvec, topK, question, s.ragKeywordWeight).map(({ record, score }) => toHit(record, score));
}

/** 同一スレッド (conversationId) の全メールを時系列で返す (経緯要約用)。 */
export async function getThread(siteUrl: string, conversationId: string): Promise<MailHit[]> {
  const eng = await getEngine(siteUrl);
  return eng.db.byConversation(conversationId).map(r => toHit(r, 1));
}
