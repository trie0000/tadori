// 検索の結線: SharePoint List から埋め込み済み行を取得 → クエリを埋め込み →
// cosine Top-K。IndexedDB キャッシュは後続フェーズ (今は都度 List から取得)。

import type { RuntimeSettings } from '../api/aiSettings';
import { SharePointClient } from '../sharepoint/client';
import { embedQueryFor } from '../embeddings/router';
import { decodeEmbedding } from '../lib/float16';
import { normalize, search as cosineSearch, type IndexedVector } from './cosine';
import { COLUMNS, TADORI_LIST_FIELDS } from '../config';

export interface MailHit {
  messageId: string;
  subject: string;
  from: string;
  date: string;
  body: string;
  score: number;
}

interface IndexedMail extends IndexedVector {
  meta: { subject: string; from: string; date: string; body: string };
}

/** List から embedding 列が入っている行を読み、検索インデックスを構築。 */
async function loadIndex(sp: SharePointClient, s: RuntimeSettings): Promise<IndexedMail[]> {
  // 注意: embedding は Note(複数行テキスト)列で SharePoint の $filter は使えない。
  // フィルタ可能な embedded_at(日時)で「埋め込み済み」を絞り、空 embedding は
  // 後段でクライアント側スキップする。
  const sel = ['Id', 'Title', 'Created', COLUMNS.embedding, 'Body', 'From'].join(',');
  const rows = await sp.getItems(
    s.listTitle,
    `$select=${sel}&$filter=${COLUMNS.embeddedAt} ne null&$top=5000`,
  );
  const out: IndexedMail[] = [];
  for (const r of rows) {
    const b64 = r[COLUMNS.embedding];
    if (typeof b64 !== 'string' || !b64) continue;
    out.push({
      messageId: String(r.Id),
      vec: normalize(decodeEmbedding(b64)),
      meta: {
        subject: String(r.Title ?? '(件名なし)'),
        from: String((r as Record<string, unknown>).From ?? ''),
        date: String(r.Created ?? ''),
        body: String((r as Record<string, unknown>).Body ?? ''),
      },
    });
  }
  return out;
}

/** getItems が「リストが存在しない」で失敗したかを判定。 */
function isListMissing(e: unknown): boolean {
  const msg = e instanceof Error ? e.message : String(e);
  return /HTTP 404/.test(msg) || msg.includes('存在しません') || /does ?n.t exist/i.test(msg);
}

export async function searchMails(
  question: string,
  s: RuntimeSettings,
  siteUrl: string,
  topK: number,
): Promise<MailHit[]> {
  const sp = new SharePointClient(siteUrl);

  let index: IndexedMail[];
  try {
    index = await loadIndex(sp, s);
  } catch (e) {
    // リストが存在しない (404) なら自動作成し、データ投入を促す。
    if (isListMissing(e)) {
      await sp.ensureList(s.listTitle, TADORI_LIST_FIELDS);
      throw new Error(
        `リスト「${s.listTitle}」が無かったので作成しました。` +
        `設定 → 開発者 → サンプル投入、または取り込みでデータを入れてから再検索してください。`,
      );
    }
    throw e;
  }
  if (index.length === 0) return [];

  const qvec = normalize(await embedQueryFor(question, s));

  const hits = cosineSearch(qvec, index, topK);
  const byId = new Map(index.map(m => [m.messageId, m]));
  return hits.map(h => {
    const m = byId.get(h.messageId)!;
    return { messageId: h.messageId, score: h.score, ...m.meta };
  });
}
