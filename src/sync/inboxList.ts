// Tadori 受信メール List の定義と自動作成。
// Power Automate からの自動投入先 (新着メールを Tadori が拾うためのバッファ)。
//
// このリストに PA が新着メールを書き込み、Tadori (writer 担当) がポーリングで
// 拾って embed → ベクトル DB へ投入する。relay COM 取り込みと並列に動かせる。

import { SharePointClient, type FieldSpec } from '../sharepoint/client';

/** Tadori 受信メール List の表示名 (既定)。Spira/Shapion 系のリスト一覧でも一目で
 *  Tadori 関連と分かる名前に。 */
export const TADORI_INBOX_LIST = 'Tadori 受信メール';

/** PA が書き込む列。型は SharePoint REST の FieldType に対応。
 *  Title は標準列なので追加不要、それ以外は ensureFields で作る。 */
export const TADORI_INBOX_FIELDS: FieldSpec[] = [
  { name: 'From',         type: 'text'     },
  { name: 'ToAddrs',      type: 'note'     },  // 改行区切り (To 列名は SP の予約と被るので別名)
  { name: 'CcAddrs',      type: 'note'     },
  { name: 'ReceivedTime', type: 'datetime' },
  { name: 'MessageId',    type: 'text'     },
  { name: 'Body',         type: 'note'     },
  { name: 'IsHtml',       type: 'boolean'  },
  { name: 'TadoriStatus', type: 'text'     },  // 'pending' / 'embedded' / 'error' (Tadori 側で更新)
];

/** List 名は設定の draft.listTitle で上書き可能。未設定なら既定名。 */
export function tadoriInboxListName(custom?: string | null): string {
  const t = (custom || '').trim();
  return t || TADORI_INBOX_LIST;
}

/** リストが無ければ作成、不足列があれば追加。Tadori Sync と違いバージョン履歴は
 *  ON のままにする (メール本文の改変履歴が残るのは妥当)。 */
export async function ensureTadoriInboxList(siteUrl: string, listTitle?: string): Promise<void> {
  const sp = new SharePointClient(siteUrl);
  const name = tadoriInboxListName(listTitle);
  await sp.ensureList(name, TADORI_INBOX_FIELDS);
}
