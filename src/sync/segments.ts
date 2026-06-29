// ベクトルDBの配布フォーマット (SharePoint 上に置く)。
//
// 設計 (このリポジトリの設計ログ参照):
//   - ベクトルDB 本体は各 relay のローカル SQLite。SharePoint には「追記専用の
//     セグメント群 + manifest」を置き、各 relay が差分だけ取り込んで同期する。
//   - セグメントは封印後 不変。更新/削除は新しいレコード (op) を後続セグメントへ
//     追記し、seq 昇順・message-id 単位の last-writer-wins でローカルに収束させる。
//   - 1 セグメント ≒ 1,000 件で封印・ロール。容量回収はコンパクションで世代更新。

export type SegmentOp = 'upsert' | 'delete';

export interface SegmentRecord {
  /** 全体で単調増加する適用順。message-id 単位で最大 seq が勝つ。 */
  seq: number;
  op: SegmentOp;
  messageId: string;
  // op='upsert' のときのみ以下を持つ (delete は messageId だけの tombstone)。
  /** RFC2822 Internet-Message-Id (Outlook での再検索キー)。無い場合あり。 */
  internetMessageId?: string;
  /** スレッド識別子 (Outlook ConversationID)。OneNote では parentDocId (pageId) を兼用。 */
  conversationId?: string;
  /** ソース種別。'mail' (既定) / 'onenote' / 'doc' / 'pptx' / 'transcript'。 */
  kind?: 'mail' | 'onenote' | 'doc' | 'pptx' | 'transcript';
  /** 取り込みバッチのラベル (主に OneNote のラベル付きバッチ用)。検索のサブ項目絞り込みに使う。
   *  doc/pptx/transcript は folder URL で絞れるので通常は未設定。 */
  label?: string;
  /** PPTX 取り込みのメタ (kind='pptx' のときだけ意味を持つ)。 */
  pptxFile?: string;
  pptxServerRelUrl?: string;
  slideNo?: number;
  slideTitle?: string;
  thumbServerRelUrl?: string;
  /** Teams 文字起こしのメタ (kind='transcript' のときだけ意味を持つ)。 */
  transcriptFile?: string;
  vttServerRelUrl?: string;
  recordingServerRelUrl?: string;
  startSec?: number;
  /** ドキュメント (kind='doc') のメタ。 */
  docFile?: string;
  docServerRelUrl?: string;
  /** ソース内容のハッシュ (pptx: title+rawText+tables+notes / transcript: 本文)。
   *  再取り込み時に「内容が変わっていないチャンクは embed をスキップ」する
   *  差分判定に使う。 */
  srcHash?: string;
  /** 親ドキュメント内のチャンク番号 (0 始まり)。単一レコードならどちらも省略。 */
  chunkIdx?: number;
  /** 親ドキュメントの総チャンク数。 */
  chunkCount?: number;
  /** ドキュメントの擬似パス (例: "onenote://Notebook/Section/Page" / "/sites/x/.../file.docx")。 */
  docPath?: string;
  subject?: string;
  from?: string;
  to?: string[];
  cc?: string[];
  /** 受信日時 (ISO)。SharePoint の Created とは別に実受信日を保持。 */
  date?: string;
  /** 表示用本文。isHtml なら HTML、そうでなければプレーンテキスト。 */
  body?: string;
  /** body が HTML 形式か。true なら表示時に HTML レンダリング。 */
  isHtml?: boolean;
  /** 埋め込みベクトル (Base64 Float16, src/lib/float16.ts)。 */
  emb?: string;
}

export interface Segment {
  id: string;
  generation: number;
  records: SegmentRecord[];
}

export interface ManifestOpen {
  id: string;
  hash: string;
  count: number;
}

export interface Manifest {
  /** 単調増加。条件付 GET の補助・変化検知に使う。 */
  version: number;
  /** コンパクション世代。変わったら relay は集合を作り直す。 */
  generation: number;
  /** 全レコードを通して単調増加する適用順の現在値 (次は maxSeq+1)。 */
  maxSeq: number;
  /** 封印済みセグメント id (不変・一度DLしたら再取得しない)。 */
  sealed: string[];
  /** 追記中セグメント (小・hash 変化で再取得)。無ければ null。 */
  open: ManifestOpen | null;
  /** 最終更新 ISO。 */
  updatedAt: string;
}

export const SEGMENT_CAP = 1000; // 1 セグメントの最大レコード数 (封印トリガ)

export function emptyManifest(): Manifest {
  return { version: 0, generation: 1, maxSeq: 0, sealed: [], open: null, updatedAt: new Date().toISOString() };
}

/** sealed の `seg-NNNNN` から次のセグメント番号を返す (無ければ 1)。 */
export function nextSegmentIndex(sealed: string[]): number {
  let max = 0;
  for (const id of sealed) {
    const m = /(\d+)$/.exec(id);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return max + 1;
}

export function segmentId(index: number): string {
  return 'seg-' + String(index).padStart(5, '0');
}

export function serializeSegment(seg: Segment): string {
  return JSON.stringify(seg);
}

export function parseSegment(text: string): Segment {
  const o = JSON.parse(text) as Segment;
  if (!o || !Array.isArray(o.records)) throw new Error('壊れたセグメント');
  return o;
}

export function serializeManifest(m: Manifest): string {
  return JSON.stringify(m);
}

export function parseManifest(text: string): Manifest {
  const o = JSON.parse(text) as Manifest;
  if (!o || !Array.isArray(o.sealed)) throw new Error('壊れた manifest');
  return o;
}

/** open セグメントの変化検知用の軽量ハッシュ (FNV-1a 32bit)。 */
export function hashString(s: string): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
  }
  return h.toString(16).padStart(8, '0');
}

/** relay が持っていない封印セグメント id を返す (差分DLの対象)。 */
export function missingSealed(manifest: Manifest, haveIds: ReadonlySet<string>): string[] {
  return manifest.sealed.filter(id => !haveIds.has(id));
}
