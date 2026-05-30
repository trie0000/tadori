// チャット履歴の永続化 (localStorage)。左ペインのセッション一覧 + 再表示用。
// 1 セッション = 1 スレッド。turns に質問/回答(Markdown)/出典メタを保存する。

export interface SavedHit {
  messageId: string;
  internetMessageId: string;
  conversationId: string;
  kind: 'mail' | 'onenote' | 'doc' | 'pptx';
  chunkIdx?: number;
  chunkCount?: number;
  docPath?: string;
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  date: string;
  body: string;
  isHtml: boolean;
  score: number;
  /** PPTX 関連メタ (kind='pptx' のときのみ意味を持つ)。 */
  pptxFile?: string;
  pptxServerRelUrl?: string;
  slideNo?: number;
  slideTitle?: string;
  thumbServerRelUrl?: string;
}

export interface SavedTurn {
  q: string;
  /** 回答の生 Markdown (再表示時に renderMarkdown する)。 */
  answer: string;
  hits: SavedHit[];
  ms: number;
  /** このやり取りで発生した AI 利用料の目安 (円)。古い履歴では未設定。 */
  yen?: number;
  /** 質問が送信された時刻 (ISO 8601)。古い履歴では未設定。 */
  createdAt?: string;
}

export interface ChatSession {
  id: string;
  title: string;
  updatedAt: string;
  turns: SavedTurn[];
}

// チャット履歴はサイト別に分離。旧 'tadori:chat-sessions' グローバルキーは
// 初回読込時に現在サイトへ移管 (1 回だけ)。
import { siteHash } from '../sharepoint/spSites';

const LEGACY_KEY = 'tadori:chat-sessions';
const MIGRATED_KEY = 'tadori:chat-sessions:legacy-migrated';
const MAX_SESSIONS = 50;

function keyFor(siteUrl: string): string {
  return `tadori:chat-sessions:${siteHash(siteUrl)}`;
}

function migrateLegacy(siteUrl: string): void {
  try {
    if (localStorage.getItem(MIGRATED_KEY)) return;
    const legacy = localStorage.getItem(LEGACY_KEY);
    if (!legacy) { localStorage.setItem(MIGRATED_KEY, '1'); return; }
    const cur = localStorage.getItem(keyFor(siteUrl));
    if (!cur) localStorage.setItem(keyFor(siteUrl), legacy);
    localStorage.removeItem(LEGACY_KEY);
    localStorage.setItem(MIGRATED_KEY, '1');
    console.log('[tadori] chat sessions: legacy 履歴を現在サイトへ移管しました');
  } catch { /* noop */ }
}

function lsGet(siteUrl: string): ChatSession[] {
  try {
    migrateLegacy(siteUrl);
    const raw = localStorage.getItem(keyFor(siteUrl));
    if (!raw) return [];
    const arr = JSON.parse(raw) as ChatSession[];
    return Array.isArray(arr) ? arr : [];
  } catch { return []; }
}

function lsSet(siteUrl: string, list: ChatSession[]): void {
  try { localStorage.setItem(keyFor(siteUrl), JSON.stringify(list)); } catch { /* quota */ }
}

/** 更新日時の新しい順。 */
export function listSessions(siteUrl: string): ChatSession[] {
  return lsGet(siteUrl).sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : -1));
}

export function getSession(siteUrl: string, id: string): ChatSession | undefined {
  return lsGet(siteUrl).find(s => s.id === id);
}

export function newSessionId(): string {
  return 's-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6);
}

function titleFrom(q: string): string {
  const t = q.replace(/\s+/g, ' ').trim();
  return t.length > 40 ? t.slice(0, 40) + '…' : (t || '新しいセッション');
}

/** turn を追記 (セッションが無ければ作成)。タイトルは最初の質問から。返り値は保存後のセッション。 */
export function appendTurn(siteUrl: string, id: string, turn: SavedTurn): ChatSession {
  const list = lsGet(siteUrl);
  let s = list.find(x => x.id === id);
  if (!s) {
    s = { id, title: titleFrom(turn.q), updatedAt: new Date().toISOString(), turns: [] };
    list.push(s);
  }
  if (s.turns.length === 0) s.title = titleFrom(turn.q);
  s.turns.push(turn);
  s.updatedAt = new Date().toISOString();
  list.sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : -1));
  lsSet(siteUrl, list.slice(0, MAX_SESSIONS));
  return s;
}

export function setTitle(siteUrl: string, id: string, title: string): void {
  const t = title.trim();
  if (!t) return;
  const list = lsGet(siteUrl);
  const s = list.find(x => x.id === id);
  if (!s) return;
  s.title = t.slice(0, 40);
  lsSet(siteUrl, list);
}

export function deleteSession(siteUrl: string, id: string): void {
  lsSet(siteUrl, lsGet(siteUrl).filter(s => s.id !== id));
}
