// チャット履歴の永続化 (localStorage)。左ペインのセッション一覧 + 再表示用。
// 1 セッション = 1 スレッド。turns に質問/回答(Markdown)/出典メタを保存する。

export interface SavedHit {
  messageId: string;
  internetMessageId: string;
  subject: string;
  from: string;
  date: string;
  body: string;
  isHtml: boolean;
  score: number;
}

export interface SavedTurn {
  q: string;
  /** 回答の生 Markdown (再表示時に renderMarkdown する)。 */
  answer: string;
  hits: SavedHit[];
  ms: number;
}

export interface ChatSession {
  id: string;
  title: string;
  updatedAt: string;
  turns: SavedTurn[];
}

const KEY = 'tadori:chat-sessions';
const MAX_SESSIONS = 50;

function lsGet(): ChatSession[] {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return [];
    const arr = JSON.parse(raw) as ChatSession[];
    return Array.isArray(arr) ? arr : [];
  } catch { return []; }
}

function lsSet(list: ChatSession[]): void {
  try { localStorage.setItem(KEY, JSON.stringify(list)); } catch { /* quota */ }
}

/** 更新日時の新しい順。 */
export function listSessions(): ChatSession[] {
  return lsGet().sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : -1));
}

export function getSession(id: string): ChatSession | undefined {
  return lsGet().find(s => s.id === id);
}

export function newSessionId(): string {
  return 's-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6);
}

function titleFrom(q: string): string {
  const t = q.replace(/\s+/g, ' ').trim();
  return t.length > 40 ? t.slice(0, 40) + '…' : (t || '新しいセッション');
}

/** turn を追記 (セッションが無ければ作成)。タイトルは最初の質問から。返り値は保存後のセッション。 */
export function appendTurn(id: string, turn: SavedTurn): ChatSession {
  const list = lsGet();
  let s = list.find(x => x.id === id);
  if (!s) {
    s = { id, title: titleFrom(turn.q), updatedAt: new Date().toISOString(), turns: [] };
    list.push(s);
  }
  if (s.turns.length === 0) s.title = titleFrom(turn.q);
  s.turns.push(turn);
  s.updatedAt = new Date().toISOString();
  // 上限を超えたら古いものから捨てる
  list.sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : -1));
  lsSet(list.slice(0, MAX_SESSIONS));
  return s;
}

export function deleteSession(id: string): void {
  lsSet(lsGet().filter(s => s.id !== id));
}
