// 書き込み担当の単一化 (リース選出)。複数人が同時に SharePoint のセグメントを
// 書き換えると壊れるため、調整用 List「Tadori Sync」の単一リース行を ETag 楽観
// ロックで奪い合い、勝った1人だけが書き込みできるようにする (ADR-012)。
//
// - 書き込み前に ensureWriter() でリースを取得/更新。取れなければ書き込み拒否。
// - リースは LEASE_MS で失効。担当が落ちたら別の人が次に奪取 (タイムアウト移譲)。
// - start()/stop() で定期ハートビート + リース更新も可能 (在席表示・自動移譲用)。

import { SharePointClient, type FieldSpec, type SpItem } from '../sharepoint/client';

const SYNC_LIST = 'Tadori Sync';
const LEASE_KEY = '__lease__';
const HEARTBEAT_MS = 30_000;
const HEARTBEAT_IDLE_MS = 5 * 60_000; // 非アクティブタブ時はこちらの長い間隔へ切替
const LEASE_MS = 2 * 60_000; // リース有効期間。短めにして担当の輪番を促す
const PEER_ACTIVE_MS = 90_000; // 在席判定: last_seen がこれ以内なら「在席中」

const SYNC_FIELDS: FieldSpec[] = [
  { name: 'last_seen', type: 'datetime' },
  { name: 'holder', type: 'text' },
  { name: 'expires', type: 'datetime' },
];

function clientId(): string {
  try {
    let id = localStorage.getItem('tadori:client-id');
    if (!id) { id = 'c-' + Math.random().toString(36).slice(2, 10); localStorage.setItem('tadori:client-id', id); }
    return id;
  } catch { return 'c-anon'; }
}

export interface LeaseStatus {
  /** 自分の client-id */
  myId: string;
  /** 自分が現在の writer か */
  isWriter: boolean;
  /** 現在の writer の client-id (空文字なら未確定) */
  holderId: string;
  /** 現在のリース失効時刻 (Unix ms)。0 ならリース行未取得 */
  expiresAt: number;
  /** 「在席中」(last_seen が PEER_ACTIVE_MS 以内) のメンバ一覧 */
  peers: Array<{ id: string; lastSeen: number }>;
}

type StatusListener = (status: LeaseStatus) => void;

export class WriterLease {
  private readonly sp: SharePointClient;
  private readonly me = clientId();
  private listReady = false;
  private writer = false;
  private timer: number | null = null;
  private started = false;
  private status: LeaseStatus = { myId: this.me, isWriter: false, holderId: '', expiresAt: 0, peers: [] };
  private listeners = new Set<StatusListener>();
  private visibilityBound = false;

  constructor(siteUrl: string) { this.sp = new SharePointClient(siteUrl); }

  get id(): string { return this.me; }
  isWriter(): boolean { return this.writer; }
  getStatus(): LeaseStatus { return { ...this.status, peers: [...this.status.peers] }; }

  /** 状態 (writer / 在席者) が変わったら呼ばれる listener。返り値で unsubscribe。 */
  subscribe(cb: StatusListener): () => void {
    this.listeners.add(cb);
    // 初回呼び出しで現状を渡す
    try { cb(this.getStatus()); } catch { /* ignore */ }
    return () => this.listeners.delete(cb);
  }

  /** 書き込み直前に呼ぶ。リースを取得/更新し、書き込み可なら true。 */
  async ensureWriter(): Promise<boolean> {
    await this.ensureList();
    await this.electOrRenew();
    return this.writer;
  }

  /** 在席ハートビート + リース更新を定期実行。
   *  Page Visibility と連動して非アクティブ時はインターバルを延長 (5 分)。
   *  二重起動防止: 既に started なら何もしない。 */
  async start(): Promise<void> {
    if (this.started) return;
    this.started = true;
    await this.ensureList();
    await this.tick();
    this.scheduleNext();
    if (!this.visibilityBound && typeof document !== 'undefined') {
      this.visibilityBound = true;
      document.addEventListener('visibilitychange', () => {
        // visible に戻った瞬間に 1 回 tick して間隔も短くリセット
        if (!document.hidden) void this.tick();
        this.scheduleNext();
      });
    }
  }

  private scheduleNext(): void {
    if (this.timer != null) { window.clearInterval(this.timer); this.timer = null; }
    if (!this.started) return;
    const ms = (typeof document !== 'undefined' && document.hidden) ? HEARTBEAT_IDLE_MS : HEARTBEAT_MS;
    this.timer = window.setInterval(() => { void this.tick(); }, ms);
  }

  stop(): void {
    this.started = false;
    if (this.timer != null) { window.clearInterval(this.timer); this.timer = null; }
    void this.release();
  }

  private async ensureList(): Promise<void> {
    if (this.listReady) return;
    // バージョン履歴を無効化: 同じ行を 30 秒毎に上書きするので、履歴は無限に増えるが意味が無い。
    await this.sp.ensureList(SYNC_LIST, SYNC_FIELDS, { disableVersioning: true });
    this.listReady = true;
  }

  private async tick(): Promise<void> {
    try { await this.heartbeat(); await this.electOrRenew(); await this.refreshPeers(); }
    catch (e) { console.warn('[tadori/lease] tick 失敗:', (e as Error).message); }
  }

  /** Tadori Sync List から在席者 (last_seen が直近 PEER_ACTIVE_MS 以内) を集めて status.peers を更新。
   *  __lease__ 行や古い退職者行は除外する。 */
  private async refreshPeers(): Promise<void> {
    try {
      const rows = await this.sp.getItems(SYNC_LIST, `$select=Title,last_seen&$top=200`);
      const now = Date.now();
      const peers: Array<{ id: string; lastSeen: number }> = [];
      for (const r of rows) {
        const title = String(r.Title ?? '');
        if (!title || title === LEASE_KEY) continue;
        const last = Date.parse(String(r.last_seen ?? '')) || 0;
        if (last && now - last <= PEER_ACTIVE_MS) peers.push({ id: title, lastSeen: last });
      }
      peers.sort((a, b) => b.lastSeen - a.lastSeen);
      this.updateStatus({ peers });
    } catch (e) { /* peers 取得失敗は致命的ではない */ console.warn('[tadori/lease] peers 取得失敗:', (e as Error).message); }
  }

  /** 部分更新で status を書き換え、変化があれば listener へ通知。 */
  private updateStatus(patch: Partial<LeaseStatus>): void {
    const next: LeaseStatus = { ...this.status, ...patch };
    // 浅い等価比較で「実質変化なし」をスキップ
    const same =
      next.isWriter === this.status.isWriter &&
      next.holderId === this.status.holderId &&
      next.expiresAt === this.status.expiresAt &&
      next.peers.length === this.status.peers.length &&
      next.peers.every((p, i) => p.id === this.status.peers[i]?.id && p.lastSeen === this.status.peers[i]?.lastSeen);
    if (same) return;
    this.status = next;
    for (const cb of this.listeners) { try { cb(this.getStatus()); } catch { /* ignore */ } }
  }

  private async findRow(title: string): Promise<SpItem | null> {
    const rows = await this.sp.getItems(SYNC_LIST, `$select=Id&$filter=Title eq '${title}'&$top=1`);
    if (rows.length === 0) return null;
    return this.sp.getItem(SYNC_LIST, Number(rows[0].Id)); // ETag 付きで取り直す
  }

  private async heartbeat(): Promise<void> {
    const now = new Date().toISOString();
    const row = await this.findRow(this.me);
    if (row) await this.sp.updateItem(SYNC_LIST, row.Id, { last_seen: now }, '*');
    else await this.sp.createItem(SYNC_LIST, { Title: this.me, last_seen: now });
  }

  private async electOrRenew(): Promise<void> {
    const nowMs = Date.now();
    const untilMs = nowMs + LEASE_MS;
    const until = () => new Date(untilMs).toISOString();
    const lease = await this.findRow(LEASE_KEY);

    if (!lease) {
      try {
        await this.sp.createItem(SYNC_LIST, { Title: LEASE_KEY, holder: this.me, expires: until() });
        this.writer = true;
        this.updateStatus({ isWriter: true, holderId: this.me, expiresAt: untilMs });
      } catch {
        this.writer = false; // 同時 create 競合 → 次 tick で再判定
        this.updateStatus({ isWriter: false });
      }
      return;
    }

    const holder = String(lease.holder ?? '');
    const expires = Date.parse(String(lease.expires ?? '')) || 0;

    if (holder === this.me || expires < nowMs) {
      // 自分の更新、または失効分の奪取。ETag 競合(412)なら奪えなかった = reader。
      const ok = await this.sp.updateItem(SYNC_LIST, lease.Id, { holder: this.me, expires: until() }, lease.__etag);
      this.writer = ok;
      this.updateStatus({ isWriter: ok, holderId: ok ? this.me : holder, expiresAt: ok ? untilMs : expires });
    } else {
      this.writer = false; // 他者が有効なリース保持中
      this.updateStatus({ isWriter: false, holderId: holder, expiresAt: expires });
    }
  }

  private async release(): Promise<void> {
    if (!this.listReady || !this.writer) return;
    try {
      const lease = await this.findRow(LEASE_KEY);
      if (lease && String(lease.holder) === this.me) {
        await this.sp.updateItem(SYNC_LIST, lease.Id, { expires: new Date().toISOString() }, lease.__etag);
      }
    } catch { /* best-effort */ }
    this.writer = false;
  }
}

// ─── siteUrl ごとの共有インスタンス ──────────────────────────────────────────
let shared: WriterLease | null = null;
let sharedSite = '';

export function getLease(siteUrl: string): WriterLease {
  if (!shared || sharedSite !== siteUrl) { shared = new WriterLease(siteUrl); sharedSite = siteUrl; }
  return shared;
}
