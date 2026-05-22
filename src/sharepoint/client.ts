// SharePoint List REST クライアント。PoC 01 で Cookie 認証 READ を確認済み。
// 全リクエストは credentials:'include' でブラウザの既存セッション Cookie を借用。
// 書き込みは FormDigest + ETag (If-Match) による楽観ロック (ADR-005)。

export interface SpItem {
  Id: number;
  __etag: string;
  [field: string]: unknown;
}

export class SharePointClient {
  private digest: string | null = null;
  private digestAt = 0;

  /** @param siteUrl 末尾スラッシュなしのサイト絶対 URL。 */
  constructor(private readonly siteUrl: string) {}

  private listApi(listTitle: string): string {
    return `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')`;
  }

  private async headers(extra: Record<string, string> = {}): Promise<Record<string, string>> {
    return {
      Accept: 'application/json;odata=nometadata',
      ...extra,
    };
  }

  /** FormDigest を取得・キャッシュ (有効期限は余裕を見て 20 分)。書き込み前に必須。 */
  async getFormDigest(): Promise<string> {
    const now = Date.now();
    if (this.digest && now - this.digestAt < 20 * 60_000) return this.digest;
    const res = await fetch(`${this.siteUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: await this.headers(),
      credentials: 'include',
    });
    if (!res.ok) throw new Error(`contextinfo HTTP ${res.status}`);
    const json = await res.json() as { FormDigestValue?: string };
    if (!json.FormDigestValue) throw new Error('FormDigestValue missing');
    this.digest = json.FormDigestValue;
    this.digestAt = now;
    return this.digest;
  }

  /** OData クエリでアイテムを取得。$filter / $select / $top などを渡す。 */
  async getItems(listTitle: string, query: string): Promise<SpItem[]> {
    const url = `${this.listApi(listTitle)}/items?${query}`;
    const res = await fetch(url, { headers: await this.headers(), credentials: 'include' });
    if (!res.ok) {
      const body = await res.text().catch(() => '');
      throw new Error(`getItems HTTP ${res.status} ${body.slice(0, 300)}`);
    }
    const json = await res.json() as { value: SpItem[] };
    return json.value ?? [];
  }

  /** 単一アイテムを ETag 付きで取得 (try-claim の前段)。 */
  async getItem(listTitle: string, id: number, select?: string): Promise<SpItem> {
    const sel = select ? `?$select=${select}` : '';
    const res = await fetch(`${this.listApi(listTitle)}/items(${id})${sel}`, {
      headers: await this.headers(),
      credentials: 'include',
    });
    if (!res.ok) throw new Error(`getItem(${id}) HTTP ${res.status}`);
    const etag = res.headers.get('ETag') ?? '*';
    const json = await res.json() as SpItem;
    return { ...json, __etag: etag };
  }

  /** ETag 楽観ロックで MERGE 更新。412 (競合) は false を返す。
   *  ETag を '*' にすると無条件更新 (claim 競合検証には使わないこと)。 */
  async updateItem(
    listTitle: string,
    id: number,
    fields: Record<string, unknown>,
    etag: string,
  ): Promise<boolean> {
    const digest = await this.getFormDigest();
    const res = await fetch(`${this.listApi(listTitle)}/items(${id})`, {
      method: 'POST',
      headers: await this.headers({
        'Content-Type': 'application/json;odata=nometadata',
        'X-RequestDigest': digest,
        'IF-MATCH': etag,
        'X-HTTP-Method': 'MERGE',
      }),
      credentials: 'include',
      body: JSON.stringify(fields),
    });
    if (res.status === 412) return false; // 楽観ロック競合 → 別クライアントが先取り
    if (!res.ok) {
      const body = await res.text().catch(() => '');
      throw new Error(`updateItem(${id}) HTTP ${res.status} ${body.slice(0, 300)}`);
    }
    return true;
  }
}
