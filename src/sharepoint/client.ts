// SharePoint List REST クライアント。PoC 01 で Cookie 認証 READ を確認済み。
// 全リクエストは credentials:'include' でブラウザの既存セッション Cookie を借用。
// 書き込みは FormDigest + ETag (If-Match) による楽観ロック (ADR-005)。

export interface SpItem {
  Id: number;
  __etag: string;
  [field: string]: unknown;
}

export type FieldType = 'text' | 'note' | 'number' | 'datetime' | 'boolean';
export interface FieldSpec {
  /** 列の表示名 = 内部名 (ASCII + アンダースコアなら内部名は表示名と一致)。 */
  name: string;
  type: FieldType;
}

/** SP REST `/fields` POST 用の型付きペイロード (odata=verbose)。 */
function toFieldSchema(f: FieldSpec): Record<string, unknown> {
  switch (f.type) {
    case 'text':     return { __metadata: { type: 'SP.FieldText' }, FieldTypeKind: 2, Title: f.name };
    case 'note':     return { __metadata: { type: 'SP.FieldMultiLineText' }, FieldTypeKind: 3, Title: f.name, RichText: false, NumberOfLines: 6 };
    case 'number':   return { __metadata: { type: 'SP.FieldNumber' }, FieldTypeKind: 9, Title: f.name };
    case 'datetime': return { __metadata: { type: 'SP.FieldDateTime' }, FieldTypeKind: 4, Title: f.name, DisplayFormat: 1 };
    case 'boolean':  return { __metadata: { type: 'SP.Field' }, FieldTypeKind: 8, Title: f.name };
  }
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

  /** 新規アイテムを作成。fields は列の内部名 → 値。作成された Id を返す。 */
  async createItem(listTitle: string, fields: Record<string, unknown>): Promise<number> {
    const digest = await this.getFormDigest();
    const res = await fetch(`${this.listApi(listTitle)}/items`, {
      method: 'POST',
      headers: await this.headers({
        'Content-Type': 'application/json;odata=nometadata',
        'X-RequestDigest': digest,
      }),
      credentials: 'include',
      body: JSON.stringify(fields),
    });
    if (!res.ok) {
      const body = await res.text().catch(() => '');
      throw new Error(`createItem HTTP ${res.status} ${body.slice(0, 300)}`);
    }
    const json = await res.json() as { Id?: number };
    return json.Id ?? 0;
  }

  /** リストが無ければ作成し、不足列を追加する (冪等)。新規作成したら true を返す。
   *  opts.disableVersioning が true なら List のバージョン履歴を無効化する
   *  (ハートビート用 List 等、頻繁に同じ行を更新する用途でバージョン履歴が
   *  無限に膨らむのを防ぐ)。 */
  async ensureList(listTitle: string, fields: FieldSpec[], opts: { disableVersioning?: boolean } = {}): Promise<boolean> {
    const existed = await this.listExists(listTitle);
    if (!existed) await this.createList(listTitle);
    // 列追加は best-effort (権限不足等で失敗しても致命にしない)。
    try { await this.ensureFields(listTitle, fields); }
    catch (e) { console.warn('[tadori] ensureFields 失敗:', (e as Error).message); }
    if (opts.disableVersioning) {
      try { await this.setListVersioning(listTitle, false); }
      catch (e) { console.warn('[tadori] バージョン無効化失敗:', (e as Error).message); }
    }
    return !existed;
  }

  /** List のバージョン履歴を on/off。新規/既存どちらの List にも適用可能。 */
  private async setListVersioning(listTitle: string, enabled: boolean): Promise<void> {
    const digest = await this.getFormDigest();
    const res = await fetch(this.listApi(listTitle), {
      method: 'POST',
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
        'X-HTTP-Method': 'MERGE',
        'If-Match': '*',
      },
      credentials: 'include',
      body: JSON.stringify({
        __metadata: { type: 'SP.List' },
        EnableVersioning: enabled,
        EnableMinorVersions: enabled,
      }),
    });
    if (!res.ok) {
      const body = await res.text().catch(() => '');
      throw new Error(`setListVersioning HTTP ${res.status} ${body.slice(0, 200)}`);
    }
  }

  private async listExists(listTitle: string): Promise<boolean> {
    const res = await fetch(`${this.listApi(listTitle)}?$select=Id`, {
      headers: await this.headers(),
      credentials: 'include',
    });
    if (res.status === 404) return false;
    return res.ok;
  }

  private async createList(listTitle: string): Promise<void> {
    const digest = await this.getFormDigest();
    const res = await fetch(`${this.siteUrl}/_api/web/lists`, {
      method: 'POST',
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
      },
      credentials: 'include',
      body: JSON.stringify({
        __metadata: { type: 'SP.List' },
        Title: listTitle,
        BaseTemplate: 100, // 汎用リスト
        AllowContentTypes: true,
        ContentTypesEnabled: false,
      }),
    });
    if (!res.ok) {
      const body = await res.text().catch(() => '');
      throw new Error(`createList HTTP ${res.status} ${body.slice(0, 300)}`);
    }
  }

  private async ensureFields(listTitle: string, fields: FieldSpec[]): Promise<void> {
    const existing = await this.listFieldNames(listTitle);
    const digest = await this.getFormDigest();
    for (const f of fields) {
      if (existing.has(f.name)) continue;
      const res = await fetch(`${this.listApi(listTitle)}/fields`, {
        method: 'POST',
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'X-RequestDigest': digest,
        },
        credentials: 'include',
        body: JSON.stringify(toFieldSchema(f)),
      });
      if (!res.ok) {
        const body = await res.text().catch(() => '');
        throw new Error(`addField(${f.name}) HTTP ${res.status} ${body.slice(0, 200)}`);
      }
    }
  }

  private async listFieldNames(listTitle: string): Promise<Set<string>> {
    const res = await fetch(`${this.listApi(listTitle)}/fields?$select=InternalName,Title,StaticName&$top=500`, {
      headers: await this.headers(),
      credentials: 'include',
    });
    const set = new Set<string>();
    if (!res.ok) return set;
    const json = await res.json() as { value?: { InternalName?: string; Title?: string; StaticName?: string }[] };
    for (const f of json.value ?? []) {
      if (f.InternalName) set.add(f.InternalName);
      if (f.StaticName) set.add(f.StaticName);
      if (f.Title) set.add(f.Title);
    }
    return set;
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

  // ─── ドキュメントライブラリ ファイル操作 (ベクトルDB セグメント配布用) ───────

  /** サイトの server-relative パス (例: /sites/n365)。 */
  serverRelativeSite(): string {
    try { return new URL(this.siteUrl).pathname.replace(/\/+$/, ''); } catch { return ''; }
  }

  /** フォルダが無ければ作成。serverRelativeUrl は site 込み (例 /sites/X/Shared Documents/Tadori)。
   *  存在確認 → 無ければ classic な verbose POST で作成 (addUsingPath は環境により 400)。 */
  async ensureFolder(serverRelativeUrl: string): Promise<void> {
    // 既に存在するか確認 (存在すれば 200 + Exists:true)
    try {
      const check = await fetch(
        `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')?$select=Exists`,
        { headers: await this.headers(), credentials: 'include' },
      );
      if (check.ok) {
        const j = await check.json() as { Exists?: boolean };
        if (j.Exists) return;
      }
    } catch { /* 確認に失敗しても作成を試す */ }

    const digest = await this.getFormDigest();
    const res = await fetch(`${this.siteUrl}/_api/web/folders`, {
      method: 'POST',
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
      },
      credentials: 'include',
      body: JSON.stringify({ __metadata: { type: 'SP.Folder' }, ServerRelativeUrl: serverRelativeUrl }),
    });
    if (res.ok) return;
    // 既に存在する系は成功扱い (冪等)
    const body = await res.text().catch(() => '');
    if (res.status === 409 || /exist|既に|already/i.test(body)) return;
    throw new Error(`ensureFolder HTTP ${res.status} ${body.slice(0, 200)}`);
  }

  /** ファイル本文をテキストで取得。存在しなければ null。 */
  async readFileText(serverRelativeUrl: string): Promise<string | null> {
    const res = await fetch(
      `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')/$value`,
      { headers: { Accept: '*/*' }, credentials: 'include' },
    );
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`readFile HTTP ${res.status} (${serverRelativeUrl})`);
    return res.text();
  }

  /** ファイル本文 + ETag を取得 (CAS 用)。存在しなければ null。 */
  async readFileTextWithEtag(serverRelativeUrl: string): Promise<{ text: string; etag: string } | null> {
    const res = await fetch(
      `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')/$value`,
      { headers: { Accept: '*/*' }, credentials: 'include' },
    );
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`readFile HTTP ${res.status} (${serverRelativeUrl})`);
    const text = await res.text();
    const etag = res.headers.get('ETag') || res.headers.get('etag') || '';
    return { text, etag };
  }

  /** 既存ファイルを If-Match (ETag) で楽観ロック更新。
   *  ETag 不一致時は 412 → `CasConflictError` を投げる。呼び出し側はキャッチして再読込→リトライする。 */
  async uploadFileTextCas(serverRelativeUrl: string, text: string, ifMatchEtag: string): Promise<void> {
    const digest = await this.getFormDigest();
    const res = await fetch(
      `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')/$value`,
      {
        method: 'POST',
        headers: await this.headers({
          'Content-Type': 'text/plain;charset=utf-8',
          'X-RequestDigest': digest,
          'X-HTTP-Method': 'PUT',
          'If-Match': ifMatchEtag,
        }),
        credentials: 'include',
        body: text,
      },
    );
    if (res.status === 412) {
      const err = new Error('CAS conflict (412)') as Error & { code: string };
      err.code = 'PRECONDITION_FAILED';
      throw err;
    }
    if (!res.ok) {
      const body = await res.text().catch(() => '');
      throw new Error(`uploadFileTextCas HTTP ${res.status} ${body.slice(0, 200)}`);
    }
  }

  /** 同名ファイルがある時は失敗するアップロード (segment 名衝突防止用)。
   *  作成成功なら true、既に存在 (409 や SP 固有の 400/500 + "already exists") なら false。 */
  async uploadFileTextNoOverwrite(folderServerRelativeUrl: string, name: string, text: string): Promise<boolean> {
    const digest = await this.getFormDigest();
    const url = `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')`
      + `/Files/add(url='${encodeURIComponent(name)}',overwrite=false)`;
    const res = await fetch(url, {
      method: 'POST',
      headers: await this.headers({ 'Content-Type': 'text/plain;charset=utf-8', 'X-RequestDigest': digest }),
      credentials: 'include',
      body: text,
    });
    if (res.ok) return true;
    // SharePoint は overwrite=false で既存ファイルに当たると 400/500 を返し、本文に「既に存在」系メッセージが入る。
    if (res.status === 409 || res.status === 400 || res.status === 500) {
      const body = await res.text().catch(() => '');
      if (/already exists|exists at|存在|already there/i.test(body)) return false;
    }
    const body = await res.text().catch(() => '');
    throw new Error(`uploadFileTextNoOverwrite HTTP ${res.status} ${body.slice(0, 200)}`);
  }

  /** テキストをファイルとしてアップロード (上書き)。folder は site 込み server-relative。 */
  async uploadFileText(folderServerRelativeUrl: string, name: string, text: string): Promise<void> {
    const digest = await this.getFormDigest();
    const url = `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')`
      + `/Files/add(url='${encodeURIComponent(name)}',overwrite=true)`;
    const res = await fetch(url, {
      method: 'POST',
      headers: await this.headers({ 'Content-Type': 'text/plain;charset=utf-8', 'X-RequestDigest': digest }),
      credentials: 'include',
      body: text,
    });
    if (!res.ok) {
      const body = await res.text().catch(() => '');
      throw new Error(`uploadFile(${name}) HTTP ${res.status} ${body.slice(0, 200)}`);
    }
  }

  /** ファイルを削除 (リセット用)。失敗は throw。 */
  async deleteFile(serverRelativeUrl: string): Promise<void> {
    const digest = await this.getFormDigest();
    const res = await fetch(
      `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')`,
      {
        method: 'POST',
        headers: await this.headers({ 'X-RequestDigest': digest, 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' }),
        credentials: 'include',
      },
    );
    if (!res.ok && res.status !== 404) {
      const b = await res.text().catch(() => '');
      throw new Error(`deleteFile HTTP ${res.status} ${b.slice(0, 200)}`);
    }
  }

  /** フォルダ直下のファイル名一覧。フォルダが無ければ空配列。 */
  async listFolderFileNames(folderServerRelativeUrl: string): Promise<string[]> {
    const res = await fetch(
      `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')/Files?$select=Name&$top=5000`,
      { headers: await this.headers(), credentials: 'include' },
    );
    if (!res.ok) return [];
    const json = await res.json() as { value?: { Name?: string }[] };
    return (json.value ?? []).map(f => f.Name ?? '').filter(Boolean);
  }
}
