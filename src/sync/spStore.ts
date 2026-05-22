// ベクトルDB セグメント群の SharePoint 配置 (ファイルサーバ不要・SharePoint だけ)。
// /<site>/<library>/Tadori/ 配下に manifest.json と seg-*.json を置く。
// 読み書きはブラウザの Cookie 認証 REST (relay は SPO に認証できないため)。

import { SharePointClient } from '../sharepoint/client';
import {
  type Manifest, type Segment,
  emptyManifest, serializeManifest, parseManifest, serializeSegment, parseSegment,
} from './segments';

const MANIFEST_NAME = 'manifest.json';
const DEFAULT_LIBRARY = 'Shared Documents'; // SPO 既定ライブラリの server-relative 名

export class SpVectorStore {
  private readonly folder: string; // site 込み server-relative (例: /sites/n365/Shared Documents/Tadori)

  constructor(private readonly sp: SharePointClient, library = DEFAULT_LIBRARY, sub = 'Tadori') {
    this.folder = `${sp.serverRelativeSite()}/${library}/${sub}`;
  }

  /** 配置フォルダを作成 (冪等)。 */
  async ensure(): Promise<void> {
    await this.sp.ensureFolder(this.folder);
  }

  async readManifest(): Promise<Manifest | null> {
    const text = await this.sp.readFileText(`${this.folder}/${MANIFEST_NAME}`);
    if (text == null) return null;
    return parseManifest(text);
  }

  async writeManifest(m: Manifest): Promise<void> {
    await this.sp.uploadFileText(this.folder, MANIFEST_NAME, serializeManifest(m));
  }

  async readSegment(id: string): Promise<Segment | null> {
    const text = await this.sp.readFileText(`${this.folder}/${id}.json`);
    if (text == null) return null;
    return parseSegment(text);
  }

  async writeSegment(seg: Segment): Promise<void> {
    await this.sp.uploadFileText(this.folder, `${seg.id}.json`, serializeSegment(seg));
  }

  /** 既存セグメント id 一覧 (manifest 不整合の復旧用)。 */
  async listSegmentIds(): Promise<string[]> {
    const names = await this.sp.listFolderFileNames(this.folder);
    return names.filter(n => n.startsWith('seg-') && n.endsWith('.json')).map(n => n.slice(0, -5));
  }

  /** SharePoint 上の Tadori フォルダ配下の manifest と全セグメントを削除。
   *  「取り込みメールの全削除」用。フォルダ自体は残す (次回取り込みで再利用)。 */
  async deleteAll(): Promise<void> {
    const names = await this.sp.listFolderFileNames(this.folder);
    for (const name of names) {
      if (name === MANIFEST_NAME || (name.startsWith('seg-') && name.endsWith('.json'))) {
        await this.sp.deleteFile(`${this.folder}/${name}`).catch(() => { /* best-effort */ });
      }
    }
  }

  /** 初回・manifest 不在時は空 manifest を作って配置する。 */
  async ensureManifest(): Promise<Manifest> {
    const m = await this.readManifest();
    if (m) return m;
    const fresh = emptyManifest();
    await this.ensure();
    await this.writeManifest(fresh);
    return fresh;
  }
}
