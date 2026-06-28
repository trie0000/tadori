// docx / xlsx をブラウザだけでテキスト抽出する (Word/Excel 不要、外部依存なし)。
//
// OOXML (docx/xlsx/pptx) は実体が ZIP。ブラウザのネイティブ DecompressionStream
// ('deflate-raw') で各エントリを展開できるので、必要な XML パートだけ取り出して
// テキスト化する。Word COM のように「アプリで開く」必要は一切ない。
//
// 対応:
//   .docx → word/document.xml から段落テキスト
//   .xlsx → xl/sharedStrings.xml + 各 sheet の inlineStr からセルテキスト

interface ZipEntry { name: string; data: Uint8Array; }

/** ZIP バイト列から必要なエントリだけ展開して返す。
 *  Central Directory を読んでエントリを特定し、stored(0)/deflate(8) に対応。 */
async function unzip(buf: ArrayBuffer, want: (name: string) => boolean): Promise<ZipEntry[]> {
  const u8 = new Uint8Array(buf);
  const dv = new DataView(buf);
  // End of Central Directory (EOCD) を末尾から探す (シグネチャ 0x06054b50)
  let eocd = -1;
  for (let i = u8.length - 22; i >= 0 && i >= u8.length - 22 - 65536; i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('ZIP の EOCD が見つかりません (壊れたファイル?)');
  const cdCount = dv.getUint16(eocd + 10, true);
  let cdOff = dv.getUint32(eocd + 16, true);

  const out: ZipEntry[] = [];
  for (let n = 0; n < cdCount; n++) {
    if (dv.getUint32(cdOff, true) !== 0x02014b50) break; // Central Directory header
    const method = dv.getUint16(cdOff + 10, true);
    const compSize = dv.getUint32(cdOff + 20, true);
    const nameLen = dv.getUint16(cdOff + 28, true);
    const extraLen = dv.getUint16(cdOff + 30, true);
    const commentLen = dv.getUint16(cdOff + 32, true);
    const localOff = dv.getUint32(cdOff + 42, true);
    const name = new TextDecoder('utf-8').decode(u8.subarray(cdOff + 46, cdOff + 46 + nameLen));
    cdOff += 46 + nameLen + extraLen + commentLen;

    if (!want(name)) continue;

    // ローカルヘッダから実データ開始位置を求める
    if (dv.getUint32(localOff, true) !== 0x04034b50) continue;
    const lNameLen = dv.getUint16(localOff + 26, true);
    const lExtraLen = dv.getUint16(localOff + 28, true);
    const dataStart = localOff + 30 + lNameLen + lExtraLen;
    const comp = u8.subarray(dataStart, dataStart + compSize);

    let data: Uint8Array;
    if (method === 0) {
      data = comp; // stored
    } else if (method === 8) {
      data = await inflateRaw(comp);
    } else {
      continue; // 未対応の圧縮方式はスキップ
    }
    out.push({ name, data });
  }
  return out;
}

/** raw deflate を展開 (ブラウザネイティブ)。 */
async function inflateRaw(comp: Uint8Array): Promise<Uint8Array> {
  // 念のため subarray のオフセットを 0 始まりの独立バッファにコピー
  const src = comp.byteOffset === 0 && comp.byteLength === comp.buffer.byteLength
    ? comp : comp.slice();
  const ds = new DecompressionStream('deflate-raw');
  const stream = new Blob([src as BlobPart]).stream().pipeThrough(ds);
  const ab = await new Response(stream).arrayBuffer();
  return new Uint8Array(ab);
}

function xmlToText(xml: string): string {
  return xml
    .replace(/<\/w:p>/g, '\n')        // 段落終わり → 改行
    .replace(/<w:tab\/?>/g, '\t')      // タブ
    .replace(/<w:br\/?>/g, '\n')       // 改行
    .replace(/<[^>]+>/g, '')           // 残りのタグを除去
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&apos;/g, "'")
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/** .docx → テキスト。 */
export async function docxToText(buf: ArrayBuffer): Promise<string> {
  const entries = await unzip(buf, n => n === 'word/document.xml');
  const doc = entries.find(e => e.name === 'word/document.xml');
  if (!doc) throw new Error('word/document.xml が見つかりません (docx ではない?)');
  return xmlToText(new TextDecoder('utf-8').decode(doc.data));
}

/** .xlsx → テキスト (sharedStrings + 各シートのセル値を改行区切りで)。 */
export async function xlsxToText(buf: ArrayBuffer): Promise<string> {
  const entries = await unzip(buf, n => n === 'xl/sharedStrings.xml' || /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
  const dec = new TextDecoder('utf-8');

  // sharedStrings: <si>..<t>text</t>..</si> の配列
  const shared: string[] = [];
  const ss = entries.find(e => e.name === 'xl/sharedStrings.xml');
  if (ss) {
    const xml = dec.decode(ss.data);
    for (const m of xml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
      const texts = [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(t => t[1]);
      shared.push(decodeEntities(texts.join('')));
    }
  }

  const lines: string[] = [];
  const sheets = entries.filter(e => /sheet\d+\.xml$/.test(e.name)).sort((a, b) => a.name.localeCompare(b.name));
  for (const sh of sheets) {
    const xml = dec.decode(sh.data);
    for (const c of xml.matchAll(/<c[^>]*?(?:\st="([^"]*)")?[^>]*>([\s\S]*?)<\/c>/g)) {
      const type = c[1];
      const inner = c[2];
      const vm = inner.match(/<v>([\s\S]*?)<\/v>/);
      const isM = inner.match(/<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/is>/);
      let val = '';
      if (type === 's' && vm) { val = shared[Number(vm[1])] ?? ''; }   // shared string index
      else if (isM) { val = decodeEntities(isM[1]); }                   // inline string
      else if (vm) { val = vm[1]; }                                     // number / direct
      if (val) lines.push(val);
    }
  }
  return lines.join('\n').trim();
}

function decodeEntities(s: string): string {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&apos;/g, "'");
}
