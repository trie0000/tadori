// PDF をブラウザだけでテキスト抽出する (Word 不要)。pdf.js を使用。
//
// worker はビルド時に生成される pdfWorker.generated.ts (pdf.worker.min.mjs の
// ソース文字列) から Blob URL を作って設定する。これで単一バンドルに閉じ、
// SP への追加ファイル配置や CDN を不要にする。
//
// 注: テキスト層のある PDF のみ抽出可能。スキャン画像のみの PDF は文字が取れない
//     (その場合は将来 Vision で対応)。

import * as pdfjsNs from 'pdfjs-dist/build/pdf.mjs';
import workerSource from './pdfWorker.generated';

const pdfjs = pdfjsNs as unknown as {
  GlobalWorkerOptions: { workerSrc: string };
  getDocument: (src: { data: Uint8Array; isEvalSupported: boolean }) => { promise: Promise<PdfDoc>; destroy: () => void };
};

let workerReady = false;
function ensureWorker(): void {
  if (workerReady) return;
  try {
    const blob = new Blob([workerSource], { type: 'text/javascript' });
    pdfjs.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
    workerReady = true;
  } catch (e) {
    console.warn('[tadori] pdf worker 設定失敗:', (e as Error).message);
  }
}

/** PDF (ArrayBuffer) → 全ページのテキストを改行区切りで連結。 */
export async function pdfToText(buf: ArrayBuffer, signal?: AbortSignal): Promise<string> {
  ensureWorker();
  // getDocument は構造化複製のため、Uint8Array を渡す (ArrayBuffer は detach されうる)
  const data = new Uint8Array(buf.slice(0));
  const loadingTask = pdfjs.getDocument({ data, isEvalSupported: false });
  const doc = await loadingTask.promise;
  try {
    const parts: string[] = [];
    for (let p = 1; p <= doc.numPages; p++) {
      if (signal?.aborted) throw new DOMException('aborted', 'AbortError');
      const page = await doc.getPage(p);
      const content = await page.getTextContent();
      // items を読み順 (Y 降順 → X 昇順) でつなぐ。簡易に items の str を空白連結 +
      // hasEOL / transform の Y 変化で改行を入れる。
      let line = '';
      let lastY: number | null = null;
      const pageLines: string[] = [];
      for (const it of content.items) {
        const item = it as PdfTextItem;
        if (typeof item.str !== 'string') continue;
        const y = item.transform ? item.transform[5] : null;
        if (lastY !== null && y !== null && Math.abs(y - lastY) > 2) {
          if (line.trim()) pageLines.push(line.trim());
          line = '';
        }
        line += item.str;
        if (item.hasEOL) { if (line.trim()) pageLines.push(line.trim()); line = ''; }
        lastY = y;
      }
      if (line.trim()) pageLines.push(line.trim());
      const pageText = pageLines.join('\n').trim();
      if (pageText) parts.push(pageText);
      // page リソース解放
      try { (page as unknown as { cleanup: () => void }).cleanup(); } catch { /* noop */ }
    }
    return parts.join('\n\n').trim();
  } finally {
    try { loadingTask.destroy(); } catch { /* noop */ }
  }
}

interface PdfDoc {
  numPages: number;
  getPage: (n: number) => Promise<PdfPage>;
}
interface PdfPage {
  getTextContent: () => Promise<{ items: unknown[] }>;
}
interface PdfTextItem {
  str: string;
  hasEOL?: boolean;
  transform?: number[];
}
