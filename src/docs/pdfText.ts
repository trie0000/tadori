// PDF テキスト抽出。
//
// ※ pdf.js をバンドルに同梱すると 2MB 超に膨らみ bookmarklet が壊れたため、
//    一旦 PDF 対応を停止 (docx/xlsx/md/txt はブラウザでパース可能)。
//    PDF は別途、軽量に遅延ロードする方式で再導入予定。

export async function pdfToText(_buf: ArrayBuffer, _signal?: AbortSignal): Promise<string> {
  throw new Error('PDF は現在未対応です (docx/xlsx/md/txt は取り込めます)。');
}
