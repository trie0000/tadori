// PDF テキスト抽出。
//
// ローカル relay の /tadori/pdf-extract (PdfPig, Apache-2.0) でバイト列から
// テキストを抽出する。Word も pdf.js も使わない (バンドルは太らせない)。
// relay 起動が必要。

/** PDF (ArrayBuffer) → テキスト。relay の PdfPig で抽出。 */
export async function pdfToText(buf: ArrayBuffer, relayBaseUrl: string, fileName: string, signal?: AbortSignal): Promise<string> {
  if (!relayBaseUrl) throw new Error('中継サーバ URL が未設定です (PDF は relay で解析します。AI 接続で設定)');
  const url = `${relayBaseUrl.replace(/\/+$/, '')}/tadori/pdf-extract`;
  let res: Response;
  try {
    res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream', 'X-Tadori-Filename': encodeURIComponent(fileName) },
      body: buf, signal,
    });
  } catch (e) {
    throw new Error(`relay (${url}) への通信に失敗: ${(e as Error).message}。relay 起動を確認してください。`);
  }
  if (!res.ok) {
    const t = await res.text().catch(() => '');
    throw new Error(`relay /pdf-extract HTTP ${res.status} ${t.slice(0, 400)}`);
  }
  const json = await res.json() as { ok?: boolean; text?: string; error?: { code?: string; detail?: string } };
  if (!json.ok) throw new Error(`relay /pdf-extract 失敗 [${json.error?.code ?? '?'}]: ${json.error?.detail ?? ''}`);
  return json.text ?? '';
}
