declare const __TADORI_BUILD_ID__: string;
declare const __TADORI_VERSION__: string;

interface Window {
  _spPageContextInfo?: { webAbsoluteUrl?: string };
}

declare module '*.css' {
  const content: string;
  export default content;
}

// pdf.js は .mjs 直接 import で型が無いので最小宣言 (利用側で構造を絞る)。
declare module 'pdfjs-dist/build/pdf.mjs' {
  export const GlobalWorkerOptions: { workerSrc: string };
  export function getDocument(src: unknown): { promise: Promise<unknown>; destroy: () => void };
}

// build.js が生成する pdf.js worker ソース文字列。
declare module './pdfWorker.generated' {
  const workerSource: string;
  export default workerSource;
}
