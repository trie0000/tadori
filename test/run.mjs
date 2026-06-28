// ローカル検証ランナー。test/*.test.ts を esbuild で Node 向けにバンドルし、
// 本番コードの外部依存だけスタブに差し替えて `node --test` で実行する。
//
//  - ../db/engine        → test/_stubs/engine.ts  (SP/IndexedDB に触れない)
//  - ../embeddings/router → test/_stubs/router.ts (Azure/Voyage に触れない)
//  - mailhtml / mailtext  → DOMPurify 不要の no-op (Node に DOM が無いため)
//  - localStorage         → メモリ実装を polyfill
//
// 新しい依存は増やさない (esbuild + Node 標準 test のみ)。

import * as esbuild from 'esbuild';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import fs from 'node:fs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outDir = path.join(root, 'dist-test');
fs.rmSync(outDir, { recursive: true, force: true });

const stubEngine = path.join(root, 'test/_stubs/engine.ts');
const stubRouter = path.join(root, 'test/_stubs/router.ts');

const aliasPlugin = {
  name: 'tadori-test-alias',
  setup(b) {
    // node:test / node:assert/strict 等は node: プレフィックス付きのまま external に保つ
    // (esbuild 既定だと "node:test" → "test" に剥がされ require が失敗する)。
    b.onResolve({ filter: /^node:/ }, args => ({ path: args.path, external: true }));
    // engine / router は外部 I/O を持つので必ずスタブへ。テストも同じスタブを import
    // するので、絶対パスに正規化して単一インスタンスを共有させる。
    b.onResolve({ filter: /(^|\/)db\/engine$/ }, () => ({ path: stubEngine }));
    b.onResolve({ filter: /(^|\/)embeddings\/router$/ }, () => ({ path: stubRouter }));
    // DOMPurify を import 時に実行する mail 系は検索ロジックに不要 → no-op 化。
    b.onResolve({ filter: /(^|\/)(mailhtml|mailtext)$/ }, args => ({ path: args.path, namespace: 'stub-mail' }));
    b.onLoad({ filter: /.*/, namespace: 'stub-mail' }, () => ({
      contents: 'export function htmlToText(s){return String(s||"")}\nexport function cleanBody(s){return String(s||"")}',
      loader: 'js',
    }));
  },
};

const banner = `globalThis.localStorage ??= (()=>{const m=new Map();return {getItem:k=>m.has(k)?m.get(k):null,setItem:(k,v)=>m.set(k,String(v)),removeItem:k=>m.delete(k),clear:()=>m.clear(),key:i=>[...m.keys()][i]??null,get length(){return m.size}};})();`;

const tests = fs.readdirSync(path.join(root, 'test')).filter(f => f.endsWith('.test.ts'));
if (tests.length === 0) { console.error('test/*.test.ts が見つかりません'); process.exit(1); }

await esbuild.build({
  entryPoints: tests.map(f => path.join(root, 'test', f)),
  outdir: outDir,
  bundle: true,
  platform: 'node',
  format: 'cjs',
  target: 'es2022',
  outExtension: { '.js': '.cjs' },
  banner: { js: banner },
  plugins: [aliasPlugin],
  logLevel: 'error',
});

const files = fs.readdirSync(outDir).filter(f => f.endsWith('.test.cjs')).map(f => path.join(outDir, f));
try {
  execFileSync('node', ['--no-warnings', '--test', ...files], { stdio: 'inherit' });
} catch {
  process.exit(1); // テスト失敗時は非0で終了
}
