// Tadori build script — esbuild + dev server (Spira の build.js 流派を踏襲)
import * as esbuild from 'esbuild';
import http from 'node:http';
import fs from 'node:fs';
import { execSync } from 'node:child_process';

const watch = process.argv.includes('--watch');
const serve = process.argv.includes('--serve');
const prod = process.argv.includes('--prod') || (!watch && !serve);

// Build identity — 「どのビルドが動いているか」をバンドルに焼き込む。
const pkg = JSON.parse(fs.readFileSync('package.json', 'utf8'));
let gitSha = 'nogit';
let gitDirty = '';
try {
  gitSha = execSync('git rev-parse --short HEAD', { stdio: ['ignore', 'pipe', 'ignore'] })
    .toString().trim();
  if (execSync('git status --porcelain', { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim()) {
    gitDirty = '+';
  }
} catch { /* not a git repo */ }
const buildTime = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
const buildId = `${pkg.version}-${gitSha}${gitDirty} (${buildTime})`;
console.log(`[build] id: ${buildId}`);

const buildOptions = {
  entryPoints: ['src/main.ts'],
  bundle: true,
  format: 'iife',
  globalName: 'Tadori',
  outfile: 'dist/tadori.js',
  target: 'es2020',
  platform: 'browser',
  minify: prod,
  sourcemap: !prod,
  loader: { '.css': 'text' },
  define: {
    'process.env.NODE_ENV': prod ? '"production"' : '"development"',
    __TADORI_BUILD_ID__: JSON.stringify(buildId),
    __TADORI_VERSION__: JSON.stringify(pkg.version),
  },
};

if (watch || serve) {
  const ctx = await esbuild.context(buildOptions);
  await ctx.watch();
  console.log('[build] watching...');
  if (serve) {
    const PORT = 5599;
    http.createServer((req, res) => {
      const file = req.url === '/' ? '/dist/tadori.js' : req.url;
      try {
        const body = fs.readFileSync('.' + file);
        res.writeHead(200, { 'Access-Control-Allow-Origin': '*' });
        res.end(body);
      } catch {
        res.writeHead(404); res.end('not found');
      }
    }).listen(PORT, () => console.log(`[serve] http://localhost:${PORT}/dist/tadori.js`));
  }
} else {
  await esbuild.build(buildOptions);
  console.log(`[build] done → dist/tadori.js (minify=${prod})`);
}
