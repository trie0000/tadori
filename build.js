// Tadori build script — esbuild + dev server + bookmarklet 配布物生成
// (Spira の build.js と同じ配布方式: install.html ドラッグ登録 + loader 方式)
import * as esbuild from 'esbuild';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';

const watch = process.argv.includes('--watch');
const serve = process.argv.includes('--serve');
const makeBookmarklet = process.argv.includes('--bookmarklet');
// install.html / index.html はユーザ配布物なので prod は minify する
// (Edge の drag-to-bookmark 上限が 1〜2MB 程度のため)。
const prod = !watch && !serve;

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
// ビルド時刻は JST 表記 (どのビルドが動いているか日本時間で判別できるように)。
const buildTime = new Date().toLocaleString('sv-SE', { timeZone: 'Asia/Tokyo' }) + ' JST';
const buildId = `${pkg.version}-${gitSha}${gitDirty} (${buildTime})`;
console.log(`[build] id: ${buildId}`);

// (pdf.js は bundle 肥大化のため一旦撤去。PDF 対応は別方式で再導入予定。)

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
  logLevel: 'info',
};

if (watch || serve) {
  const ctx = await esbuild.context(buildOptions);
  await ctx.watch();
  console.log('[esbuild] watching...');
  if (serve) {
    const PORT = 5599;
    http.createServer((req, res) => {
      let url = (req.url || '/').split('?')[0];
      if (url === '/') url = '/test/harness.html';
      const filePath = path.join(process.cwd(), url);
      if (!fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
        res.writeHead(404); res.end('not found'); return;
      }
      const types = {
        '.html': 'text/html; charset=utf-8',
        '.js': 'application/javascript; charset=utf-8',
        '.css': 'text/css; charset=utf-8',
        '.json': 'application/json; charset=utf-8',
        '.map': 'application/json; charset=utf-8',
      };
      res.writeHead(200, { 'Content-Type': types[path.extname(filePath)] || 'text/plain; charset=utf-8', 'Access-Control-Allow-Origin': '*' });
      fs.createReadStream(filePath).pipe(res);
    }).listen(PORT, () => console.log(`[dev] http://localhost:${PORT}/  (→ test/harness.html)`));
  }
} else {
  await esbuild.build(buildOptions);
  console.log('[esbuild] build complete');

  const js = fs.readFileSync('dist/tadori.js', 'utf8');
  const sizeKb = (s) => (fs.statSync(s).size / 1024).toFixed(1);

  // 単一ファイル HTML: SharePoint ドキュメントライブラリに置いて直接開く用途。
  fs.writeFileSync('dist/index.html', `<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Tadori</title>
<style>html,body{margin:0;padding:0;background:#fafaf7}</style>
</head>
<body>
<script>${js}
</script>
</body>
</html>`);
  console.log(`[html] dist/index.html: ${sizeKb('dist/index.html')} KB`);

  // install.html: バンドル全体を javascript: URL に inline したドラッグ登録ページ。
  // void function(){...}() で包み、再クリックでもグローバル汚染しないようにする
  // (再マウントの冪等性は shell.boot() が担保)。
  const inlined = `void function(){${js}}()`;
  const bookmarkletHref = 'javascript:' + encodeURIComponent(inlined);
  fs.writeFileSync('dist/install.html', renderInstallHtml(bookmarkletHref, false));
  console.log(`[install] dist/install.html: ${sizeKb('dist/install.html')} KB (バンドル inline)`);

  if (makeBookmarklet) {
    // ── ローダー方式 (自動更新) ─────────────────────────────────────────
    // 配置先: SharePoint「ドキュメント」(Shared Documents) の Tadori フォルダ。
    // ローダーが実行ページのサイト URL から相対でパスを組み立てるので、同じ
    // ライブラリ/フォルダ命名ならどのサイトでも動く。
    const libPath = process.env.TADORI_BUNDLE_LIB || '/Shared%20Documents/Tadori';
    const overrideBase = JSON.stringify(process.env.TADORI_BUNDLE_BASE || '');

    fs.copyFileSync('dist/tadori.js', 'dist/tadori.bundle.js');
    fs.writeFileSync('dist/version.txt', buildId + '\n');

    // ローダー: ①サイト URL+libPath でベース決定 ②version.txt を毎回 fetch して
    // 最新版確認 ③tadori.bundle.js?v=<version> を <script> 注入 (版が同じなら
    // ブラウザキャッシュ即ロード) ④失敗時は SharePoint に自動フォールバック。
    const loader =
      `(function(){var d=document,w=window;` +
      `function SP(){try{var c=w._spPageContextInfo;if(c&&c.webServerRelativeUrl)return c.webServerRelativeUrl.replace(/\\/$/,'')+${JSON.stringify(libPath)};}catch(e){}return '';}` +
      `var sp=SP(),dev='';` +
      `try{if(w.localStorage&&localStorage.getItem('tadori.dev.bundle-source')==='local')dev=(localStorage.getItem('tadori.dev.local-base')||'http://127.0.0.1:18080/tadori').replace(/\\/+$/,'');}catch(e){}` +
      `var primary=dev||${overrideBase}||sp;var fb=(primary!==sp&&sp)?sp:'';var isLocal=!!dev;` +
      // ローカル参照失敗はサイレントに SP へ落とさず警告。
      `function fail(base,why){var msg='[Tadori] ローカルバンドル読み込み失敗: '+base+(why?' ('+why+')':'')+'\\nrelay 起動 / CORS / PNA / CSP を確認してください。';if(isLocal){alert(msg);console.error(msg);}else{console.warn(msg);}}` +
      // SP の CSP は http://127.0.0.1 を script-src に許可していないので <script src> 注入は
      // ブロックされる。ローカル参照時のみ fetch で取得して unsafe-eval で実行する
      // (SP の CSP には unsafe-eval が含まれているため通る)。SP 配信時は同一オリジンなので
      // 従来どおり <script src> でロードする。
      `function evalLoad(base,ver){fetch(base+'/tadori.bundle.js?v='+encodeURIComponent(ver),{credentials:'same-origin'}).then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r.text();}).then(function(t){var o=d.getElementById('tadori-script');if(o)o.remove();try{(0,eval)(t);}catch(e){fail(base,'eval: '+(e&&e.message||e));}}).catch(function(e){fail(base,e&&e.message||'fetch error');});}` +
      `function inject(base,ver){if(isLocal){evalLoad(base,ver);return;}var o=d.getElementById('tadori-script');if(o)o.remove();var s=d.createElement('script');s.id='tadori-script';s.src=base+'/tadori.bundle.js?v='+encodeURIComponent(ver);s.onerror=function(){fail(base,'script load error');if(fb){var x=fb;fb='';go(x);}};d.body.appendChild(s);}` +
      `function go(base){fetch(base+'/version.txt?t='+Date.now(),{credentials:'same-origin'}).then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r.text();}).then(function(t){inject(base,(t||'').trim()||String(Date.now()));}).catch(function(e){fail(base,e&&e.message||'fetch error');if(isLocal)return;if(fb){var x=fb;fb='';go(x);}else{inject(base,String(Date.now()));}});}` +
      `go(primary);})();`;

    fs.writeFileSync('dist/tadori.loader.js', loader);
    const loaderHref = 'javascript:' + encodeURIComponent(loader);
    fs.writeFileSync('dist/bookmarklet.txt', loaderHref);
    fs.writeFileSync('dist/install-loader.html', renderInstallHtml(loaderHref, true));

    console.log(`[loader] dist/tadori.bundle.js: ${sizeKb('dist/tadori.bundle.js')} KB  ← ライブラリに配置`);
    console.log(`[loader] dist/version.txt: "${buildId}"  ← ライブラリに配置`);
    console.log(`[loader] dist/tadori.loader.js / bookmarklet.txt / install-loader.html`);
    console.log('');
    console.log('  ▶ 配置: SharePoint「ドキュメント」→ Tadori フォルダに tadori.bundle.js と version.txt を置く');
    console.log(`     ローダーは実行ページのサイト URL + "${libPath}" から自動でパスを組み立てます。`);
    console.log('  ▶ ライブラリ/フォルダが違う場合は TADORI_BUNDLE_LIB で上書き:');
    console.log('     TADORI_BUNDLE_LIB="/SiteAssets/tadori" node build.js --bookmarklet');
  }
}

function renderInstallHtml(bookmarkletHref, loaderMode) {
  const updateNote = loaderMode
    ? '<strong>更新方法</strong>: 本体 (tadori.bundle.js) を新ビルドに差し替えるだけ。ブックマークの再登録は不要です (ローダーが version.txt を見て自動更新)。'
    : '<strong>更新方法</strong>: 新しいバージョンが出たら、このページを再度開いて再度ドラッグしてください (古いブックマークは上書き or 削除)。';
  return `<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Tadori インストール</title>
<style>
  :root {
    --ink: #2a2a26; --ink-3: #7a766c; --ink-4: #a8a39a;
    --paper: #fafaf7; --paper-2: #f3f1ea; --paper-3: #e8e4d8;
    --line: rgba(42,42,38,0.12);
    --accent: #7a8a78; --accent-strong: #5e6f5c;
    --code-fg: #8b3a30; --code-bg: rgba(122,118,108,0.16);
    --font: "Meiryo","メイリオ","Hiragino Sans","Yu Gothic UI",-apple-system,"Segoe UI",system-ui,sans-serif;
    --font-mono: ui-monospace,"Cascadia Mono","Consolas",monospace;
  }
  * { box-sizing: border-box; }
  body { font-family: var(--font); max-width: 580px; margin: 60px auto; padding: 0 24px; color: var(--ink); line-height: 1.75; background: var(--paper); }
  h1 { font-size: 28px; font-weight: 700; margin: 0 0 8px; letter-spacing: -0.01em; display: flex; align-items: center; gap: 12px; }
  h1 .mark { font-family: var(--font-mono); color: var(--accent-strong); }
  .sub { color: var(--ink-3); font-size: 14px; margin: 0 0 40px; }
  .step { display: flex; gap: 16px; margin-bottom: 28px; align-items: flex-start; }
  .step-num { width: 28px; height: 28px; background: var(--accent); color: #fff; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 13px; font-weight: 700; flex-shrink: 0; margin-top: 2px; }
  .step-body h3 { font-size: 16px; font-weight: 600; margin: 0 0 4px; color: var(--ink); }
  .step-body p { font-size: 14px; color: var(--ink-3); margin: 0; }
  .bm-wrap { background: var(--paper-2); border: 2px dashed var(--paper-3); border-radius: 8px; padding: 28px; text-align: center; margin: 20px 0 32px; }
  .bm-wrap p { font-size: 13px; color: var(--ink-3); margin: 0 0 16px; }
  #bm-link {
    display: inline-flex; align-items: center; gap: 8px;
    background: var(--accent); color: #fff; text-decoration: none;
    padding: 12px 24px; border-radius: 6px; font-size: 16px; font-weight: 600;
    box-shadow: 0 2px 8px rgba(122,138,120,.25); cursor: grab; user-select: none;
  }
  #bm-link:hover { background: var(--accent-strong); }
  #bm-link .mark { font-family: var(--font-mono); }
  hr { border: none; border-top: 1px solid var(--paper-3); margin: 32px 0; }
  .alt { font-size: 13px; color: var(--ink-3); }
  code { background: var(--code-bg); color: var(--code-fg); padding: 2px 6px; border-radius: 3px; font-size: 12px; font-family: var(--font-mono); }
  .note { background: rgba(196,127,28,0.10); border-left: 3px solid #c47f1c; padding: 12px 16px; border-radius: 4px; font-size: 13px; color: var(--ink); margin-top: 24px; }
</style>
</head>
<body>

<h1><span class="mark">辿</span> Tadori インストール</h1>
<p class="sub">SharePoint 上で動く ML メール検索 (RAG) — bookmarklet 形式${loaderMode ? ' / ローダー方式' : ''}</p>

<div class="step">
  <div class="step-num">1</div>
  <div class="step-body">
    <h3>ブックマークバーを表示する</h3>
    <p>Chrome / Edge: <code>Ctrl+Shift+B</code>（Mac: <code>Cmd+Shift+B</code>）</p>
  </div>
</div>

<div class="step">
  <div class="step-num">2</div>
  <div class="step-body">
    <h3>下のボタンをブックマークバーにドラッグ</h3>
    <p>右クリック → 「リンクをブックマーク」 でも OK です。</p>
  </div>
</div>

<div class="bm-wrap">
  <p>↓ このボタンをブックマークバーにドラッグ ↓</p>
  <a id="bm-link" href="${bookmarkletHref}" onclick="alert('ドラッグしてブックマークバーに登録してください。クリックでは起動しません（このページは SharePoint ではないため）。'); return false;"><span class="mark">辿</span> Tadori</a>
</div>

<div class="step">
  <div class="step-num">3</div>
  <div class="step-body">
    <h3>SharePoint サイトを開いて、ブックマークをクリック</h3>
    <p>同一テナントの SP サイト上でブックマークを実行すると Tadori が起動します。<br>
       初回は 歯車 → 設定 → 取り込みで List 名、AI 接続で中継サーバ / モデルを設定してください。</p>
  </div>
</div>

<hr>

<p class="alt">${updateNote}</p>

<div class="note">
  ⚠ <strong>SharePoint 上で実行する必要があります</strong>。Graph API・外部 SaaS 不要で、SP REST API（同一オリジン Cookie 認証）のみを使用します。
</div>

</body>
</html>`;
}
