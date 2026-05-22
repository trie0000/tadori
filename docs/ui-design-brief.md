# Tadori UI デザインブリーフ（claude.design 用）

このファイルを claude.design にそのまま貼り、必要なら Spira の実画面スクショを
添付する。**目的: Spira と同一のデザイン言語で Tadori の画面を作る。**

---

## 0. 大前提

- Tadori は SharePoint / OWA のページ**上に被せて表示する注入型 UI**（独立アプリではない）。
- ホストページの CSS に汚染されないよう、ルート要素に `all: initial` を当て、
  超高 z-index（20億台）で最前面に出す。Spira と同じ方式。
- **デザイントークン（下記）は Spira からそのまま流用する。新しい色を作らない。**
  これが「Spira と同じ見た目」を保証する唯一の確実な方法。

## 1. デザイン言語（一言で）

「紙と墨と苔色」。暖色オフホワイトの紙地（paper）に、墨色（ink）の文字、
差し色は彩度を抑えた苔グリーン（moss green）。青を使わない暖かい影。
和文フォント（Meiryo 基準）。情報密度は高め、角丸は控えめ。

## 2. デザイントークン（Spira app.css より、そのままコピー可）

```css
:root {
  /* text */
  --ink: #2a2a26; --ink-3: #7a766c; --ink-4: #a8a39a;
  /* surface (paper) */
  --paper: #fafaf7; --paper-2: #f3f1ea; --paper-2-strong: #ece8de; --paper-3: #e8e4d8;
  /* line */
  --line: rgba(42,42,38,0.12); --line-strong: rgba(42,42,38,0.18);
  /* accent — moss green */
  --accent: #7a8a78; --accent-soft: rgba(122,138,120,0.18); --accent-strong: #5e6f5c;
  /* status */
  --danger: #b8534a; --warn: #c47f1c; --ok: #2f6f5e;
  --hl: rgba(196,174,96,0.35); /* 検索ハイライト用の山吹色 */
  /* type */
  --font-sans: "Meiryo","メイリオ","Hiragino Sans","Yu Gothic UI",-apple-system,"Segoe UI",system-ui,sans-serif;
  --font-mono: ui-monospace,"Cascadia Mono","Consolas",monospace;
  --fs-xs:11px; --fs-sm:12px; --fs-md:13px; --fs-base:15px; --fs-lg:16px; --fs-xl:18px;
  --fs-h3:22px; --fs-h2:28px; --fs-h1:36px;
  --lh-base:1.75; --lh-tight:1.35;
  /* spacing (4px base) */
  --s-1:4px; --s-2:6px; --s-3:8px; --s-4:10px; --s-5:12px; --s-6:14px; --s-7:18px; --s-8:22px; --s-9:28px; --s-10:40px;
  /* radius (控えめ) */
  --r-1:2px; --r-2:4px; --r-3:6px; --r-4:8px;
  /* shadow (暖色、青なし) */
  --shadow-panel: 0 8px 20px rgba(42,42,38,0.10);
  --shadow-modal: 0 0 0 1px rgba(42,42,38,0.06), 0 4px 12px rgba(42,42,38,0.10), 0 16px 40px rgba(42,42,38,0.18);
  /* chrome 寸法 */
  --topbar-h:44px; --side-w:200px; --toolbar-h:38px; --gutter:16px;
}
```

ダークモードあり（紙地 #1d1b18 / 墨 #e8e4d8 に反転）。フォントサイズは sm/中/lg の
3 段階を `data-font-size` 属性で切替（Spira と同じ）。

## 3. コンポーネント語彙（Spira と同じ命名・見た目）

- **ボタン**: primary（accent 塗り・白文字）/ ghost（枠線のみ）/ danger / icon。高さ 28〜32px、角丸 --r-2。
- **バッジ/チップ**: fill / ok / warn / danger / muted。小さめ、--fs-sm。
- **入力**: テキスト/セレクト/検索ボックス（左に虫眼鏡、paper-2 地）。
- **モーダル**: 暖色の三段影、backdrop は半透明。
- **トースト**: 右上、--z-toast。
- **テーブル/リスト**: 行区切りは --line、ホバーで paper-2。

## 4. Tadori で必要な画面（ここを claude.design に設計させる）

Spira は topbar+sidebar+main のフルシェルだが、Tadori は**検索が主役**なので
以下の 3 面を中心に。レイアウトは Spira の chrome（topbar/sidebar）を踏襲するか、
軽量パネル単体にするかは提案してほしい。

### (A) 検索パネル ★最重要
- 上部に大きめの検索ボックス（プレースホルダ「メール内容を意味で検索…」）
- 直下に件数・所要時間（例: 「1,240 件中 上位 20 件 / 38ms」）
- 結果リスト: 1 行 = 件名（強調）/ 送信者・日時（muted）/ 本文スニペット（--hl で
  クエリ周辺をハイライト）/ 右に「OWA で開く」リンク。
- 各行クリックで詳細プレビュー（右ペイン or 展開）。

### (B) 取り込み状況（Sticky モード）
- 小さなステータスバッジ: 稼働中/停止、claimed/embedded/indexed の件数、最終同期時刻。
- トグルで取り込みループの ON/OFF。

### (C) 設定モーダル
- 中継サーバ URL、デプロイ名、API キー、対象 ML アドレス、次元数、同期間隔。
- Spira の設定モーダル（settingsModal）と同じ見た目に。

## 5. claude.design への指示文（コピペ用）

> 添付のデザイントークンとスクリーンショットに完全準拠して、「Tadori」という
> SharePoint 上に注入されるメール意味検索ツールの UI を設計してください。配色・
> フォント・余白・角丸・影は添付トークンの CSS 変数だけを使い、新しい色や
> フォントを足さないこと。和文（日本語）UI。画面は (A) 検索パネル、(B) 取り込み
> 状況バッジ、(C) 設定モーダル の 3 つ。トーンは Spira と同じ「紙と墨と苔色」。
> 出力は各画面の HTML + これらの CSS 変数を参照する CSS で。
