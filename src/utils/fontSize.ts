// 文字サイズ設定 (小 / 中 / 大)。Spira の utils/fontSize.ts と同じ作法。
// 中 = 現在の既定値、小/大 はスケール表を #tadori-root にインラインで上書き。

export type FontSize = 'sm' | 'md' | 'lg';

const KEY = 'tadori:font-size';

export function getFontSize(): FontSize {
  try {
    const v = localStorage.getItem(KEY);
    if (v === 'sm' || v === 'lg' || v === 'md') return v;
  } catch { /* ignore */ }
  return 'md';
}

export function setFontSize(v: FontSize): void {
  try {
    if (v === 'md') localStorage.removeItem(KEY);
    else localStorage.setItem(KEY, v);
  } catch { /* ignore */ }
  applyFontSize();
}

// 中 = 現在の app.css と同じ値。小/大 はそこから前後にスケール。
const SCALES: Record<FontSize, Record<string, string>> = {
  sm: { '--fs-xs': '10px', '--fs-sm': '11px', '--fs-md': '12px', '--fs-base': '13px', '--fs-lg': '14px', '--fs-xl': '16px' },
  md: { '--fs-xs': '11px', '--fs-sm': '12px', '--fs-md': '13px', '--fs-base': '15px', '--fs-lg': '16px', '--fs-xl': '18px' },
  lg: { '--fs-xs': '13px', '--fs-sm': '14px', '--fs-md': '15px', '--fs-base': '17px', '--fs-lg': '19px', '--fs-xl': '22px' },
};

/** #tadori-root に CSS 変数 + data-font-size 属性 + base font-size を適用。 */
export function applyFontSize(): void {
  const root = document.getElementById('tadori-root');
  if (!root) return;
  const size = getFontSize();
  const scale = SCALES[size];
  for (const [k, v] of Object.entries(scale)) root.style.setProperty(k, v);
  root.setAttribute('data-font-size', size);
  root.style.fontSize = scale['--fs-base']!;
}
