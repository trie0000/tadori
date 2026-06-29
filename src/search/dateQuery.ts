// 質問文から期間 (from..to, いずれも 'YYYY-MM-DD') を解釈する。
// 「今月の」「先週の」「直近30日」「2024年5月」等を検索の日付フィルタに変換する。
// 純関数 (now を引数で受ける) なのでテスト可能。判定は UTC 基準で行う (決定論のため)。

export interface DateRange { from?: string; to?: string; }

function ymd(ms: number): string {
  const d = new Date(ms);
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${d.getUTCFullYear()}-${m}-${day}`;
}
function ymdOf(y: number, m1: number, d: number): string {
  return `${y}-${String(m1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}
function lastDayOfMonth(y: number, m1: number): number {
  return new Date(Date.UTC(y, m1, 0)).getUTCDate(); // m1 は 1始まり、Date.UTC(y, m1, 0) = 前月末日
}

const DAY = 86400_000;
const z2h = (s: string): string => s.replace(/[０-９]/g, c => String('０１２３４５６７８９'.indexOf(c)));

/** query に期間表現があれば DateRange を返す。無ければ null。 */
export function parseDateRange(query: string, nowMs: number): DateRange | null {
  const q = z2h(query);
  const d = new Date(nowMs);
  const Y = d.getUTCFullYear(), M = d.getUTCMonth() + 1;
  const today = ymd(nowMs);

  // 直近/過去/ここ N 日|週(間)|ヶ月|年
  const rel = q.match(/(?:直近|過去|ここ)\s*(\d+)\s*(日|週間|週|ヶ月|カ月|か月|月|年)/);
  if (rel) {
    const n = parseInt(rel[1], 10), unit = rel[2];
    let fromMs = nowMs;
    if (unit === '日') fromMs = nowMs - n * DAY;
    else if (unit === '週' || unit === '週間') fromMs = nowMs - n * 7 * DAY;
    else if (unit === '年') fromMs = Date.UTC(Y - n, d.getUTCMonth(), d.getUTCDate());
    else fromMs = Date.UTC(Y, d.getUTCMonth() - n, d.getUTCDate()); // ヶ月/月
    return { from: ymd(fromMs), to: today };
  }

  // YYYY年M月 / YYYY年
  const ym = q.match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
  if (ym) {
    const y = +ym[1], m = +ym[2];
    return { from: ymdOf(y, m, 1), to: ymdOf(y, m, lastDayOfMonth(y, m)) };
  }
  const yOnly = q.match(/(\d{4})\s*年(?!\s*\d)/);
  if (yOnly) { const y = +yOnly[1]; return { from: ymdOf(y, 1, 1), to: ymdOf(y, 12, 31) }; }

  if (/今日|本日/.test(q)) return { from: today, to: today };
  if (/昨日/.test(q)) return { from: ymd(nowMs - DAY), to: ymd(nowMs - DAY) };
  if (/今週/.test(q)) {
    const dow = (new Date(nowMs).getUTCDay() + 6) % 7; // 月曜=0
    return { from: ymd(nowMs - dow * DAY), to: today };
  }
  if (/先週|前週/.test(q)) {
    const dow = (new Date(nowMs).getUTCDay() + 6) % 7;
    const thisMon = nowMs - dow * DAY;
    return { from: ymd(thisMon - 7 * DAY), to: ymd(thisMon - DAY) };
  }
  if (/今月/.test(q)) return { from: ymdOf(Y, M, 1), to: today };
  if (/先月|前月/.test(q)) {
    const py = M === 1 ? Y - 1 : Y, pm = M === 1 ? 12 : M - 1;
    return { from: ymdOf(py, pm, 1), to: ymdOf(py, pm, lastDayOfMonth(py, pm)) };
  }
  if (/今年/.test(q)) return { from: ymdOf(Y, 1, 1), to: today };
  if (/去年|昨年/.test(q)) return { from: ymdOf(Y - 1, 1, 1), to: ymdOf(Y - 1, 12, 31) };

  return null;
}

/** record.date (ISO) が [from,to] (両端含む) に入るか。日付の無いレコードは安全側で true。 */
export function inDateRange(dateIso: string | undefined, range: DateRange): boolean {
  if (!dateIso) return true;                 // 日付不明は除外しない
  const day = dateIso.slice(0, 10);
  if (range.from && day < range.from) return false;
  if (range.to && day > range.to) return false;
  return true;
}
