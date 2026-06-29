// 期間解釈 (parseDateRange) と範囲判定 (inDateRange) の検証。
// now は固定 (UTC 2026-06-29 = 月曜) で決定論的に。
import test from 'node:test';
import assert from 'node:assert/strict';
import { parseDateRange, inDateRange } from '../src/search/dateQuery';

const NOW = Date.UTC(2026, 5, 29, 12, 0, 0); // 2026-06-29 (月)

test('今日 / 昨日', () => {
  assert.deepEqual(parseDateRange('今日のメール', NOW), { from: '2026-06-29', to: '2026-06-29' });
  assert.deepEqual(parseDateRange('昨日の議事録', NOW), { from: '2026-06-28', to: '2026-06-28' });
});

test('今月 / 先月', () => {
  assert.deepEqual(parseDateRange('今月のお知らせ', NOW), { from: '2026-06-01', to: '2026-06-29' });
  assert.deepEqual(parseDateRange('先月の申請', NOW), { from: '2026-05-01', to: '2026-05-31' });
});

test('今年 / 去年', () => {
  assert.deepEqual(parseDateRange('今年の方針', NOW), { from: '2026-01-01', to: '2026-06-29' });
  assert.deepEqual(parseDateRange('去年の実績', NOW), { from: '2025-01-01', to: '2025-12-31' });
});

test('直近N日 / 過去Nヶ月', () => {
  assert.deepEqual(parseDateRange('直近7日の変更', NOW), { from: '2026-06-22', to: '2026-06-29' });
  assert.deepEqual(parseDateRange('過去3ヶ月の傾向', NOW), { from: '2026-03-29', to: '2026-06-29' });
});

test('YYYY年M月 / YYYY年', () => {
  assert.deepEqual(parseDateRange('2024年5月の資料', NOW), { from: '2024-05-01', to: '2024-05-31' });
  assert.deepEqual(parseDateRange('2023年のまとめ', NOW), { from: '2023-01-01', to: '2023-12-31' });
});

test('今週 (月曜起点)', () => {
  // 2026-06-29 は月曜 → 週初は当日
  assert.deepEqual(parseDateRange('今週の予定', NOW), { from: '2026-06-29', to: '2026-06-29' });
});

test('期間表現が無ければ null', () => {
  assert.equal(parseDateRange('経費精算のやり方', NOW), null);
});

test('inDateRange: 範囲内/外/日付なし', () => {
  const r = { from: '2026-06-01', to: '2026-06-30' };
  assert.equal(inDateRange('2026-06-15T09:00:00Z', r), true);
  assert.equal(inDateRange('2026-05-31T09:00:00Z', r), false);
  assert.equal(inDateRange('2026-07-01T00:00:00Z', r), false);
  assert.equal(inDateRange(undefined, r), true); // 日付不明は通す
});
