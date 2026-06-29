import test from 'node:test';
import assert from 'node:assert/strict';
import { computeFacets } from '../src/search/facets';

const hits = [
  { kind: 'mail', date: '2026-06-10T00:00:00Z', from: '田中' },
  { kind: 'mail', date: '2026-06-20T00:00:00Z', from: '田中' },
  { kind: 'mail', date: '2026-05-02T00:00:00Z', from: '佐藤' },
  { kind: 'doc',  date: '2026-06-01T00:00:00Z', from: 'manual.pdf' },
  { kind: 'onenote', from: 'NB › Sec' },   // 日付なし
];

test('種別の集計 (件数降順)', () => {
  const f = computeFacets(hits);
  assert.equal(f.kind[0].key, 'mail');
  assert.equal(f.kind[0].count, 3);
  assert.equal(f.kind[0].label, 'メール');
});

test('時期 (YYYY-MM) の集計 / 日付なしは除外', () => {
  const f = computeFacets(hits);
  const jun = f.month.find(b => b.key === '2026-06');
  assert.equal(jun?.count, 3);
  assert.equal(f.month.reduce((n, b) => n + b.count, 0), 4); // 日付なし1件は除外
});

test('差出人の集計', () => {
  const f = computeFacets(hits);
  assert.equal(f.from.find(b => b.key === '田中')?.count, 2);
});
