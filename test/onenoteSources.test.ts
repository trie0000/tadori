// OneNote ラベルバッチ: コンテナ(ノート/セクション)+ページの和集合マージ / 旧形式移行 / pageIds 差替。
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  recordOneNoteBatch, listOneNoteBatches, setOneNoteBatchPageIds, removeOneNoteBatch, renameOneNoteBatch,
} from '../src/sync/onenoteSources';

const SITE = 'https://x.sharepoint.com/sites/test-onenote';

function reset(): void { removeOneNoteBatch(SITE, '設計'); }

test('新規バッチ: ノート/セクション/ページを記録', () => {
  reset();
  recordOneNoteBatch(SITE, '設計', { notebookIds: ['nb1'], sectionIds: ['s1'], pageIds: ['p1', 'p2'] });
  const b = listOneNoteBatches(SITE).find(x => x.label === '設計')!;
  assert.deepEqual(b.notebookIds, ['nb1']);
  assert.deepEqual(b.sectionIds, ['s1']);
  assert.deepEqual(b.pageIds, ['p1', 'p2']);
});

test('同ラベルに追記すると和集合マージ (複数ノート/重複OK)', () => {
  reset();
  recordOneNoteBatch(SITE, '設計', { notebookIds: ['nb1'], sectionIds: ['s1'], pageIds: ['p1'] });
  recordOneNoteBatch(SITE, '設計', { notebookIds: ['nb2'], sectionIds: ['s1', 's2'], pageIds: ['p1', 'p3'] });
  const b = listOneNoteBatches(SITE).find(x => x.label === '設計')!;
  assert.deepEqual(b.notebookIds.sort(), ['nb1', 'nb2']);
  assert.deepEqual(b.sectionIds.sort(), ['s1', 's2']);   // 重複 s1 は1つに
  assert.deepEqual(b.pageIds.sort(), ['p1', 'p3']);       // 重複 p1 は1つに
});

test('setOneNoteBatchPageIds で解決ページを差し替え', () => {
  reset();
  recordOneNoteBatch(SITE, '設計', { sectionIds: ['s1'], pageIds: ['p1'] });
  setOneNoteBatchPageIds(SITE, '設計', ['p1', 'p2', 'p3']);
  const b = listOneNoteBatches(SITE).find(x => x.label === '設計')!;
  assert.deepEqual(b.pageIds.sort(), ['p1', 'p2', 'p3']);
  assert.deepEqual(b.sectionIds, ['s1']); // コンテナは維持
});

test('renameOneNoteBatch: 新名へ変更 / 既存名ならマージ', () => {
  removeOneNoteBatch(SITE, 'A'); removeOneNoteBatch(SITE, 'B'); removeOneNoteBatch(SITE, 'C');
  recordOneNoteBatch(SITE, 'A', { sectionIds: ['s1'], pageIds: ['p1'] });
  renameOneNoteBatch(SITE, 'A', 'C');
  let l = listOneNoteBatches(SITE);
  assert.ok(!l.find(b => b.label === 'A') && l.find(b => b.label === 'C'));
  // 既存名へリネーム → マージ
  recordOneNoteBatch(SITE, 'B', { sectionIds: ['s2'], pageIds: ['p2'] });
  renameOneNoteBatch(SITE, 'B', 'C');
  l = listOneNoteBatches(SITE);
  const c = l.find(b => b.label === 'C')!;
  assert.ok(!l.find(b => b.label === 'B'));
  assert.deepEqual(c.sectionIds.sort(), ['s1', 's2']);
  assert.deepEqual(c.pageIds.sort(), ['p1', 'p2']);
  removeOneNoteBatch(SITE, 'C');
});

test('旧形式 (pageIds のみ) を読み込んでも欠損フィールドを補完', () => {
  // 旧スキーマを直接 localStorage に書いて load 経由で読む
  const { siteHash } = require('../src/sharepoint/spSites');
  const key = `tadori:onenote:batches:${siteHash(SITE)}`;
  localStorage.setItem(key, JSON.stringify([{ label: '旧', pageIds: ['a', 'b'] }]));
  const b = listOneNoteBatches(SITE).find(x => x.label === '旧')!;
  assert.deepEqual(b.notebookIds, []);
  assert.deepEqual(b.sectionIds, []);
  assert.deepEqual(b.pageIds, ['a', 'b']);
  removeOneNoteBatch(SITE, '旧');
});
