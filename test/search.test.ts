// 本番 searchVectors を engine/router だけスタブして実行する E2E 検証。
// 過去に発生した不具合を回帰テストとして固定 + 種別ごとのサブ項目スコープ (＋ピッカー) を検証:
//  - スコープ空で全件が消える (2026-06 不具合) を各種別で防止
//  - kind フィルタ / doc・pptx フォルダ絞り込み (一致/不一致)
//  - メール= to/cc アドレス絞り込み / OneNote= ラベル絞り込み

import test from 'node:test';
import assert from 'node:assert/strict';

import { searchVectors } from '../src/search/vectorSearch';
import { VectorDb } from '../src/db/store';
import { normalize } from '../src/search/cosine';
import { __setDb } from './_stubs/engine';
import { __setQuery } from './_stubs/router';
import { makeRecord, makeSegment, vec, DIM } from './_fixtures';

const FOLDER = '/sites/n365/Shared Documents/マニュアル';
const S: any = { ragKeywordWeight: 0, relayBaseUrl: '' };

function setupDb(): void {
  const db = new VectorDb();
  const recs = [];
  for (let i = 0; i < 10; i++) recs.push(makeRecord({ i, kind: 'doc', folder: FOLDER }));
  for (let i = 0; i < 10; i++) recs.push(makeRecord({ i, kind: 'pptx', folder: '/sites/n365/Shared Documents/PPTX' }));
  db.applySegment(makeSegment('seg-00000', recs));
  __setDb(db);
  __setQuery(normalize(vec(3 + 1000)));   // doc#3 に一致
}

test('doc スコープ空 [] でも doc がヒットする (回帰: 空で全消え)', async () => {
  setupDb();
  const hits = await searchVectors('マニュアル', S, 'site', 10, { kinds: ['doc'], scope: { folders: [] } });
  assert.ok(hits.length > 0, '空スコープで 0 件になってはいけない');
  assert.ok(hits.every(h => h.kind === 'doc'));
  assert.equal(hits[0].docFile, 'f3.pdf');
});

test('kind フィルタ=pptx は pptx のみ', async () => {
  setupDb();
  const hits = await searchVectors('マニュアル', S, 'site', 10, { kinds: ['pptx'] });
  assert.ok(hits.length > 0);
  assert.ok(hits.every(h => h.kind === 'pptx'));
});

test('doc フォルダ絞り込み: 一致は通る / 不一致は0件', async () => {
  setupDb();
  const ok = await searchVectors('マニュアル', S, 'site', 10, { kinds: ['doc'], scope: { folders: [FOLDER] } });
  assert.ok(ok.length > 0 && ok.every(h => h.kind === 'doc'));
  const ng = await searchVectors('マニュアル', S, 'site', 10, { kinds: ['doc'], scope: { folders: ['/sites/n365/Shared Documents/別'] } });
  assert.equal(ng.length, 0);
});

test('pptx フォルダ絞り込み: 一致フォルダのみ', async () => {
  setupDb();
  __setQuery(normalize(vec(2 + 5000)));   // pptx#2 に一致
  const hits = await searchVectors('x', S, 'site', 10, { kinds: ['pptx'], scope: { folders: ['/sites/n365/Shared Documents/PPTX'] } });
  assert.ok(hits.length > 0 && hits.every(h => h.kind === 'pptx'));
});

test('メール: to/cc アドレスで絞り込み', async () => {
  const db = new VectorDb();
  const recs = [];
  for (let i = 0; i < 5; i++) recs.push(makeRecord({ i, kind: 'mail', to: ['alpha@example.com'], cc: [] }));
  for (let i = 5; i < 10; i++) recs.push(makeRecord({ i, kind: 'mail', to: ['beta@example.com'], cc: [] }));
  db.applySegment(makeSegment('seg-00000', recs));
  __setDb(db);
  __setQuery(normalize(vec(7))); // beta 側 (i=7)
  const hits = await searchVectors('x', S, 'site', 10, { kinds: ['mail'], scope: { mailAddresses: ['beta@example.com'] } });
  assert.ok(hits.length > 0, 'beta 宛が取れるべき');
  assert.equal(hits[0].messageId.includes('f7'), true, 'beta#7 がトップ');
});

test('OneNote: ラベル(pageId 集合)で絞り込み', async () => {
  const F = '/sites/x/Shared Documents/F';
  const db = new VectorDb();
  const recs = [];
  for (let i = 0; i < 5; i++) recs.push(makeRecord({ i, kind: 'onenote', label: '議事録' }));
  for (let i = 5; i < 10; i++) recs.push(makeRecord({ i, kind: 'onenote', label: '仕様メモ' }));
  db.applySegment(makeSegment('seg-00000', recs));
  __setDb(db);
  __setQuery(normalize(vec(8 + 9000)));
  // 「仕様メモ」ラベル = pageId(conversationId) i=5..9。チャット側がバッチ設定から解決する想定。
  const specPageIds = [5, 6, 7, 8, 9].map(i => `${F}/f${i}`);
  const hits = await searchVectors('x', S, 'site', 10, { kinds: ['onenote'], scope: { onenotePageIds: specPageIds } });
  assert.ok(hits.length > 0);
  assert.ok(hits.every(h => specPageIds.includes(h.conversationId)), '選んだラベルのページのみ');
});

test('スコープ未指定なら全件対象', async () => {
  setupDb();
  const hits = await searchVectors('マニュアル', S, 'site', 20, {});
  const kinds = new Set(hits.map(h => h.kind));
  assert.ok(kinds.has('doc'));
});
