// 本番 searchVectors を engine/router だけスタブして実行する E2E 検証。
// 過去に発生した不具合を回帰テストとして固定する:
//  - doc スコープ空配列で doc 全件が消える (2026-06 不具合)
//  - kind フィルタ
//  - doc フォルダ絞り込み (一致/不一致)

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
  // クエリ = doc#3 に一致
  __setQuery(normalize(vec(3 + 1000)));
}

test('doc スコープ空配列 [] でも doc がヒットする (回帰: 空配列で全消え)', async () => {
  setupDb();
  const hits = await searchVectors('マニュアル', S, 'site', 10, { kinds: ['doc'], docFolderPrefixes: [] });
  assert.ok(hits.length > 0, '空配列スコープで 0 件になってはいけない');
  assert.ok(hits.every(h => h.kind === 'doc'), 'doc 以外が混ざってはいけない');
  assert.equal(hits[0].docFile, 'f3.pdf', 'トップは一致した doc#3 のはず');
});

test('kind フィルタ=pptx は pptx のみ返す', async () => {
  setupDb();
  const hits = await searchVectors('マニュアル', S, 'site', 10, { kinds: ['pptx'] });
  assert.ok(hits.length > 0);
  assert.ok(hits.every(h => h.kind === 'pptx'));
});

test('doc フォルダ絞り込み: 一致フォルダは通る', async () => {
  setupDb();
  const hits = await searchVectors('マニュアル', S, 'site', 10, { kinds: ['doc'], docFolderPrefixes: [FOLDER] });
  assert.ok(hits.length > 0, '一致フォルダなのに 0 件はおかしい');
  assert.ok(hits.every(h => h.kind === 'doc'));
});

test('doc フォルダ絞り込み: 不一致フォルダは 0 件 (絞り込みは機能する)', async () => {
  setupDb();
  const hits = await searchVectors('マニュアル', S, 'site', 10, {
    kinds: ['doc'], docFolderPrefixes: ['/sites/n365/Shared Documents/別フォルダ'],
  });
  assert.equal(hits.length, 0, '不一致フォルダは弾くべき');
});

test('kinds 未指定なら doc/pptx 両方が対象になりうる', async () => {
  setupDb();
  const hits = await searchVectors('マニュアル', S, 'site', 20, {});
  const kinds = new Set(hits.map(h => h.kind));
  assert.ok(kinds.has('doc'), 'doc が含まれるべき');
});

test('次元不一致のレコードはスコア0扱いで上位に来ない', async () => {
  const db = new VectorDb();
  const recs = [];
  for (let i = 0; i < 5; i++) recs.push(makeRecord({ i, kind: 'doc', folder: FOLDER }));      // 1024 次元
  recs.push(makeRecord({ i: 99, kind: 'doc', folder: FOLDER, dim: 256 }));                     // 異次元
  db.applySegment(makeSegment('seg-00000', recs));
  __setDb(db);
  __setQuery(normalize(vec(99 + 1000, DIM)));   // 異次元レコードに対応する seed だが次元は 1024
  const hits = await searchVectors('x', S, 'site', 10, { kinds: ['doc'] });
  // 異次元 (f99) はベクトル比較不能なので上位(スコア>0)に来ないこと
  assert.ok(hits.length > 0);
  assert.notEqual(hits[0].docFile, 'f99.pdf');
});
