// buildScope: フォルダ群(doc/pptx/transcript)の kinds 補完を検証。
// 「doc だけ active でも pptx/transcript が検索対象に含まれる」回帰を固定。
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildScope, type SubSelection } from '../src/search/scopeSelection';

const SITE = 'https://x.sharepoint.com/sites/scope-test';
const EMPTY: SubSelection = { mail: [], onenote: [], folders: [] };

test('doc だけ active でも pptx/transcript が kinds に補完される', () => {
  const { kinds } = buildScope(SITE, ['doc'], EMPTY);
  assert.ok(kinds.includes('doc') && kinds.includes('pptx') && kinds.includes('transcript'));
});

test('pptx だけ active でも doc/transcript が補完される', () => {
  const { kinds } = buildScope(SITE, ['pptx'], EMPTY);
  assert.ok(kinds.includes('doc') && kinds.includes('pptx') && kinds.includes('transcript'));
});

test('フォルダ群が無効 (mail のみ) なら補完しない', () => {
  const { kinds } = buildScope(SITE, ['mail'], EMPTY);
  assert.deepEqual(kinds, ['mail']);
});

test('フォルダ選択時は scope.folders を serverRelative で設定 + 3種別補完', () => {
  const { kinds, scope } = buildScope(SITE, ['mail'], { ...EMPTY, folders: ['https://x.sharepoint.com/sites/foo/Shared Documents/資料'] });
  assert.ok(scope.folders && scope.folders[0].startsWith('/sites/foo/'));
  assert.ok(kinds.includes('pptx'));
});
