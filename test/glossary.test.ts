// 用語辞書の純ロジック検証 (クエリ展開 / Excel貼り付けパース)。
import test from 'node:test';
import assert from 'node:assert/strict';
import { expandQueryTerms, parseGlossaryTable, type GlossaryEntry } from '../src/search/glossary';

const G: GlossaryEntry[] = [
  { canonical: 'ワークフロー', aliases: ['P-WF', 'PWF', '業務フロー'] },
  { canonical: '勤怠管理システム', aliases: ['勤怠', 'KMS'] },
];

test('略語で引くと正式名・他表記が展開される', () => {
  const add = expandQueryTerms('P-WF の申請手順', G);
  assert.ok(add.includes('ワークフロー'));
  assert.ok(add.includes('業務フロー'));
  assert.ok(!add.includes('P-WF'), '既に出ている語は足さない');
});

test('正式名で引くと略語が展開される', () => {
  const add = expandQueryTerms('ワークフローとは', G);
  assert.ok(add.includes('P-WF'));
});

test('該当なしクエリは展開ゼロ', () => {
  assert.equal(expandQueryTerms('経費精算の締め日', G).length, 0);
});

test('maxAdds で上限', () => {
  const add = expandQueryTerms('ワークフロー 勤怠', G, 2);
  assert.equal(add.length, 2);
});

test('TSV 貼り付けをパース (見出し行スキップ)', () => {
  const tsv = '正式名\t別名\t意味\nワークフロー\tP-WF, PWF\t申請経路\n勤怠管理システム\t勤怠;KMS\t';
  const e = parseGlossaryTable(tsv);
  assert.equal(e.length, 2);
  assert.deepEqual(e[0].aliases, ['P-WF', 'PWF']);
  assert.equal(e[0].def, '申請経路');
  assert.deepEqual(e[1].aliases, ['勤怠', 'KMS']);
});

test('CSV もパースできる', () => {
  const csv = 'ワークフロー,P-WF,';
  const e = parseGlossaryTable(csv);
  assert.equal(e[0].canonical, 'ワークフロー');
  assert.deepEqual(e[0].aliases, ['P-WF']);
});
