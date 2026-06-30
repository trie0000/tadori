// 診断「取り込み済み一覧」の集計 (ingestedSummary) を検証。
import test from 'node:test';
import assert from 'node:assert/strict';
import { VectorDb } from '../src/db/store';
import { makeRecord, makeSegment } from './_fixtures';

test('メール: 件数・ML(to/cc)・期間を集計', () => {
  const db = new VectorDb();
  const recs = [
    { ...makeRecord({ i: 0, kind: 'mail', to: ['ml-a@example.com'], cc: ['x@example.com'] }), date: '2026-03-01T00:00:00Z' },
    { ...makeRecord({ i: 1, kind: 'mail', to: ['ml-b@example.com'], cc: [] }), date: '2026-05-20T00:00:00Z' },
  ];
  db.applySegment(makeSegment('seg-00000', recs));
  const s = db.ingestedSummary();
  assert.equal(s.mail.count, 2);
  assert.equal(s.mail.dateMin, '2026-03-01T00:00:00Z');
  assert.equal(s.mail.dateMax, '2026-05-20T00:00:00Z');
  assert.ok(s.mail.mls.includes('ml-a@example.com') && s.mail.mls.includes('ml-b@example.com'));
});

test('文書/PPTX: ファイル単位で集約 (チャンク数・場所・名前・取り込み時間)', () => {
  const F = '/sites/n365/Shared Documents/マニュアル';
  const db = new VectorDb();
  // 同一 doc ファイルの2チャンク + pptx 1件
  const r0 = { ...makeRecord({ i: 0, kind: 'doc', folder: F }), chunkIdx: 0, ingestedAt: '2026-06-30T01:00:00Z' };
  const r1 = { ...makeRecord({ i: 0, kind: 'doc', folder: F }), messageId: 'doc://x#1', chunkIdx: 1, ingestedAt: '2026-06-30T02:00:00Z' };
  const p0 = { ...makeRecord({ i: 0, kind: 'pptx', folder: F }), ingestedAt: '2026-06-30T03:00:00Z' };
  db.applySegment(makeSegment('seg-00000', [r0, r1, p0]));
  const s = db.ingestedSummary();
  const doc = s.docs.find(d => d.kind === 'doc')!;
  assert.equal(doc.chunks, 2, '同一ファイルの2チャンクは1文書に集約');
  assert.equal(doc.title, 'f0.pdf');
  assert.equal(doc.location, F);
  assert.equal(doc.ingestedAt, '2026-06-30T02:00:00Z', '取り込み時間は最大(最新)');
  assert.ok(s.docs.some(d => d.kind === 'pptx'));
});
