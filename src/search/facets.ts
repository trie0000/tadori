// 検索ヒット (出典) の内訳を集計する。チャットの「根拠の内訳」チップ表示＋絞り込み用。
// 種別 / 時期(YYYY-MM) / 差出人 で件数を出す。純関数なのでテスト可能。

export interface FacetHit {
  kind?: string;
  date?: string;   // ISO
  from?: string;   // 差出人 (mail) / OneNote はノート›セクション 等
}

export interface FacetBucket { key: string; label: string; count: number; }
export interface Facets {
  kind: FacetBucket[];
  month: FacetBucket[];
  from: FacetBucket[];
}

const KIND_LABEL: Record<string, string> = {
  mail: 'メール', onenote: 'OneNote', pptx: 'PPTX', transcript: '会議', doc: '文書',
};

function tally(items: { key: string; label: string }[]): FacetBucket[] {
  const m = new Map<string, FacetBucket>();
  for (const it of items) {
    const b = m.get(it.key);
    if (b) b.count++;
    else m.set(it.key, { key: it.key, label: it.label, count: 1 });
  }
  return [...m.values()].sort((a, b) => b.count - a.count);
}

/** ヒット配列から種別/時期/差出人のファセットを集計。各 top で件数上位に絞る。 */
export function computeFacets(hits: FacetHit[], top = 8): Facets {
  const kind = tally(hits.map(h => {
    const k = h.kind ?? 'mail';
    return { key: k, label: KIND_LABEL[k] ?? k };
  }));
  const month = tally(
    hits.filter(h => h.date && /^\d{4}-\d{2}/.test(h.date))
      .map(h => { const ym = h.date!.slice(0, 7); return { key: ym, label: ym }; }),
  ).slice(0, top);
  const from = tally(
    hits.filter(h => (h.from ?? '').trim())
      .map(h => { const f = h.from!.trim(); return { key: f, label: f }; }),
  ).slice(0, top);
  return { kind, month, from };
}
