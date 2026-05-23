// RAG 検索の対象から除外するメールを定義するルール群。
// Outlook の振り分けルール的に、件名/送信者/To/Cc/本文 のいずれかに指定文字列を
// 含むメールを検索結果から除外する。ルールは localStorage に保存。

import { htmlToText } from '../lib/mailhtml';

export type ExclusionField = 'subject' | 'from' | 'to' | 'cc' | 'body';

export interface ExclusionRule {
  id: string;
  field: ExclusionField;
  /** 含まれていれば除外 (case-insensitive substring)。 */
  value: string;
  /** 既定 true。ON/OFF できる。 */
  enabled?: boolean;
}

/** ルール判定で参照するメールの最小フィールド (MailHit / MailRecord いずれも適合)。 */
export interface RuleTarget {
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  body: string;
  isHtml: boolean;
}

const KEY = 'tadori:exclusion-rules';

function lsRead(): ExclusionRule[] {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return [];
    const arr = JSON.parse(raw) as ExclusionRule[];
    return Array.isArray(arr) ? arr : [];
  } catch { return []; }
}
function lsWrite(rules: ExclusionRule[]): void {
  try { localStorage.setItem(KEY, JSON.stringify(rules)); } catch { /* quota */ }
}

export function loadRules(): ExclusionRule[] { return lsRead(); }
export function saveRules(rules: ExclusionRule[]): void { lsWrite(rules); }

export function newRuleId(): string {
  return 'r-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6);
}

export function addRule(rule: Omit<ExclusionRule, 'id'>): ExclusionRule {
  const list = lsRead();
  const r: ExclusionRule = { id: newRuleId(), enabled: true, ...rule };
  list.push(r);
  lsWrite(list);
  return r;
}

export function deleteRule(id: string): void {
  lsWrite(lsRead().filter(r => r.id !== id));
}

export function updateRule(id: string, patch: Partial<ExclusionRule>): void {
  const list = lsRead();
  const r = list.find(x => x.id === id);
  if (!r) return;
  Object.assign(r, patch);
  lsWrite(list);
}

/** ルールのいずれか 1 つにでも一致したら除外対象 (true) を返す。 */
export function matchesAnyRule(target: RuleTarget, rules: ExclusionRule[]): boolean {
  const active = rules.filter(r => r.enabled !== false && r.value);
  if (active.length === 0) return false;
  for (const r of active) {
    if (matchOne(target, r)) return true;
  }
  return false;
}

function matchOne(t: RuleTarget, r: ExclusionRule): boolean {
  const v = r.value.toLowerCase();
  switch (r.field) {
    case 'subject': return t.subject.toLowerCase().includes(v);
    case 'from':    return t.from.toLowerCase().includes(v);
    case 'to':      return t.to.some(a => a.toLowerCase().includes(v));
    case 'cc':      return t.cc.some(a => a.toLowerCase().includes(v));
    case 'body': {
      const body = t.isHtml ? htmlToText(t.body) : t.body;
      return body.toLowerCase().includes(v);
    }
  }
}

export const FIELD_LABELS: Record<ExclusionField, string> = {
  subject: '件名',
  from: '送信者',
  to: 'To',
  cc: 'Cc',
  body: '本文',
};
