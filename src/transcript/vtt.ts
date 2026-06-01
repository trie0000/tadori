// WebVTT (.vtt) パーサ。Teams 会議の文字起こしファイルを cue 配列へ。
//
// Teams の .vtt 例:
//   WEBVTT
//
//   00:00:01.230 --> 00:00:05.670
//   <v 山田太郎>では本日の議題ですが、申請フローの見直しについて
//
// 仕様参照: docs/teams-transcript-design.md §4
//
// 注: 会議名・日時は .vtt 内に無い (ファイル名から取る)。ここでは発言だけを扱う。

export interface VttCue {
  /** 開始秒 (録画ジャンプ用)。 */
  startSec: number;
  /** 終了秒。 */
  endSec: number;
  /** 話者の表示名 (`<v 名前>` voice タグ)。無ければ ''。 */
  speaker: string;
  /** 装飾タグ除去済みの発言テキスト。 */
  text: string;
}

/** "00:01:05.230" or "01:05.230" → 秒 (float)。解釈不能なら 0。 */
export function vttTimeToSec(s: string): number {
  const m = s.trim().match(/^(?:(\d+):)?(\d{1,2}):(\d{2})(?:[.,](\d{1,3}))?$/);
  if (!m) return 0;
  const h = m[1] ? Number(m[1]) : 0;
  const min = Number(m[2]);
  const sec = Number(m[3]);
  const ms = m[4] ? Number(m[4].padEnd(3, '0')) : 0;
  return h * 3600 + min * 60 + sec + ms / 1000;
}

/** cue テキストから装飾タグを除去しつつ話者名を取り出す。
 *  `<v 山田太郎>本文` / `<v.loud Bob>...` / `<c>...</c>` / `&amp;` 等に対応。 */
function extractSpeakerAndText(raw: string): { speaker: string; text: string } {
  let speaker = '';
  let body = raw;

  // voice タグ <v Name> or <v.class Name> ... (終了タグ </v> は任意)
  const vm = body.match(/<v(?:\.[^\s>]+)?\s+([^>]+)>/i);
  if (vm) {
    speaker = vm[1].trim();
  }
  // すべての山括弧タグを除去 (<v ...>, </v>, <c>, <i>, <00:00:01.000> 等)
  body = body.replace(/<[^>]*>/g, '');
  // HTML エンティティのデコード (主要なものだけ)
  body = body
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ');
  return { speaker, text: body.trim() };
}

/** .vtt テキスト全体を cue 配列へパース。空や不正は空配列。 */
export function parseVtt(input: string): VttCue[] {
  if (!input) return [];
  // BOM 除去 + 改行正規化
  let text = input.replace(/^﻿/, '').replace(/\r\n?/g, '\n');
  // "WEBVTT" ヘッダ行は無くてもパースは続行 (寛容に)
  const blocks = text.split(/\n{2,}/);
  const cues: VttCue[] = [];

  for (const block of blocks) {
    const lines = block.split('\n').map(l => l).filter(l => l.length > 0);
    if (lines.length === 0) continue;
    // タイムスタンプ行を探す (cue 識別子行が先頭に来ることがあるため)
    let tsIdx = -1;
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].includes('-->')) { tsIdx = i; break; }
    }
    if (tsIdx < 0) continue; // WEBVTT ヘッダ / NOTE / STYLE 等はスキップ

    const tsLine = lines[tsIdx];
    const tm = tsLine.match(/([\d:.,]+)\s*-->\s*([\d:.,]+)/);
    if (!tm) continue;
    const startSec = vttTimeToSec(tm[1]);
    const endSec = vttTimeToSec(tm[2]);

    // タイムスタンプ行の後ろが発言テキスト (複数行ありうる)
    const textLines = lines.slice(tsIdx + 1);
    if (textLines.length === 0) continue;
    const rawText = textLines.join('\n');
    const { speaker, text } = extractSpeakerAndText(rawText);
    if (!text) continue;
    cues.push({ startSec, endSec, speaker, text });
  }
  return cues;
}

/** cue 配列を「[mm:ss 話者] 発言」形式の 1 本のテキストに整形。
 *  チャンク分割 (splitIntoChunks) に渡す前段。話者が連続して同じなら省略しない
 *  (各発言の頭に必ず話者+時刻を付けて、チャンク境界でも話者が分かるように)。 */
export function cuesToText(cues: VttCue[]): string {
  return cues.map(c => {
    const t = secToMmSs(c.startSec);
    const who = c.speaker || '不明';
    return `[${t} ${who}] ${c.text}`;
  }).join('\n');
}

/** 秒 → "h:mm:ss" or "mm:ss"。 */
export function secToMmSs(sec: number): string {
  const s = Math.max(0, Math.floor(sec));
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const ss = s % 60;
  const pad = (n: number): string => String(n).padStart(2, '0');
  return h > 0 ? `${h}:${pad(m)}:${pad(ss)}` : `${m}:${pad(ss)}`;
}
