/* 照合ツールの検証スクリプト
 *
 *   node verify.js                                  … データ非依存のユニットテストのみ
 *   node verify.js <勤務表.xlsx> <勤怠.csv> [<タイミー明細.csv>] [<シート名>]
 *                                                   … 実ファイルを流して結果を表示
 *
 * index.html の <script> からロジック部分だけを取り出して評価する。
 * 画面まわりは document が未定義のときスキップされるため、そのまま読み込める。
 * 実ファイル検証には xlsx パッケージが必要:  npm install xlsx
 */
'use strict';
const fs = require('fs');
const vm = require('vm');
const path = require('path');

function loadLogic() {
  const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
  // ロジック本体は <script id="tool-logic"> に入っている。
  // CDNフォールバック用の小さな <script> と取り違えないよう id で特定する。
  const m = html.match(/<script id="tool-logic">([\s\S]*?)<\/script>/);
  if (!m) throw new Error('tool-logic スクリプトが見つかりません');
  const code = m[1];
  const sandbox = { module: { exports: {} }, console };
  sandbox.exports = sandbox.module.exports;
  vm.createContext(sandbox);
  vm.runInContext(code, sandbox, { filename: 'index.html' });
  if (!sandbox.module.exports.parseSchedule) throw new Error('ロジックを読み込めませんでした');
  return sandbox.module.exports;
}

const L = loadLogic();
let pass = 0, fail = 0;
const ok = (cond, label, detail) => {
  cond ? pass++ : fail++;
  console.log(`  ${cond ? 'PASS' : 'FAIL'}  ${label}${detail !== undefined ? `  → ${detail}` : ''}`);
};

/* ---------------------------------------------------------------- ユニット */
console.log('■ ユニットテスト（データ非依存）');

// B-2 異体字
ok(L.baseNormName('岩﨑　美紀') === L.baseNormName('岩崎 美紀'), 'B-2 異体字 﨑/崎 を同一視', L.baseNormName('岩﨑　美紀'));
ok(L.baseNormName('髙橋明子') === L.baseNormName('高橋明子'), 'B-2 異体字 髙/高 を同一視');
ok(L.baseNormName('渡邊一郎') === L.baseNormName('渡辺一郎'), 'B-2 異体字 邊/辺 を同一視');

// B-6 末尾の数字は残す。括弧付きの連番だけ落とす。
ok(L.baseNormName('タイミー日①') !== L.baseNormName('タイミー日③'), 'B-6 末尾の丸数字で別枠を保持',
  `${L.baseNormName('タイミー日①')} / ${L.baseNormName('タイミー日③')}`);
ok(L.baseNormName('職員01') !== L.baseNormName('職員02'), 'B-6 末尾の連番で別人を保持');
ok(L.baseNormName('山田太郎(2)') === L.baseNormName('山田太郎'), 'B-6 括弧付き連番は同一視');

// B-8 あいまい一致
const dummy = { '早番': { start: 540, end: 1080 }, '早番B': { start: 540, end: 1080 } };
ok(L.getMaster(dummy, '早') === null, 'B-8 マスタ名がシフト側より長い一致は不採用', 'getMaster("早") → null');
ok(L.getMaster(dummy, '早番') === dummy['早番'], 'B-8 完全一致は従来どおり');

// B-9 月の日数
ok(L.daysInMonth(2025, 2) === 28 && L.daysInMonth(2024, 2) === 29 && L.daysInMonth(2026, 7) === 31,
  'B-9 月の日数', `2025/2=${L.daysInMonth(2025, 2)} 2024/2=${L.daysInMonth(2024, 2)} 2026/7=${L.daysInMonth(2026, 7)}`);

// B-10 エイリアスの推移性
const am = L.buildAliasMap('A=B\nB=C');
ok(am[L.baseNormName('A')] === am[L.baseNormName('C')], 'B-10 A=B と B=C を1グループに統合', JSON.stringify(am));

// B-3 終了0:00の夜勤
const yakin = { start: 22 * 60, end: 1440, breakStart: null, breakEnd: null, hours: 0 };
ok(L.classifyShift(yakin, '深夜') === 'nightIn', 'B-3 22:00〜0:00 を夜勤入り判定', L.classifyShift(yakin, '深夜'));
ok(L.expectedHours(yakin) === 2, 'B-3 予定時間は 2.0h のまま', L.expectedHours(yakin) + 'h');

// B-1 休憩表記
ok(L.breakHours('休憩1h') === 1 && L.breakHours('休憩深夜2h') === 2 && L.breakHours('休憩1.5時間') === 1.5,
  'B-1 休憩「1h」「深夜2h」「1.5時間」を控除', `${L.breakHours('休憩1h')} / ${L.breakHours('休憩深夜2h')} / ${L.breakHours('休憩1.5時間')}`);
ok(L.commentOutMin('8/1休憩深夜1h、退勤5:01') === 301, 'B-1 コメントの退勤時刻');

// C-1 必須ヘッダ
let msg = '';
try { L.requireHeaders([['氏名', '日付']], ['従業員氏名', '年/月/日'], 'テストCSV'); } catch (e) { msg = e.message; }
ok(msg.includes('従業員氏名'), 'C-1 必須ヘッダ欠落をエラーにする', msg.slice(0, 60) + '…');

// C-2 日付
ok(L.dateKey('2026/7/1') === '2026/7/1' && L.dateKey('01/05') === null,
  'C-2 解釈できない日付は null を返す', `"2026/7/1"→${L.dateKey('2026/7/1')} / "01/05"→${L.dateKey('01/05')}`);

// C-7 年月推定
ok(L.inferYearMonth({ name: '2607勤務表.xlsx' }).m === 7, 'C-7 ファイル名 2607 → 2026/7');
ok(L.inferYearMonth({ name: 'report-9913.xlsx' }) === null, 'C-7 妥当でない年月は採用しない');

// C-3補助 施設名
ok(L.inferFacility({ name: '【タイミー】RASIEL宇都宮_明細.csv' }, L.DEFAULT_FACILITIES) === '宇都宮',
  'C-3 【タイミー】を施設名と誤認しない', L.inferFacility({ name: '【タイミー】RASIEL宇都宮_明細.csv' }, L.DEFAULT_FACILITIES));

// A-4 日跨ぎの時刻差
ok(L.timeDiffMin(23 * 60, 10) === 70, 'A-4 23:00 と 0:10 の差は 70分', L.timeDiffMin(23 * 60, 10) + '分');

// F-1 マスタY列は日中勤務時間であって総労働時間ではない
{
  const pa = { start: 18 * 60, end: 1440, breakStart: null, breakEnd: null, hours: 4 };
  const pm = { start: 0, end: 9 * 60, breakStart: 0, breakEnd: 120, hours: 4 };
  const sy = { start: 22 * 60, end: 1440, breakStart: null, breakEnd: null, hours: 0 };
  const sm = { start: 0, end: 5 * 60, breakStart: 120, breakEnd: 180, hours: 0 };
  ok(L.expectedHours(pa) === 6 && L.expectedHours(pm) === 7,
    'F-1 パ夜/パ明けの予定時間は Y列(4h)ではなく実労働(6h/7h)',
    `パ夜=${L.expectedHours(pa)}h パ明け=${L.expectedHours(pm)}h → サイクル ${L.expectedHours(pa) + L.expectedHours(pm)}h`);
  ok(L.expectedHours(sy) + L.expectedHours(sm) === 6,
    'F-1 深夜サイクルは 6.0h のまま（タイミー実働6.0hと一致）',
    `${L.expectedHours(sy)}h + ${L.expectedHours(sm)}h`);
  ok(L.daytimeHours(pa) === 4, 'F-1 日中勤務時間はY列から取れる', L.daytimeHours(pa) + 'h');
  ok(L.expectedHours({ start: null, end: null, hours: 8 }) === 8,
    'F-1 時刻を持たないシフトはY列にフォールバック');
  ok(L.plannedShiftHours('有4', {}) === 4 && L.plannedShiftHours('振10', {}) === 10,
    'F-2 「有4」「振10」は時間数つきの休暇として積み上げる',
    `有4=${L.plannedShiftHours('有4', {})}h 振10=${L.plannedShiftHours('振10', {})}h`);
}

// D-2 部分有給
ok(L.leaveHours('有4') === 4 && L.leaveHours('有') === null, 'D-2 「有4」から休暇時間を取り出す', `有4→${L.leaveHours('有4')} / 有→${L.leaveHours('有')}`);
ok(L.isPartialLeave('有4', 8) && !L.isPartialLeave('有8', 8) && !L.isPartialLeave('有13', 8),
  'D-2 標準所定8hなら 有4 のみ部分有給', `有4=${L.isPartialLeave('有4', 8)} 有8=${L.isPartialLeave('有8', 8)} 有13=${L.isPartialLeave('有13', 8)}`);

// D-3 休み系シフト
ok(L.isOffShift('受診') && L.isOffShift('振8') && L.isOffShift('振10'),
  'D-3 受診・振8・振10 を休み系として扱う');

// D-5 早出・残業と遅刻・早退を区別
ok(L.signedDiffMin(9 * 60, 8 * 60 + 24) === -36, 'D-5 予定9:00に8:24出勤 → -36分（早い）', L.signedDiffMin(9 * 60, 8 * 60 + 24));
ok(L.signedDiffMin(18 * 60, 18 * 60 + 41) === 41, 'D-5 予定18:00に18:41退勤 → +41分（遅い）', L.signedDiffMin(18 * 60, 18 * 60 + 41));
ok(L.signedDiffMin(23 * 60, 10) === 70, 'D-5 0時をまたいでも符号つきで求まる', L.signedDiffMin(23 * 60, 10));

// D-4 1日ずれの集約
{
  const src = [
    { 職員名: 'A', 日付: '2026/7/7', シフト: '深夜', 内容: '夜勤入り・出勤打刻なし', CSV出勤: '', CSV退勤: '', 判定: '要確認' },
    { 職員名: 'A', 日付: '2026/7/6', シフト: '空欄', 内容: '空欄＝休み扱いですが、CSVに打刻があります', CSV出勤: '21:47', CSV退勤: '', 判定: '要確認' },
    { 職員名: 'B', 日付: '2026/7/1', シフト: '日', 内容: '出勤打刻なし', CSV出勤: '', CSV退勤: '', 判定: '要確認' }
  ];
  const out = L.mergeShiftedDays(src);
  ok(out.length === 2 && out.some(r => r.判定 === '日ずれ'), 'D-4 隣り合う「打刻なし」と「空欄+打刻あり」を1件に',
    `${src.length}件 → ${out.length}件 / ${out.find(r => r.判定 === '日ずれ')?.内容 || ''}`);
}

// E-4 未登録シフト・エラーの集約
{
  const agg = L.aggregateIssues(
    [{ 施設名: 'X', 職員名: 'A', 日付: '2026/7/1', シフト: '謎', 内容: 'マスタなし' },
     { 施設名: 'X', 職員名: 'B', 日付: '2026/7/2', シフト: '謎', 内容: 'マスタなし' }],
    [{ 施設名: 'X', 職員名: 'C', 日付: '2026/7/3', セル値: '#エラー', 内容: '数式エラー' }]);
  ok(agg.length === 2 && agg[0].件数 === 2, 'E-4 シフト単位に集約して件数で示す',
    agg.map(a => `${a.種別}:${a.シフト}×${a.件数}`).join(' / '));
}

// G-1 レイアウト自動検出（管理用シート＝2列おき / 入力シート＝1列ずつ）
{
  const SERIAL_2026_07_01 = 46204;
  const mkSheet = (nameCol, roleCol, firstDay, step, extra) => {
    const header = [];
    for (let i = 0; i < 31; i++) header[firstDay + i * step] = SERIAL_2026_07_01 + i;
    const mk = (name, role, val) => {
      const r = []; r[nameCol] = name; r[roleCol] = role;
      for (let i = 0; i < 31; i++) r[firstDay + i * step] = val;
      return r;
    };
    return [[], [], header,
      mk('山田　太郎', '生活支援員', '日'),
      mk('鈴木　花子', '生活支援員', '日'),
      mk('佐藤　次郎', '世話人', '日'),
      ...(extra || [])].concat();
  };
  const a = L.detectScheduleLayout(mkSheet(12, 1, 34, 2), () => {}, { maxDays: 31 });
  ok(a && a.name === 12 && a.role === 1 && a.firstDay === 34 && a.step === 2,
    'G-1 管理用シート型（氏名M列・日付2列おき）を検出', JSON.stringify(a));
  const b = L.detectScheduleLayout(mkSheet(2, 1, 5, 1), () => {}, { maxDays: 31 });
  ok(b && b.name === 2 && b.role === 1 && b.firstDay === 5 && b.step === 1,
    'G-1 入力シート型（氏名C列・日付1列ずつ）を検出', JSON.stringify(b));
  ok(b && b.cu === -1, 'G-3 月間時間の列が無いことを検出できる', 'cu=' + (b && b.cu));
  const ym = L.ymFromLayout(mkSheet(2, 1, 5, 1), b);
  ok(ym && ym.y === 2026 && ym.m === 7, 'G-1 日付見出し行から年月を取る', JSON.stringify(ym));

  // G-2 シート下部のシフト別集計行を職員として扱わない
  const master = { '日': { start: 540, end: 1080, breakStart: 780, breakEnd: 840, hours: 8 } };
  const tally = []; tally[2] = '日'; tally[1] = null;
  for (let i = 0; i < 31; i++) tally[5 + i] = 3;          // 集計行は数値
  const sheet = { rows: mkSheet(2, 1, 5, 1, [tally]), errs: new Set() };
  const lay = L.detectScheduleLayout(sheet.rows, () => {}, { maxDays: 31 });
  const sch = L.parseSchedule(sheet, master,
    { facility: 'X', y: 2026, m: 7, aliasMap: {}, includeDispatch: true, warn: () => {}, layout: lay });
  const names = Object.values(sch.metaByName).map(v => v.職員名);
  ok(names.length === 3 && !names.includes('日'),
    'G-2 シフト名が氏名欄に来る集計行を職員にしない', names.join('、'));
  ok(sch.hasCuColumn === false, 'G-3 月間時間の列が無いシートを判別', 'hasCuColumn=' + sch.hasCuColumn);
}

// I-1 連続夜勤（同じ行に出勤と退勤が両方ある日）
{
  const mk = (d, i, o, c) => ({ 日付: d, inMin: i, outMin: o, 出勤: '', 退勤: '', コメント: c || '' });
  const r = L.actualMonthlyHours({ A: [
    mk('2026/7/20', 17 * 60 + 45, null),
    mk('2026/7/21', 17 * 60 + 44, 9 * 60 + 2, '休憩深夜2時間'),
    mk('2026/7/22', null, 9 * 60 + 4, '休憩深夜2時間')] }, 2026, 7, true);
  ok(Math.abs(r.hours.A - 26.62) < 0.05, 'I-1 明け＋入りが同じ行にある夜勤を2回分として計上',
    r.hours.A + 'h（v2.18は13.3h）');
  const r2 = L.actualMonthlyHours({ A: [mk('2026/7/5', 9 * 60, 18 * 60, '休憩1時間')] }, 2026, 7, true);
  ok(r2.hours.A === 8, 'I-1 通常の日勤は従来どおり', r2.hours.A + 'h');
  const r3 = L.actualMonthlyHours({ A: [mk('2026/7/5', 22 * 60, null)] }, 2026, 7, true);
  ok(r3.issues.A.length === 1, 'I-1 相手のいない打刻を警告として返す', r3.issues.A[0]);
}

// I-2 休暇を勤務予定から分ける
{
  const master = { '日': { start: 540, end: 1080, breakStart: 780, breakEnd: 840, hours: 8 } };
  const sch = { normal: { A: { '2026/7/1': [{ シフト: '日' }], '2026/7/2': [{ シフト: '有10' }] } } };
  const p = L.computePlannedHours(sch, master);
  ok(p.workHoursByName.A === 8 && p.leaveHoursByName.A === 10 && p.shiftHoursByName.A === 18,
    'I-2 予定18h = 勤務8h + 休暇10h に分解', JSON.stringify(p));
}

// H-2 氏名の突合
{
  const sch = { normal: { '旭里香(A)': { '2026/7/1': [{ シフト: '日' }] }, 'THANDARAYE': { '2026/7/1': [{ シフト: '日' }] } },
                metaByName: { '旭里香(A)': { 職員名: '旭　里香（A）' }, 'THANDARAYE': { 職員名: 'THANDAR　AYE' } } };
  const kin = { byName: { '旭里香': [{ 職員名: '旭　里香', 日付: '2026/7/1', inMin: 540, outMin: 1080 }],
                          'THANDERAYE': [{ 職員名: 'THANDER　AYE', 日付: '2026/7/1', inMin: 540, outMin: 1080 }] } };
  const r = L.reconcileStaff(sch, kin, {});
  const a = r.find(x => x.氏名 === '旭　里香（A）'), b = r.find(x => x.氏名 === 'THANDAR　AYE');
  ok(a && a.候補 === '旭　里香', 'H-2 末尾の(A)を落として候補を提示', a && a.根拠);
  ok(b && b.候補 === 'THANDER　AYE', 'H-2 1文字違いを編集距離で提示', b && b.根拠);
  ok(L.stripNameSuffix('（宮崎　一徳明け）') === '宮崎　一徳' && L.stripNameSuffix('松本錠二①') === '松本錠二',
    'H-2 施設ごとに違う枝番の書き方を落とせる');
}

// J-2 マスタの統合
{
  const A = { '①明初正': { start: 0, end: 600 }, 'ⅱ夜正': { start: null, end: null } };
  const B = { '①明初正': { start: null, end: null }, 'ⅱ夜正': { start: 960, end: 1440 } };
  const warns = [];
  const M = L.mergeMasters(A, B, m => warns.push(m), 'テスト');
  ok(M['①明初正'].start === 0 && M['ⅱ夜正'].start === 960,
    'J-2 時刻を持つ定義を優先して統合', `①明初正=${M['①明初正'].start} ⅱ夜正=${M['ⅱ夜正'].start}`);
  const w2 = [];
  L.mergeMasters({ 'X': { start: 540, end: 1080 } }, { 'X': { start: 600, end: 1080 } }, m => w2.push(m));
  ok(w2.length === 1, 'J-2 定義が食い違う場合は警告する', w2[0] && w2[0].slice(0, 50) + '…');
}

// L-1〜L-4 「末」シフト（月末出勤・退勤は翌月）
{
  const suePa = { start: 18 * 60, end: 9 * 60, breakStart: null, breakEnd: null, hours: 0 };   // 18:00〜翌09:00
  const sueUt = { start: 18 * 60, end: 1440, breakStart: null, breakEnd: null, hours: 0 };     // 18:00〜24:00
  ok(L.classifyShift(suePa, '末パ夜①') === 'nightIn' && L.classifyShift(sueUt, '末パ夜') === 'nightIn',
    'L-1 日をまたぐ「末」シフトも夜勤入りと判定', `翌朝まで=${L.classifyShift(suePa, '末パ夜①')} / 24:00まで=${L.classifyShift(sueUt, '末パ夜')}`);
  ok(L.classifyShift({ start: 540, end: 18 * 60, breakStart: 780, breakEnd: 840, hours: 8 }, '日') === 'day',
    'L-1 通常の日勤は day のまま');
  ok(L.isMonthEndShift('末パ夜①') && L.isMonthEndShift('②深末') && !L.isMonthEndShift('パ夜'),
    'L-2 「末」を含むシフトを月末出勤とみなす');
  ok(L.monthEndSpansNextDay(suePa) === true && L.monthEndSpansNextDay(sueUt) === false,
    'L-4 翌朝までの登録か24:00までの登録かを判別', `翌朝まで=${L.monthEndSpansNextDay(suePa)} / 24:00まで=${L.monthEndSpansNextDay(sueUt)}`);

  // コメントの退勤時刻（施設ごとに表記が違う）
  const cases = [['8/1休憩深夜1h、退勤5:01', 301], ['8/1深夜休憩1時間　5：00退勤', 300],
                 ['休憩深夜1時間、退勤5時（8/1の退勤時間）', 300], ['休憩深夜2時間、退勤10時（8/1の退勤時間）', 600],
                 ['休憩1時間', null]];
  ok(cases.every(([c, e]) => L.commentOutMin(c) === e), 'L-3 施設ごとに違う退勤コメントの表記に対応',
    cases.map(([c, e]) => `${L.commentOutMin(c)}`).join(' / '));

  // 月末の計上：翌朝までの登録はコメントの退勤まで、24:00までの登録は24:00で切る
  const mk = (d, i, o, c) => ({ 日付: d, inMin: i, outMin: o, 出勤: '', 退勤: '', コメント: c || '' });
  const byName = { A: [mk('2026/7/31', 17 * 60 + 51, null, '8/1 深夜休憩2時間　9：00退勤')] };
  const full = L.actualMonthlyHours(byName, 2026, 7, true,
    { normal: { A: { '2026/7/31': [{ シフト: '末パ夜①', master: suePa }] } } });
  ok(Math.abs(full.hours.A - 13.15) < 0.05, 'L-4 翌朝までの「末」はコメントの退勤まで計上',
    full.hours.A + 'h  ' + full.notes.A[0]);
  const split = L.actualMonthlyHours(byName, 2026, 7, true,
    { normal: { A: { '2026/7/31': [{ シフト: '末パ夜', master: sueUt }] } } });
  ok(Math.abs(split.hours.A - 6.15) < 0.05, 'L-4 24:00までの「末」は24:00で分割',
    split.hours.A + 'h  ' + split.notes.A[0]);
  ok(full.issues.A.length === 0 && split.issues.A.length === 0,
    'L-2 「末」シフトの退勤打刻なしを警告にしない');
}

console.log(`\n  ${pass} passed / ${fail} failed\n`);

/* ------------------------------------------------------------ 実データ検証 */
const [xlsxPath, kintaiPath, timeePath, sheetArg] = process.argv.slice(2);
const SHEET = sheetArg || '入力シート【現場配布用】';
if (!xlsxPath || !kintaiPath) {
  console.log('実データ検証をするには:  node verify.js <勤務表.xlsx> <勤怠.csv> [<タイミー明細.csv>]');
  process.exit(fail ? 1 : 0);
}

let XLSX;
try { XLSX = require('xlsx'); }
catch (e) { console.error('xlsx パッケージが必要です:  npm install xlsx'); process.exit(1); }

function readSheet(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) throw new Error(`シートがありません: ${name} / 実際: ${wb.SheetNames.join('、')}`);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  const errs = new Set();
  Object.keys(ws).forEach(k => {
    if (k[0] === '!') return;
    const c = ws[k];
    if (c && c.t === 'e') { const a = XLSX.utils.decode_cell(k); errs.add(a.r + ',' + a.c); }
  });
  return { rows, errs };
}
function readTextSmart(file) {
  const buf = fs.readFileSync(file);
  if (buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) return buf.toString('utf8').replace(/^﻿/, '');
  const utf8 = new TextDecoder('utf-8').decode(buf);
  if (!utf8.includes('�')) return utf8;
  return new TextDecoder('shift_jis').decode(buf);
}

const logs = [];
const warn = m => { if (!logs.includes(m)) logs.push(m); };
const aliasMap = L.buildAliasMap('');
const wb = XLSX.readFile(xlsxPath);
const { master } = L.parseMaster(readSheet(wb, 'マスタ設定').rows, warn);
const sheet = readSheet(wb, SHEET);
const layout = L.detectScheduleLayout(sheet.rows, warn, { maxDays: 31 });
const ym = L.inferYearMonth({ name: path.basename(xlsxPath) }) || L.ymFromLayout(sheet.rows, layout) || L.parseYmFromSheet(sheet.rows);
if (!ym) { console.error('年月を判定できませんでした'); process.exit(1); }
const { y, m } = ym;

const sch = L.parseSchedule(sheet, master, { facility: '(単一施設)', y, m, aliasMap, includeDispatch: false, warn, layout });
Object.assign(sch, L.computePlannedHours(sch, master));
const kin = L.parseKintai(readTextSmart(kintaiPath), {
  facility: '(単一施設)', excludedKeys: sch.excludedKeys, aliasMap, warn, fileName: path.basename(kintaiPath)
});
let tim = [];
if (timeePath) {
  tim = L.parseTimee(readTextSmart(timeePath), {
    facility: '(単一施設)', storeMap: {}, knownFacilities: L.DEFAULT_FACILITIES, warn, fileName: path.basename(timeePath)
  });
  const before = tim.length;
  tim = tim.filter(r => { const p = r.求人日.split('/').map(Number); return p[0] === y && p[1] === m; });
  if (before !== tim.length) warn(`タイミー明細のうち ${before - tim.length} 件は ${y}/${m} 以外のため除外`);
}

const slots = L.buildExternalSlots(sch.external, master, y, m);
const TOL = { late: 15, over: 60, standardHours: 8 };
const dailyRaw = L.compareDaily(sch, kin, TOL, y, m, warn);
const daily = L.mergeShiftedDays(dailyRaw);
const monthly = L.compareMonthly(sch, kin, { hour: 0.25, pct: 5 }, y, m, true, warn);
const staffMatch = L.reconcileStaff(sch, kin, aliasMap);
const issues = L.aggregateIssues(sch.unknown, sch.errorCells);
const GROUP = '(単一施設)';
const counts = L.compareTimeeCounts(slots, tim, GROUP);
const details = L.compareTimeeDetails(slots, tim, 15, 0.25, GROUP);

console.log('='.repeat(74));
console.log(`■ 実データ検証  ${y}/${m}   シート「${SHEET}」`);
console.log(`  レイアウト          氏名列 ${layout && layout.name} / 職種列 ${layout && layout.role} / 日付 ${layout && layout.firstDay}列から${layout && layout.step}列ずつ / 月間時間列 ${layout && layout.cu >= 0 ? layout.cu : 'なし'}`);
console.log('='.repeat(74));
console.log(`  日別ズレ            ${daily.length}件（1日ずれ集約前 ${dailyRaw.length}件 / うち日ずれ ${daily.filter(r => r.判定 === '日ずれ').length}件）`);
console.log(`  月間集計            ${monthly.length}名（要確認 ${monthly.filter(r => r.判定 !== 'OK').length}名）`);
console.log(`  氏名の未突合        ${staffMatch.filter(r => r.判定 !== '打刻なし').length}名`);
console.log(`  未登録シフト/エラー ${issues.length}種（延べ ${issues.reduce((a, r) => a + r.件数, 0)}件）`);
console.log(`  タイミー枠          ${slots.length}件（タイミー ${slots.filter(s => s.枠種別 === 'タイミー').length} / 派遣 ${slots.filter(s => s.枠種別 === '派遣').length}）`);
console.log(`  タイミー明細        ${tim.length}件`);

console.log('\n■ 月間集計');
console.log('  職員名         予定  休暇 勤務予定 予定日   実績  打刻日     差分   許容  判定');
monthly.forEach(r => console.log(
  `  ${String(r.職員名).padEnd(13).slice(0,13)}${String(r.予定時間).padStart(5)}${String(r.うち休暇).padStart(6)}${String(r.勤務予定).padStart(8)}${String(r.予定勤務日数).padStart(6)}${String(r.CSV実績概算時間).padStart(8)}${String(r.実績打刻日数).padStart(6)}${String(r.差分).padStart(9)}${String(r.許容).padStart(6)}  ${r.判定}`));
if (staffMatch.length) {
  console.log('\n■ 氏名の突合');
  staffMatch.forEach(r => console.log(`  ${r.区分}  ${String(r.氏名).padEnd(16).slice(0,16)} ${r.判定}${r.候補 ? '  → 候補「' + r.候補 + '」(' + r.根拠 + ')' : ''}`));
}

if (counts.length) {
  console.log('\n■ タイミー枠数照合（レベル1）');
  ['日勤', '夜勤'].forEach(k => {
    const c = counts.filter(r => r.区分 === k);
    if (!c.length) return;
    console.log(`  ${k}: ${c.length}日中 一致 ${c.filter(r => r.判定 === 'OK').length}日 / 不一致 ${c.filter(r => r.判定 !== 'OK').length}日` +
      `   枠計 ${c.reduce((a, r) => a + r.シフト枠数, 0)} vs 明細計 ${c.reduce((a, r) => a + r.明細件数, 0)}`);
  });
  const ng = counts.filter(r => r.判定 !== 'OK');
  if (ng.length) {
    console.log('\n  日付        区分  枠  タイミー 派遣  明細  過不足  派遣除く');
    ng.forEach(r => console.log(
      `  ${r.日付.padEnd(11)}${r.区分}  ${String(r.シフト枠数).padStart(2)}  ${String(r.うちタイミー).padStart(6)} ${String(r.うち派遣).padStart(4)}  ${String(r.明細件数).padStart(4)}  ${r.過不足.padStart(5)}  ${r['過不足(派遣除く)'].padStart(6)}`));
  }
  console.log('\n■ タイミー明細照合（レベル2）判定内訳');
  const tc = {};
  details.forEach(r => { const k = r.判定 + (r.内容 ? ' / ' + r.内容 : ''); tc[k] = (tc[k] || 0) + 1; });
  Object.entries(tc).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => console.log(`  ${String(v).padStart(3)}  ${k}`));
}

console.log(`\n■ 警告 ${logs.length}件`);
logs.forEach(l => console.log('  ・' + l));
process.exit(fail ? 1 : 0);
