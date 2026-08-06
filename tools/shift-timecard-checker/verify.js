/* 照合ツールの検証スクリプト
 *
 *   node verify.js                                  … データ非依存のユニットテストのみ
 *   node verify.js <勤務表.xlsx> <勤怠.csv> [<タイミー明細.csv>]
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
  const code = html.split('<script>')[1].split('</script>')[0];
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

console.log(`\n  ${pass} passed / ${fail} failed\n`);

/* ------------------------------------------------------------ 実データ検証 */
const [xlsxPath, kintaiPath, timeePath] = process.argv.slice(2);
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
const sheet = readSheet(wb, '管理用シート');
const ym = L.inferYearMonth({ name: path.basename(xlsxPath) }) || L.parseYmFromSheet(sheet.rows);
if (!ym) { console.error('年月を判定できませんでした'); process.exit(1); }
const { y, m } = ym;

const sch = L.parseSchedule(sheet, master, { facility: '(単一施設)', y, m, aliasMap, includeDispatch: true, warn });
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
const daily = L.compareDaily(sch, kin, 15, y, m, warn);
const monthly = L.compareMonthly(sch, kin, 0.25, y, m, true);
const GROUP = '(単一施設)';
const counts = L.compareTimeeCounts(slots, tim, GROUP);
const details = L.compareTimeeDetails(slots, tim, 15, 0.25, GROUP);

console.log('='.repeat(74));
console.log(`■ 実データ検証  ${y}/${m}`);
console.log('='.repeat(74));
console.log(`  日別ズレ            ${daily.length}件`);
console.log(`  月間集計            ${monthly.length}名（要確認 ${monthly.filter(r => r.判定 !== 'OK').length}名）`);
console.log(`  未登録シフト        ${sch.unknown.length}件`);
console.log(`  数式エラーセル      ${sch.errorCells.length}件`);
console.log(`  タイミー枠          ${slots.length}件（タイミー ${slots.filter(s => s.枠種別 === 'タイミー').length} / 派遣 ${slots.filter(s => s.枠種別 === '派遣').length}）`);
console.log(`  タイミー明細        ${tim.length}件`);

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
