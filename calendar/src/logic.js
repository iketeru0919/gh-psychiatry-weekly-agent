const WEEKDAY_LABELS = ['日', '月', '火', '水', '木', '金', '土'];
const NOW = new Date();

// v1 は「2026年8月固定」の想定でキーに月が埋め込まれていた。v2 では月に依存しない。
// 旧キーは読み込み時にだけ参照し、移行後は新キーへ書き出す。
const STORAGE_KEY = 'ss-kanri-v2';
const LEGACY_STORAGE_KEY = 'ss-kanri-august-2026-v1';
const AUTO_BACKUP_KEY = 'ss-kanri-auto-backup-v2';
const LEGACY_AUTO_BACKUP_KEY = 'ss-kanri-august-2026-auto-backup-v1';

// タイムラインの1日あたりの横幅(px)。ドラッグ量から日数を割り出すため、
// 可変幅ではなく固定幅にしている。
const DAY_W = 38;

function pad2(n) { return String(n).padStart(2, '0'); }
function fmtDate(d) { return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }

// new Date('2026-07-01') parses as UTC midnight, which reads back as the PREVIOUS
// day through getDate() anywhere west of UTC. Date strings are always parsed as
// local dates so the calendar never shifts by a day.
function parseDate(str) {
  const p = String(str).split('-').map(Number);
  return new Date(p[0], (p[1] || 1) - 1, p[2] || 1);
}
function shiftDate(dateStr, n) { const d = parseDate(dateStr); d.setDate(d.getDate() + n); return fmtDate(d); }
function monthKeyOf(dateStr) { return String(dateStr).slice(0, 7); }
function daysInMonthOf(monthKey) { const p = monthKey.split('-').map(Number); return new Date(p[0], p[1], 0).getDate(); }
function monthStartOf(monthKey) { return monthKey + '-01'; }
function monthEndOf(monthKey) { return monthKey + '-' + pad2(daysInMonthOf(monthKey)); }
function shiftMonth(monthKey, n) {
  const p = monthKey.split('-').map(Number);
  const d = new Date(p[0], p[1] - 1 + n, 1);
  return fmtDate(d).slice(0, 7);
}
function monthLabelOf(monthKey) {
  const p = monthKey.split('-').map(Number);
  return p[0] + '年' + p[1] + '月';
}
function dateLabelOf(dateStr) {
  const d = parseDate(dateStr);
  return (d.getMonth() + 1) + '月' + d.getDate() + '日（' + WEEKDAY_LABELS[d.getDay()] + '）';
}
function daysBetween(a, b) {
  return Math.round((parseDate(b).getTime() - parseDate(a).getTime()) / 86400000);
}

const TODAY = fmtDate(NOW);
const THIS_MONTH = monthKeyOf(TODAY);

// 退所未定の滞在は終わりが無い。全期間を数えようとすると際限なく日付を歩くので、
// 「いつまでを未定の滞在とみなすか」の上限をここで決めておく。
const RANGE_MIN = '1900-01-01';
const RANGE_MAX = (() => { const d = new Date(NOW); d.setDate(d.getDate() + 400); return fmtDate(d); })();

// ISO date strings compare chronologically, so ranges can be walked as strings.
function eachDate(startStr, endStr, fn) {
  if (!startStr || !endStr || endStr < startStr) return;
  let cur = startStr;
  for (let guard = 0; cur <= endStr && guard < 4000; guard++) { fn(cur); cur = shiftDate(cur, 1); }
}

const STATUS_LIST = ['確定', '仮予約', '調整中', '利用済み', 'キャンセル'];
// 稼働率・請求の母数に入れてよいのは確定系のみ。仮予約・調整中は「見込み」。
const TENTATIVE_STATUSES = ['仮予約', '調整中'];
function isCancelled(status) { return (status || '確定') === 'キャンセル'; }
function isTentative(status) { return TENTATIVE_STATUSES.indexOf(status || '確定') >= 0; }

function uid(prefix) { return prefix + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

function defaultMaster() {
  return {
    facilityName: 'RASIEL萩丘ⅰ',
    buildingLabel: 'ⅰ棟 1F',
    capacity: 2,
    rooms: [{ id: 'r1', name: '部屋1' }, { id: 'r2', name: '部屋2' }],
    users: [
      { id: 'u1', name: '蓬町様（漢字要確認）' },
      { id: 'u2', name: '麻美様' },
      { id: 'u3', name: '人見様' },
    ],
    categories: [
      { id: 'c1', label: 'SS', color: '#3a6ea5' },
      { id: 'c2', label: '無料', color: '#3f8a5c' },
      { id: 'c3', label: '契約', color: '#7a4f8c' },
    ],
    pickupOptions: [
      { id: 'p1', label: 'あり', highlight: true },
      { id: 'p2', label: 'なし', highlight: false },
      { id: 'p3', label: '調整中', highlight: true },
    ],
    defaultCheckinTime: '16:00',
    defaultCheckoutTime: '09:00',
    operator: '',
    // ── FAX帳票用 ──────────────────────────────────────────────
    // 送信元として帳票の頭に刷り込む施設の連絡先。
    postalCode: '',
    address: '',
    tel: '',
    fax: '',
    contactPerson: '',
    // 送付先（相談支援事業所・給食業者・家族など）。誤送信を避けるため、
    // 番号は都度手入力せずここから選ぶ運用にする。
    faxRecipients: [],
    // 予約確認書に差し込む定型文。施設ごとに違うので編集できるようにする。
    boilerplate: {
      belongings: '受給者証・保険証、内服薬（お薬手帳）、着替え、洗面用具、タオル、上履き',
      notices: '・お薬は1回分ずつ分包し、お名前をご記入ください。\n・貴重品はお預かりできません。\n・体調に変化がある場合は、事前にご連絡ください。',
    },
  };
}

// ── 食事 ────────────────────────────────────────────────────────────────
// 食事は「1滞在に朝昼夕の3フラグ」ではなく日別に持つ。連泊では初日は夕食から、
// 最終日は朝食まで、中日は3食、という実際の運用がそのまま表現できる。
function defaultMealsFor(dateStr, startDate, endDate) {
  if (startDate === endDate) return { b: false, l: true, d: false }; // 日帰り(体験利用)
  if (dateStr === startDate) return { b: false, l: false, d: true };
  if (dateStr === endDate) return { b: true, l: false, d: false };
  return { b: true, l: true, d: true };
}
function buildMealsByDate(startDate, endDate, existing, legacyFlat) {
  const out = {};
  // 旧データ(滞在に1組だけの朝昼夕)は、1つでもONなら全日に展開する。
  // 全部OFFのものは「未入力」とみなし標準パターンで初期化する。
  const useLegacy = !!legacyFlat && (legacyFlat.b || legacyFlat.l || legacyFlat.d);
  eachDate(startDate, endDate, (d) => {
    if (existing && existing[d]) {
      out[d] = { b: !!existing[d].b, l: !!existing[d].l, d: !!existing[d].d };
    } else if (useLegacy) {
      out[d] = { b: !!legacyFlat.b, l: !!legacyFlat.l, d: !!legacyFlat.d };
    } else {
      out[d] = defaultMealsFor(d, startDate, endDate);
    }
  });
  return out;
}
// 手で触った日だけを残す。日程を伸縮したとき、以前は中日だった日の設定が
// そのまま最終日に残ってしまい「最終日なのに3食」になっていた。
function pickTouched(table, touched) {
  if (!table || !touched) return null;
  const out = {};
  Object.keys(touched).forEach((d) => { if (touched[d] && table[d]) out[d] = table[d]; });
  return out;
}
function allTouched(table) {
  const out = {};
  Object.keys(table || {}).forEach((d) => { out[d] = true; });
  return out;
}

function mealsOfStay(stay) {
  const end = stay.end == null ? stay.start : stay.end;
  return buildMealsByDate(stay.start, end, stay.mealsByDate, stay.meals);
}

// Anchored to the month the app is opened in, so the demo rows never fall off
// the end of the calendar as time passes.
function sampleData() {
  const makeStay = (stayId, name, userId, startDate, startTime, endDate, endTime, notes) => {
    const shared = {
      stayId, name, userId, category: 'SS',
      meals: { b: false, l: false, d: false },
      mealsByDate: buildMealsByDate(startDate, endDate, null, null),
      pickup: '調整中', notes: notes || '', room: 'r1', status: '確定',
      pickupMethod: '', pickupLocation: '', externalService: '',
      updatedAt: '', updatedBy: '',
    };
    return {
      startDate, endDate,
      checkin: { ...shared, id: stayId + '-in', type: '入所', time: startTime },
      checkout: { ...shared, id: stayId + '-out', type: '退所', time: endTime },
    };
  };
  const stays = [
    makeStay('aug26-yomogi-01', '蓬町様（漢字要確認）', 'u1', '2026-08-01', '10:00', '2026-08-02', '16:00', '氏名の漢字は要確認。食事・送迎は未確認。'),
    // 元の管理表どおり、この滞在は入退所時刻が記録されていない（空のまま残す）。
    makeStay('aug26-asami-01', '麻美様', 'u2', '2026-08-05', '', '2026-08-18', '', '入所・退所時刻は記載なし。確定後に入力してください。食事・送迎は未確認。'),
    makeStay('aug26-yomogi-02', '蓬町様（漢字要確認）', 'u1', '2026-08-22', '10:00', '2026-08-23', '16:00', '氏名の漢字は要確認。食事・送迎は未確認。'),
    makeStay('aug26-hitomi-01', '人見様', 'u3', '2026-08-28', '15:40', '2026-08-30', '13:00', '食事・送迎は未確認。'),
  ];
  const out = {};
  stays.forEach((s) => {
    const mkIn = monthKeyOf(s.startDate);
    const mkOut = monthKeyOf(s.endDate);
    out[mkIn] = out[mkIn] || {};
    out[mkOut] = out[mkOut] || {};
    (out[mkIn][s.startDate] = out[mkIn][s.startDate] || []).push(s.checkin);
    (out[mkOut][s.endDate] = out[mkOut][s.endDate] || []).push(s.checkout);
  });
  return out;
}

// Rewrite every stored record. Master edits cascade through this so a rename or
// a deletion can never leave records pointing at something that no longer exists.
function mapRecords(data, fn) {
  const out = {};
  Object.keys(data || {}).forEach((mk) => {
    const month = {};
    Object.keys(data[mk] || {}).forEach((dateStr) => {
      month[dateStr] = (data[mk][dateStr] || []).map(fn);
    });
    out[mk] = month;
  });
  return out;
}

function applyRecord(data, dateStr, record, isNew) {
  // The bucket comes from the record's own date, never from the month on screen,
  // so a day borrowed from the next month lands in the right place.
  const mk = monthKeyOf(dateStr);
  const out = { ...data };
  const month = { ...(out[mk] || {}) };
  const list = [...(month[dateStr] || [])];
  const idx = isNew ? -1 : list.findIndex((x) => x.id === record.id);
  if (idx >= 0) list[idx] = record; else list.push(record);
  month[dateStr] = list;
  out[mk] = month;
  return out;
}

function removeRecords(data, predicate) {
  const out = {};
  Object.keys(data || {}).forEach((mk) => {
    const month = {};
    Object.keys(data[mk] || {}).forEach((dateStr) => {
      month[dateStr] = (data[mk][dateStr] || []).filter((r) => !predicate(r, dateStr));
    });
    out[mk] = month;
  });
  return out;
}

function recordEntries(data) {
  const out = [];
  Object.keys(data || {}).forEach((mk) => Object.keys(data[mk] || {}).forEach((dateStr) => {
    (data[mk][dateStr] || []).forEach((rec) => out.push({ dateStr, rec }));
  }));
  return out;
}

// ── 滞在の導出 ──────────────────────────────────────────────────────────
// stayId があるレコードは stayId でペアリングする。氏名+部屋でペアリングしていた
// 旧実装は、滞在中の部屋替えや氏名の修正で入所と退所が分離して滞在が壊れた。
// stayId を持たない旧形式レコードだけ、従来どおり氏名+部屋で連結する。
function deriveStays(data, roomIds, fallbackRoom) {
  const roomOf = (r) => (r.room && roomIds.has(r.room) ? r.room : fallbackRoom);
  const entries = recordEntries(data);
  const stays = [];
  const paired = new Set();

  const baseOf = (rec, room) => ({
    name: rec.name,
    userId: rec.userId || null,
    room,
    category: rec.category,
    status: rec.status || '確定',
    notes: rec.notes || '',
    pickup: rec.pickup,
    pickupMethod: rec.pickupMethod || '',
    pickupLocation: rec.pickupLocation || '',
    externalService: rec.externalService || '',
    meals: rec.meals || null,
    mealsByDate: rec.mealsByDate || null,
    updatedAt: rec.updatedAt || '',
    updatedBy: rec.updatedBy || '',
  });

  const byStayId = {};
  entries.forEach((e) => {
    if (e.rec.stayId) (byStayId[e.rec.stayId] = byStayId[e.rec.stayId] || []).push(e);
  });
  Object.keys(byStayId).forEach((sid) => {
    const group = byStayId[sid];
    const ci = group.find((e) => e.rec.type === '入所');
    const co = group.find((e) => e.rec.type === '退所');
    if (!ci) return; // 入所が無いものは旧形式扱いへ回す
    group.forEach((e) => paired.add(e.rec.id));
    stays.push({
      ...baseOf(ci.rec, roomOf(ci.rec)),
      stayId: sid,
      legacy: false,
      checkinId: ci.rec.id,
      checkoutId: co ? co.rec.id : null,
      start: ci.dateStr,
      // 重複判定のために補完した時刻と、実際に記録されている生の時刻は分けて持つ。
      // 補完値をそのまま編集画面へ出すと、「時刻未記入」が 00:00 として確定保存されてしまう。
      startTime: String(ci.rec.time || '00:00').padStart(5, '0'),
      startTimeRaw: ci.rec.time || '',
      end: co ? co.dateStr : null,
      endTime: co ? String(co.rec.time || '23:59').padStart(5, '0') : null,
      endTimeRaw: co ? (co.rec.time || '') : '',
    });
  });

  // 旧形式（stayId 無し）: 氏名+部屋で入所→退所を連結する。
  const legacyEvents = entries
    .filter((e) => !paired.has(e.rec.id))
    .sort((a, b) => (a.dateStr === b.dateStr
      ? String(a.rec.time).localeCompare(String(b.rec.time))
      : a.dateStr.localeCompare(b.dateStr)));
  const groups = {};
  legacyEvents.forEach((e) => {
    const key = e.rec.name + ' ' + roomOf(e.rec);
    (groups[key] = groups[key] || []).push(e);
  });
  Object.keys(groups).forEach((key) => {
    let open = null;
    groups[key].forEach((e) => {
      const room = roomOf(e.rec);
      const eventTime = String(e.rec.time || (e.rec.type === '入所' ? '00:00' : '23:59')).padStart(5, '0');
      const base = { ...baseOf(e.rec, room), stayId: null, legacy: true };
      if (e.rec.type === '入所') {
        if (open) stays.push(open); // two 入所 in a row: the earlier one stays open
        open = { ...base, checkinId: e.rec.id, checkoutId: null, start: e.dateStr, startTime: eventTime, startTimeRaw: e.rec.time || '', end: null, endTime: null, endTimeRaw: '' };
      } else if (open) {
        open.end = e.dateStr;
        open.endTime = eventTime;
        open.endTimeRaw = e.rec.time || '';
        open.checkoutId = e.rec.id;
        stays.push(open);
        open = null;
      } else {
        // 対応する入所がない退所記録は、滞在区間を持たない単独イベントとして扱う。
        stays.push({ ...base, checkinId: null, checkoutId: e.rec.id, start: e.dateStr, startTime: eventTime, startTimeRaw: e.rec.time || '', end: e.dateStr, endTime: eventTime, endTimeRaw: e.rec.time || '' });
      }
    });
    if (open) stays.push(open); // still staying
  });
  return stays;
}

function activeStaysOf(stays) { return stays.filter((s) => !isCancelled(s.status)); }

// Days the guest is present, departure day included — this is 利用日数.
// An open stay runs to the end of the window being asked about.
function stayDays(stay, winStart, winEnd) {
  const start = stay.start > winStart ? stay.start : winStart;
  const rawEnd = stay.end == null ? winEnd : stay.end;
  const end = rawEnd < winEnd ? rawEnd : winEnd;
  const out = [];
  eachDate(start, end, (d) => out.push(d));
  return out;
}

// Nights the bed is actually held. The departure day is excluded so a same-day
// turnover (A leaves in the morning, B arrives in the afternoon) is not a clash.
// 稼働率はこちらを母数にする（利用日ベースだと同日入替で100%を超えてしまう）。
function stayNights(stay, winStart, winEnd) {
  const lastNight = stay.end == null ? winEnd : shiftDate(stay.end, -1);
  const start = stay.start > winStart ? stay.start : winStart;
  const end = lastNight < winEnd ? lastNight : winEnd;
  const out = [];
  eachDate(start, end, (d) => out.push(d));
  return out;
}

function stayStartAt(stay) {
  return String(stay.start) + 'T' + String(stay.startTime || '00:00').padStart(5, '0');
}
function stayEndAt(stay) {
  return stay.end == null ? '9999-12-31T23:59' : String(stay.end) + 'T' + String(stay.endTime || '23:59').padStart(5, '0');
}
// [入所, 退所) の半開区間で比較する。同時刻の退所→入所は重複しない。
function staysOverlap(a, b) {
  return stayStartAt(a) < stayEndAt(b) && stayStartAt(b) < stayEndAt(a);
}
function stayIdentity(stay) {
  return stay.stayId || [stay.name, stay.room, stay.start, stay.startTime || '00:00'].join(' ');
}

// 重複しているペアと、カレンダー上で警告バッジを出すべき「日付×部屋」を返す。
// 保存時のブロックだけでは、バックアップ復元や旧形式データで入り込んだ重複に
// 誰も気づけない。
function conflictInfo(stays) {
  const pairs = [];
  const dateRoom = {};
  const mark = (dateStr, room) => { dateRoom[dateStr + '|' + room] = true; };
  for (let i = 0; i < stays.length; i++) {
    for (let j = i + 1; j < stays.length; j++) {
      const a = stays[i], b = stays[j];
      if (a.room !== b.room || !staysOverlap(a, b)) continue;
      pairs.push({ a, b });
      const an = new Set(stayNights(a, RANGE_MIN, RANGE_MAX));
      const bn = stayNights(b, RANGE_MIN, RANGE_MAX);
      let marked = false;
      bn.forEach((d) => { if (an.has(d)) { mark(d, a.room); marked = true; } });
      // 時刻だけが重なる同日入替（夜をまたがない重複）は、開始日にだけ印を付ける。
      if (!marked) mark(a.start > b.start ? a.start : b.start, a.room);
    }
  }
  return { pairs, dateRoom, count: pairs.length };
}

function rangeStats(stays, startStr, endStr) {
  const days = new Set();
  const users = new Set();
  let bedDays = 0;
  let bedNights = 0;
  let count = 0;
  stays.forEach((s) => {
    const ds = stayDays(s, startStr, endStr);
    if (!ds.length) return;
    count += 1;
    users.add(s.userId || s.name);
    bedDays += ds.length;
    bedNights += stayNights(s, startStr, endStr).length;
    ds.forEach((d) => days.add(d));
  });
  return { days: days.size, users: users.size, bedDays, bedNights, stays: count };
}

// dateStr -> roomId -> [stay], covering every day of a stay rather than only the
// two days that carry an 入所/退所 record.
function occupancyMap(stays, winStart, winEnd) {
  const map = {};
  stays.forEach((s) => {
    stayDays(s, winStart, winEnd).forEach((d) => {
      const byRoom = (map[d] = map[d] || {});
      (byRoom[s.room] = byRoom[s.room] || []).push(s);
    });
  });
  return map;
}

// 指定期間まるごと空いている部屋を返す。相談員からの「◯日〜◯日空いてますか」に
// 即答するための照会。
function vacancyFor(stays, rooms, startStr, endStr, checkinTime, checkoutTime) {
  return rooms.map((room) => {
    const probe = { start: startStr, startTime: checkinTime, end: endStr, endTime: checkoutTime, room: room.id };
    const blocking = stays.filter((s) => s.room === room.id && staysOverlap(s, probe));
    return {
      roomId: room.id,
      roomName: room.name,
      free: blocking.length === 0,
      blockers: blocking.map((s) => ({
        name: s.name,
        range: String(s.start).replace(/-/g, '/') + '〜' + (s.end ? String(s.end).replace(/-/g, '/') : '退所未定'),
      })),
    };
  });
}

function normalizeMaster(savedMaster) {
  const base = defaultMaster();
  const master = { ...base, ...(savedMaster || {}) };
  if (!Array.isArray(master.rooms) || !master.rooms.length) master.rooms = base.rooms;
  if (!Array.isArray(master.categories) || !master.categories.length) master.categories = base.categories;
  if (!Array.isArray(master.pickupOptions) || !master.pickupOptions.length) master.pickupOptions = base.pickupOptions;
  // v1 は利用者マスタが文字列配列だった。改名を記録へ波及させるため実体(id)を持たせる。
  if (!Array.isArray(master.users)) master.users = base.users;
  master.users = master.users.map((u, i) => (typeof u === 'string'
    ? { id: 'u' + (i + 1) + '-' + Math.random().toString(36).slice(2, 6), name: u }
    : { id: u.id || uid('u'), name: u.name || '' }));
  master.defaultCheckinTime = master.defaultCheckinTime || base.defaultCheckinTime;
  master.defaultCheckoutTime = master.defaultCheckoutTime || base.defaultCheckoutTime;
  master.operator = master.operator || '';

  // FAX帳票用の項目。既存の保存データには無いので補完する。
  ['postalCode', 'address', 'tel', 'fax', 'contactPerson'].forEach((k) => {
    master[k] = typeof master[k] === 'string' ? master[k] : '';
  });
  if (!Array.isArray(master.faxRecipients)) master.faxRecipients = [];
  master.faxRecipients = master.faxRecipients.map((r) => ({
    id: (r && r.id) || uid('f'),
    name: (r && r.name) || '',
    contact: (r && r.contact) || '',
    fax: (r && r.fax) || '',
    tel: (r && r.tel) || '',
    note: (r && r.note) || '',
  }));
  // 空文字は「意図して消した」なので既定値へ戻さない。未設定のときだけ埋める。
  const bp = master.boilerplate || {};
  master.boilerplate = {
    belongings: typeof bp.belongings === 'string' ? bp.belongings : base.boilerplate.belongings,
    notices: typeof bp.notices === 'string' ? bp.notices : base.boilerplate.notices,
  };
  // この画面は1部屋につき1名の運用。定員は部屋数から導出し、二重管理を防ぐ。
  master.capacity = master.rooms.length;
  return master;
}

// 保存済みレコードに userId / mealsByDate を補完する。
// userId は氏名一致で結び付ける（v1 は氏名しか持っていない）。
function normalizeData(data, master) {
  const byName = {};
  (master.users || []).forEach((u) => { byName[u.name] = u.id; });
  const roomIds = new Set(master.rooms.map((r) => r.id));
  const fallbackRoom = master.rooms[0].id;
  const stays = deriveStays(data, roomIds, fallbackRoom);
  const mealsByStayId = {};
  const mealsByRecordId = {};
  stays.forEach((s) => {
    const built = mealsOfStay(s);
    if (s.stayId) mealsByStayId[s.stayId] = built;
    if (s.checkinId) mealsByRecordId[s.checkinId] = built;
    if (s.checkoutId) mealsByRecordId[s.checkoutId] = built;
  });
  return mapRecords(data, (r) => {
    const next = { ...r };
    if (!next.userId && byName[next.name]) next.userId = byName[next.name];
    if (!next.mealsByDate) {
      next.mealsByDate = (next.stayId && mealsByStayId[next.stayId]) || mealsByRecordId[next.id] || {};
    }
    if (!next.meals) next.meals = { b: false, l: false, d: false };
    return next;
  });
}

function isBackupPayload(value) {
  return !!value && typeof value === 'object' && !!value.data && !!value.master;
}

function loadPersisted() {
  try {
    let raw = window.localStorage.getItem(STORAGE_KEY);
    let migrated = false;
    if (!raw) {
      raw = window.localStorage.getItem(LEGACY_STORAGE_KEY);
      migrated = !!raw;
    }
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!isBackupPayload(parsed)) return null;
    const master = normalizeMaster(parsed.master);
    return { data: normalizeData(parsed.data, master), master, migrated };
  } catch (err) {
    return null;
  }
}

function persist(state) {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify({ data: state.data, master: state.master }));
    return true;
  } catch (err) {
    return false;
  }
}

function saveAutoBackup(state) {
  try {
    window.localStorage.setItem(AUTO_BACKUP_KEY, JSON.stringify({
      format: 'ss-kanri-backup-v2', exportedAt: new Date().toISOString(),
      data: state.data, master: state.master,
    }));
    return true;
  } catch (err) {
    return false;
  }
}

function loadAutoBackup() {
  try {
    const value = JSON.parse(window.localStorage.getItem(AUTO_BACKUP_KEY)
      || window.localStorage.getItem(LEGACY_AUTO_BACKUP_KEY) || 'null');
    return isBackupPayload(value) ? value : null;
  } catch (err) {
    return null;
  }
}

// ── CSV ────────────────────────────────────────────────────────────────
// Excel で開くことが前提なので BOM 付き UTF-8 + CRLF。
function csvCell(v) {
  const s = v == null ? '' : String(v);
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}
function toCsv(rows) {
  return '﻿' + rows.map((r) => r.map(csvCell).join(',')).join('\r\n');
}
function downloadBlob(content, mime, filename) {
  const url = URL.createObjectURL(new Blob([content], { type: mime }));
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function monthsList() {
  // rolling window: 6 months back through 11 months ahead of today, so the month
  // picker's shortcut list crosses year boundaries automatically
  const out = [];
  const base = new Date(NOW.getFullYear(), NOW.getMonth(), 1);
  for (let i = -6; i <= 11; i++) {
    const d = new Date(base.getFullYear(), base.getMonth() + i, 1);
    out.push({ key: fmtDate(d).slice(0, 7), label: d.getFullYear() + '年' + (d.getMonth() + 1) + '月' });
  }
  return out;
}

class Component extends DCLogic {
  constructor(props) {
    super(props);
    const saved = loadPersisted();
    // 初期表示は「今月」。v1 は '2026-08' が固定で埋め込まれていたため、
    // 別の月に開いても毎回タブを押し直す必要があった。
    this.state = {
      view: 'month',
      currentMonth: THIS_MONTH,
      focusDate: TODAY,
      data: saved ? saved.data : sampleData(),
      drawer: null,
      master: saved ? saved.master : defaultMaster(),
      settingsOpen: false,
      userUsageOpen: false,
      vacancyOpen: false,
      vacancyRange: { start: TODAY, end: shiftDate(TODAY, 1) },
      kpiIncludeTentative: true,
      storageNotice: saved && saved.migrated
        ? '旧バージョンのデータを引き継ぎました。食事は日別管理になったため、既存予約の食事内容をご確認ください。'
        : '',
      reportRange: { start: monthStartOf(THIS_MONTH), end: monthEndOf(THIS_MONTH) },
      dragPreview: null,
    };
    // ドラッグ中の一時状態。1ドラッグにつき何度も setState したくないので、
    // 「日」単位で変化したときだけ dragPreview を更新する。
    this.drag = null;
    // ドラッグ終了後に発火する click で編集画面が開いてしまうのを防ぐ。
    this.suppressClick = false;
    this.pendingImportMerge = false;
    this.migratedFromLegacy = !!(saved && saved.migrated);
    this.onPointerMove = this.onPointerMove.bind(this);
    this.onPointerUp = this.onPointerUp.bind(this);
  }

  componentDidMount() {
    // 旧キーから読み込んだ場合、そのままでは何か編集するまで新キーへ書かれず、
    // 移行のお知らせが毎回出続ける。読み込めた時点で新形式として書き出す。
    if (this.migratedFromLegacy) persist(this.state);
    document.addEventListener('pointermove', this.onPointerMove);
    document.addEventListener('pointerup', this.onPointerUp);
    document.addEventListener('pointercancel', this.onPointerUp);
  }

  componentWillUnmount() {
    document.removeEventListener('pointermove', this.onPointerMove);
    document.removeEventListener('pointerup', this.onPointerUp);
    document.removeEventListener('pointercancel', this.onPointerUp);
  }

  componentDidUpdate(prevProps, prevState) {
    if (!prevState || prevState.data !== this.state.data || prevState.master !== this.state.master) {
      const failure = '保存に失敗しました。ブラウザの保存容量・プライベートモードを確認し、直ちにバックアップを保存してください。';
      if (!persist(this.state) && this.state.storageNotice !== failure) {
        this.setState({ storageNotice: failure });
      }
    }
  }

  // ── ナビゲーション ────────────────────────────────────────────────
  setView = (view) => this.setState({ view });
  selectMonth = (key) => this.setState({ currentMonth: key });
  prevMonth = () => this.setState((s) => ({ currentMonth: shiftMonth(s.currentMonth, -1) }));
  nextMonth = () => this.setState((s) => ({ currentMonth: shiftMonth(s.currentMonth, 1) }));
  goToday = () => this.setState({ currentMonth: THIS_MONTH, focusDate: TODAY });
  onMonthInputChange = (e) => { if (e.target.value) this.setState({ currentMonth: e.target.value }); };
  shiftFocusDate = (n) => this.setState((s) => ({ focusDate: shiftDate(s.focusDate, n) }));
  onFocusDateChange = (e) => { if (e.target.value) this.setState({ focusDate: e.target.value }); };

  toggleUserUsage = () => this.setState((s) => ({ userUsageOpen: !s.userUsageOpen }));
  toggleVacancy = () => this.setState((s) => ({ vacancyOpen: !s.vacancyOpen }));
  toggleKpiTentative = () => this.setState((s) => ({ kpiIncludeTentative: !s.kpiIncludeTentative }));
  updateVacancyStart = (e) => this.setState((s) => ({ vacancyRange: { ...s.vacancyRange, start: e.target.value } }));
  updateVacancyEnd = (e) => this.setState((s) => ({ vacancyRange: { ...s.vacancyRange, end: e.target.value } }));

  updateReportStart = (e) => this.setState((s) => ({ reportRange: { ...s.reportRange, start: e.target.value } }));
  updateReportEnd = (e) => this.setState((s) => ({ reportRange: { ...s.reportRange, end: e.target.value } }));
  setReportThisMonth = () => this.setState((s) => ({ reportRange: { start: monthStartOf(s.currentMonth), end: monthEndOf(s.currentMonth) } }));

  openSettings = () => this.setState({ settingsOpen: true });
  closeSettings = () => this.setState({ settingsOpen: false });
  updateMaster = (patch) => this.setState((s) => ({ master: normalizeMaster({ ...s.master, ...patch }) }));

  // ── バックアップ / 書き出し ───────────────────────────────────────
  exportBackup = () => {
    const payload = { format: 'ss-kanri-backup-v2', exportedAt: new Date().toISOString(), data: this.state.data, master: this.state.master };
    downloadBlob(JSON.stringify(payload, null, 2), 'application/json', 'ショートステイ管理表バックアップ-' + TODAY + '.json');
    this.setState({ storageNotice: 'バックアップファイルを保存しました。' });
  };

  // 明細CSV。請求・実績突合はExcelで行われるので、JSONバックアップとは別に必要。
  exportDetailCsv = () => {
    const master = this.state.master;
    const roomIds = new Set(master.rooms.map((r) => r.id));
    const roomName = (id) => (master.rooms.find((r) => r.id === id) || {}).name || '';
    const range = this.state.reportRange;
    const stays = deriveStays(this.state.data, roomIds, master.rooms[0].id)
      .filter((s) => stayDays(s, range.start, range.end).length)
      .sort((a, b) => (a.start + a.startTime).localeCompare(b.start + b.startTime));
    const rows = [['氏名', '区分', '部屋', '入所日', '入所時刻', '退所日', '退所時刻', '利用日数', '宿泊数', '期間内利用日数', 'ステータス', '送迎', '送迎方法', '送迎場所', '外部サービス', '朝食数', '昼食数', '夕食数', '備考', '最終更新', '更新者']];
    stays.forEach((s) => {
      const meals = mealsOfStay(s);
      const inRange = stayDays(s, range.start, range.end);
      let b = 0, l = 0, d = 0;
      inRange.forEach((day) => {
        const m = meals[day];
        if (!m) return;
        if (m.b) b++;
        if (m.l) l++;
        if (m.d) d++;
      });
      rows.push([
        s.name, s.category, roomName(s.room),
        s.start, s.startTimeRaw || '', s.end || '', s.endTimeRaw || '',
        stayDays(s, RANGE_MIN, RANGE_MAX).length,
        stayNights(s, RANGE_MIN, RANGE_MAX).length,
        inRange.length,
        s.status, s.pickup, s.pickupMethod, s.pickupLocation, s.externalService,
        b, l, d, s.notes, s.updatedAt, s.updatedBy,
      ]);
    });
    downloadBlob(toCsv(rows), 'text/csv', 'ショートステイ明細-' + range.start + '_' + range.end + '.csv');
    this.setState({ storageNotice: '明細CSVを保存しました（期間: ' + range.start + '〜' + range.end + '）。' });
  };

  // 区分別集計CSV。上司報告・請求区分の内訳がこれで一発で出る。
  exportSummaryCsv = () => {
    const master = this.state.master;
    const roomIds = new Set(master.rooms.map((r) => r.id));
    const range = this.state.reportRange;
    const stays = activeStaysOf(deriveStays(this.state.data, roomIds, master.rooms[0].id));
    const byCat = {};
    stays.forEach((s) => {
      const days = stayDays(s, range.start, range.end);
      if (!days.length) return;
      const row = (byCat[s.category] = byCat[s.category] || { users: new Set(), days: 0, nights: 0, count: 0 });
      row.users.add(s.userId || s.name);
      row.days += days.length;
      row.nights += stayNights(s, range.start, range.end).length;
      row.count += 1;
    });
    const rows = [['集計期間', range.start + '〜' + range.end], [], ['区分', '利用者数', '利用回数', '延べ利用日数', '延べ宿泊数']];
    let tu = new Set(), tc = 0, td = 0, tn = 0;
    Object.keys(byCat).sort().forEach((cat) => {
      const r = byCat[cat];
      rows.push([cat, r.users.size, r.count, r.days, r.nights]);
      r.users.forEach((u) => tu.add(u));
      tc += r.count; td += r.days; tn += r.nights;
    });
    rows.push(['合計', tu.size, tc, td, tn]);
    rows.push([]);
    rows.push(['利用者', '利用回数', '延べ利用日数']);
    const byUser = {};
    stays.forEach((s) => {
      const days = stayDays(s, range.start, range.end);
      if (!days.length) return;
      const row = (byUser[s.name] = byUser[s.name] || { count: 0, days: 0 });
      row.count += 1;
      row.days += days.length;
    });
    Object.keys(byUser).sort().forEach((n) => rows.push([n, byUser[n].count, byUser[n].days]));
    downloadBlob(toCsv(rows), 'text/csv', 'ショートステイ集計-' + range.start + '_' + range.end + '.csv');
    this.setState({ storageNotice: '集計CSVを保存しました。' });
  };

  // 食数一覧CSV（厨房への連絡用）。
  exportMealCsv = () => {
    const master = this.state.master;
    const roomIds = new Set(master.rooms.map((r) => r.id));
    const range = this.state.reportRange;
    const stays = activeStaysOf(deriveStays(this.state.data, roomIds, master.rooms[0].id));
    const perDay = {};
    stays.forEach((s) => {
      const meals = mealsOfStay(s);
      stayDays(s, range.start, range.end).forEach((day) => {
        const m = meals[day];
        if (!m) return;
        const row = (perDay[day] = perDay[day] || { b: 0, l: 0, d: 0, names: [] });
        if (m.b) row.b++;
        if (m.l) row.l++;
        if (m.d) row.d++;
        row.names.push(s.name + '(' + (m.b ? '朝' : '') + (m.l ? '昼' : '') + (m.d ? '夕' : '') + ')');
      });
    });
    const rows = [['日付', '曜日', '朝食', '昼食', '夕食', '内訳']];
    eachDate(range.start, range.end, (day) => {
      const r = perDay[day] || { b: 0, l: 0, d: 0, names: [] };
      rows.push([day, WEEKDAY_LABELS[parseDate(day).getDay()], r.b, r.l, r.d, r.names.join(' / ')]);
    });
    downloadBlob(toCsv(rows), 'text/csv', 'ショートステイ食数-' + range.start + '_' + range.end + '.csv');
    this.setState({ storageNotice: '食数CSVを保存しました。' });
  };

  importBackup = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const merge = !!this.pendingImportMerge;
    this.pendingImportMerge = false;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(String(reader.result || ''));
        if (!isBackupPayload(parsed)) throw new Error('形式不正');
        const master = normalizeMaster(parsed.master);
        const incoming = normalizeData(parsed.data, master);
        if (!merge) {
          if (!window.confirm('現在のデータを、このバックアップで置き換えます。よろしいですか？')) return;
          this.setState({ data: incoming, master, drawer: null, storageNotice: 'バックアップを復元しました。' });
          return;
        }
        if (!window.confirm('このバックアップを現在のデータへ取り込みます（同じ予約は更新日時が新しい方を採用）。よろしいですか？')) return;
        this.setState((s) => {
          const merged = this.mergeData(s.data, incoming);
          return {
            data: merged.data, drawer: null,
            master: { ...s.master, users: this.mergeUsers(s.master.users, master.users) },
            storageNotice: '取り込みました（追加 ' + merged.added + ' 件 / 更新 ' + merged.updated + ' 件）。',
          };
        });
      } catch (err) {
        this.setState({ storageNotice: '復元できませんでした。ショートステイ管理表のバックアップJSONを選択してください。' });
      }
    };
    reader.readAsText(file, 'UTF-8');
    e.target.value = '';
  };

  markImportMerge = () => { this.pendingImportMerge = true; };

  // 端末をまたいだ持ち寄りマージ。同一レコードidは updatedAt が新しい方を採用する。
  mergeData = (current, incoming) => {
    const out = JSON.parse(JSON.stringify(current || {}));
    const index = {};
    recordEntries(out).forEach((e) => { index[e.rec.id] = e; });
    let added = 0, updated = 0;
    recordEntries(incoming).forEach((e) => {
      const existing = index[e.rec.id];
      if (!existing) {
        const mk = monthKeyOf(e.dateStr);
        out[mk] = out[mk] || {};
        out[mk][e.dateStr] = out[mk][e.dateStr] || [];
        out[mk][e.dateStr].push(e.rec);
        added += 1;
        return;
      }
      const a = String(existing.rec.updatedAt || '');
      const b = String(e.rec.updatedAt || '');
      if (b > a) {
        const mk = monthKeyOf(existing.dateStr);
        out[mk][existing.dateStr] = out[mk][existing.dateStr].filter((r) => r.id !== e.rec.id);
        const nmk = monthKeyOf(e.dateStr);
        out[nmk] = out[nmk] || {};
        out[nmk][e.dateStr] = out[nmk][e.dateStr] || [];
        out[nmk][e.dateStr].push(e.rec);
        updated += 1;
      }
    });
    return { data: out, added, updated };
  };

  mergeUsers = (current, incoming) => {
    const out = [...current];
    const seen = new Set(current.map((u) => u.id));
    const names = new Set(current.map((u) => u.name));
    incoming.forEach((u) => {
      if (seen.has(u.id) || names.has(u.name)) return;
      out.push(u);
      seen.add(u.id);
      names.add(u.name);
    });
    return out;
  };

  restoreAutoBackup = () => {
    const backup = loadAutoBackup();
    if (!backup) { this.setState({ storageNotice: '初期化前の自動バックアップはありません。' }); return; }
    if (!window.confirm('初期化前の自動バックアップを復元します。現在のデータは置き換えられます。よろしいですか？')) return;
    const master = normalizeMaster(backup.master);
    this.setState({ data: normalizeData(backup.data, master), master, drawer: null, storageNotice: '初期化前の自動バックアップを復元しました。' });
  };

  resetAll = () => {
    if (!window.confirm('現在のデータを初期データに戻します。初期化前の状態は自動バックアップされます。よろしいですか？')) return;
    this.setState((s) => {
      const backedUp = saveAutoBackup(s);
      if (!backedUp) {
        return { storageNotice: '初期化前の自動バックアップ保存に失敗したため、初期化を中止しました。先に手動バックアップを保存してください。' };
      }
      return {
        data: sampleData(), master: defaultMaster(), drawer: null, settingsOpen: false,
        storageNotice: '初期データへ戻しました。必要なら自動バックアップから復元できます。',
      };
    });
  };

  // ── マスタ編集 ────────────────────────────────────────────────────
  // 氏名の改名は記録側へも波及させる。区分・送迎は波及していたのに利用者だけ
  // 抜けており、マスタと記録が食い違う状態になっていた。
  updateUser = (idx, val) => this.setState((s) => {
    const prev = s.master.users[idx];
    if (!prev) return s;
    const users = s.master.users.map((u, i) => (i === idx ? { ...u, name: val } : u));
    return {
      master: { ...s.master, users },
      data: mapRecords(s.data, (r) => (
        (r.userId && r.userId === prev.id) || (!r.userId && r.name === prev.name)
          ? { ...r, name: val, userId: prev.id }
          : r
      )),
    };
  });
  addUser = () => this.setState((s) => ({ master: { ...s.master, users: [...s.master.users, { id: uid('u'), name: '' }] } }));
  removeUser = (idx) => this.setState((s) => ({ master: { ...s.master, users: s.master.users.filter((_, i) => i !== idx) } }));

  // Records reference categories by LABEL, so a rename has to be carried through
  // the stored records or every existing reservation falls back to grey.
  updateCategory = (idx, patch) => this.setState((s) => {
    const prev = s.master.categories[idx];
    if (!prev) return s;
    const categories = s.master.categories.map((c, i) => (i === idx ? { ...c, ...patch } : c));
    const out = { master: { ...s.master, categories } };
    if (typeof patch.label === 'string' && patch.label !== prev.label) {
      out.data = mapRecords(s.data, (r) => (r.category === prev.label ? { ...r, category: patch.label } : r));
    }
    return out;
  });
  addCategory = () => this.setState((s) => ({ master: { ...s.master, categories: [...s.master.categories, { id: uid('c'), label: '新区分', color: '#7a4f8c' }] } }));
  removeCategory = (idx) => this.setState((s) => {
    if (s.master.categories.length <= 1) return s;
    const gone = s.master.categories[idx];
    const categories = s.master.categories.filter((_, i) => i !== idx);
    return {
      master: { ...s.master, categories },
      data: mapRecords(s.data, (r) => (r.category === gone.label ? { ...r, category: categories[0].label } : r)),
    };
  });

  updatePickupOption = (idx, patch) => this.setState((s) => {
    const prev = s.master.pickupOptions[idx];
    if (!prev) return s;
    const pickupOptions = s.master.pickupOptions.map((p, i) => (i === idx ? { ...p, ...patch } : p));
    const out = { master: { ...s.master, pickupOptions } };
    if (typeof patch.label === 'string' && patch.label !== prev.label) {
      out.data = mapRecords(s.data, (r) => (r.pickup === prev.label ? { ...r, pickup: patch.label } : r));
    }
    return out;
  });
  addPickupOption = () => this.setState((s) => ({ master: { ...s.master, pickupOptions: [...s.master.pickupOptions, { id: uid('p'), label: '新規', highlight: false }] } }));
  removePickupOption = (idx) => this.setState((s) => {
    if (s.master.pickupOptions.length <= 1) return s;
    const gone = s.master.pickupOptions[idx];
    const pickupOptions = s.master.pickupOptions.filter((_, i) => i !== idx);
    return {
      master: { ...s.master, pickupOptions },
      data: mapRecords(s.data, (r) => (r.pickup === gone.label ? { ...r, pickup: pickupOptions[0].label } : r)),
    };
  });

  // ── FAX送付先マスタ ──────────────────────────────────────────────
  addFaxRecipient = () => this.setState((s) => ({
    master: { ...s.master, faxRecipients: [...s.master.faxRecipients, { id: uid('f'), name: '', contact: '', fax: '', tel: '', note: '' }] },
  }));
  updateFaxRecipient = (idx, patch) => this.setState((s) => ({
    master: { ...s.master, faxRecipients: s.master.faxRecipients.map((r, i) => (i === idx ? { ...r, ...patch } : r)) },
  }));
  removeFaxRecipient = (idx) => this.setState((s) => ({
    master: { ...s.master, faxRecipients: s.master.faxRecipients.filter((_, i) => i !== idx) },
  }));

  updateBoilerplate = (patch) => this.setState((s) => ({
    master: { ...s.master, boilerplate: { ...s.master.boilerplate, ...patch } },
  }));

  updateRoom = (idx, val) => this.setState((s) => {
    const rooms = s.master.rooms.map((r, i) => (i === idx ? { ...r, name: val } : r));
    return { master: { ...s.master, rooms } };
  });
  addRoom = () => this.setState((s) => {
    const rooms = [...s.master.rooms, { id: uid('r'), name: '部屋' + (s.master.rooms.length + 1) }];
    return { master: { ...s.master, rooms, capacity: rooms.length } };
  });
  // Reservations in a deleted room would keep matching no lane at all: invisible
  // and uneditable, yet still counted. Move them to the first remaining room.
  removeRoom = (idx) => this.setState((s) => {
    if (s.master.rooms.length <= 1) return s;
    const gone = s.master.rooms[idx];
    const rooms = s.master.rooms.filter((_, i) => i !== idx);
    return {
      master: { ...s.master, rooms, capacity: rooms.length },
      data: mapRecords(s.data, (r) => (r.room === gone.id ? { ...r, room: rooms[0].id } : r)),
    };
  });

  // ── ドロワー ──────────────────────────────────────────────────────
  openAdd = (dateStr, roomId, preset) => this.setState((s) => {
    const m = s.master;
    const startDate = (preset && preset.startDate) || dateStr;
    const endDate = (preset && preset.endDate) || shiftDate(startDate, 1);
    const startTime = m.defaultCheckinTime;
    const endTime = m.defaultCheckoutTime;
    return { drawer: {
      isNew: true, isStayEditor: true, stayId: null,
      dateStr: startDate, originalDateStr: null, id: null,
      startDate, startTime, endDate, endTime,
      type: '入所', time: startTime,
      name: (preset && preset.name) || '',
      userId: (preset && preset.userId) || null,
      category: (preset && preset.category) || m.categories[0]?.label || '',
      mealsByDate: buildMealsByDate(startDate, endDate, null, null),
      mealsTouched: {}, mealsOpen: false,
      pickup: m.pickupOptions[0]?.label || '', status: '確定',
      pickupMethod: '', pickupLocation: '', externalService: '', detailsOpen: false,
      notes: '', room: roomId || m.rooms[0]?.id, error: null,
      updatedAt: '', updatedBy: '',
    } };
  });

  // 旧形式（stayId 無し）の単独レコードを編集するための入口。
  openEditLegacy = (dateStr, id) => {
    const list = (this.state.data[monthKeyOf(dateStr)] || {})[dateStr] || [];
    const rec = list.find((x) => x.id === id);
    if (!rec) return;
    this.setState({ drawer: {
      isNew: false, isStayEditor: false, stayId: null, dateStr, originalDateStr: dateStr, id,
      type: rec.type, time: rec.time, name: rec.name, userId: rec.userId || null, category: rec.category,
      mealB: !!(rec.meals && rec.meals.b), mealL: !!(rec.meals && rec.meals.l), mealD: !!(rec.meals && rec.meals.d),
      mealsByDate: rec.mealsByDate || {}, mealsTouched: allTouched(rec.mealsByDate), mealsOpen: false,
      pickup: rec.pickup, status: rec.status || '確定',
      pickupMethod: rec.pickupMethod || '', pickupLocation: rec.pickupLocation || '', externalService: rec.externalService || '',
      detailsOpen: !!(rec.pickupMethod || rec.pickupLocation || rec.externalService),
      notes: rec.notes, room: rec.room || this.state.master.rooms[0]?.id,
      updatedAt: rec.updatedAt || '', updatedBy: rec.updatedBy || '',
      error: null,
    } });
  };

  openEditStay = (stay) => {
    if (!stay.stayId) {
      this.openEditLegacy(stay.start, stay.checkinId || stay.checkoutId);
      return;
    }
    const endDate = stay.end || shiftDate(stay.start, 1);
    this.setState({ drawer: {
      isNew: false, isStayEditor: true, stayId: stay.stayId, id: stay.checkinId,
      dateStr: stay.start, originalDateStr: stay.start,
      startDate: stay.start, startTime: stay.startTimeRaw != null ? stay.startTimeRaw : (stay.startTime || ''),
      endDate, endTime: stay.endTimeRaw != null ? stay.endTimeRaw : (stay.endTime || ''),
      type: '入所', time: stay.startTime, name: stay.name, userId: stay.userId || null, category: stay.category,
      // 保存済みの食事内容は職員が決めた結果なので、すべて「手入力済み」として扱う。
      mealsByDate: buildMealsByDate(stay.start, endDate, stay.mealsByDate, stay.meals),
      mealsTouched: allTouched(stay.mealsByDate), mealsOpen: false,
      pickup: stay.pickup, status: stay.status || '確定',
      pickupMethod: stay.pickupMethod, pickupLocation: stay.pickupLocation, externalService: stay.externalService,
      detailsOpen: !!(stay.pickupMethod || stay.pickupLocation || stay.externalService),
      notes: stay.notes, room: stay.room,
      updatedAt: stay.updatedAt, updatedBy: stay.updatedBy,
      error: null,
    } });
  };

  openEdit = (dateStr, id) => {
    const master = this.state.master;
    const roomIds = new Set(master.rooms.map((r) => r.id));
    const stays = deriveStays(this.state.data, roomIds, master.rooms[0].id);
    const stay = stays.find((s) => s.checkinId === id || s.checkoutId === id);
    if (stay) { this.openEditStay(stay); return; }
    this.openEditLegacy(dateStr, id);
  };

  closeDrawer = () => this.setState({ drawer: null });

  setDrawerField = (patch) => this.setState((s) => {
    const next = { ...s.drawer, ...patch, error: null };
    // 期間が変わったら食事表を作り直す。手入力した日はそのまま残し、触っていない
    // 日は新しい期間に合わせた標準パターンへ引き直す。
    if (patch.startDate || patch.endDate) {
      next.mealsByDate = buildMealsByDate(
        next.startDate, next.endDate,
        pickTouched(s.drawer.mealsByDate, s.drawer.mealsTouched), null,
      );
    }
    return { drawer: next };
  });

  toggleDrawerMeal = (dateStr, key) => this.setState((s) => {
    const table = { ...(s.drawer.mealsByDate || {}) };
    const cur = table[dateStr] || { b: false, l: false, d: false };
    table[dateStr] = { ...cur, [key]: !cur[key] };
    return { drawer: { ...s.drawer, mealsByDate: table, mealsTouched: { ...(s.drawer.mealsTouched || {}), [dateStr]: true }, error: null } };
  });

  applyMealPreset = (values) => this.setState((s) => {
    const d = s.drawer;
    const table = {};
    eachDate(d.startDate, d.endDate, (day) => { table[day] = { ...values }; });
    return { drawer: { ...d, mealsByDate: table, mealsTouched: allTouched(table), error: null } };
  });

  // 「標準」は手入力の解除でもある。以後は日程変更に追従させる。
  applyMealStandard = () => this.setState((s) => {
    const d = s.drawer;
    const table = {};
    eachDate(d.startDate, d.endDate, (day) => { table[day] = defaultMealsFor(day, d.startDate, d.endDate); });
    return { drawer: { ...d, mealsByDate: table, mealsTouched: {}, error: null } };
  });

  toggleDrawerMeals = () => this.setState((s) => ({ drawer: { ...s.drawer, mealsOpen: !s.drawer.mealsOpen } }));

  // 延泊・短縮。月表示から日付を打ち直さずに1日単位で調整する。
  extendStay = (n) => this.setState((s) => {
    const d = s.drawer;
    const endDate = shiftDate(d.endDate, n);
    if (endDate < d.startDate) return { drawer: { ...d, error: '退所日を入所日より前にはできません。' } };
    return { drawer: { ...d, endDate, mealsByDate: buildMealsByDate(d.startDate, endDate, pickTouched(d.mealsByDate, d.mealsTouched), null), error: null } };
  });

  onNameChange = (e) => {
    const val = e.target.value;
    const match = this.state.master.users.find((u) => u.name === val);
    this.setDrawerField({ name: val, userId: match ? match.id : null });
  };

  applyPrevious = () => this.setState((s) => {
    const d = s.drawer;
    const name = String(d?.name || '').trim();
    if (!d || !name) return { drawer: { ...d, error: '先に氏名を入力または選択してください。' } };
    const previous = recordEntries(s.data)
      .filter((e) => e.rec.name === name && e.rec.type === '入所' && !isCancelled(e.rec.status) && e.dateStr <= d.startDate)
      .sort((a, b) => (b.dateStr + String(b.rec.time || '')).localeCompare(a.dateStr + String(a.rec.time || '')))[0];
    if (!previous) return { drawer: { ...d, error: 'この利用者の過去の入所記録がありません。' } };
    const r = previous.rec;
    return { drawer: { ...d,
      category: r.category,
      pickup: r.pickup, pickupMethod: r.pickupMethod || '', pickupLocation: r.pickupLocation || '', externalService: r.externalService || '',
      notes: r.notes || '', room: r.room || d.room,
      detailsOpen: !!(r.pickupMethod || r.pickupLocation || r.externalService), error: null } };
  });

  duplicateNextWeek = () => this.setState((s) => {
    const d = s.drawer;
    if (!d) return null;
    const startDate = shiftDate(d.startDate, 7);
    const endDate = shiftDate(d.endDate, 7);
    return { drawer: { ...d, isNew: true, stayId: null, id: null,
      startDate, endDate, dateStr: startDate, originalDateStr: null,
      mealsByDate: buildMealsByDate(startDate, endDate, null, null), mealsTouched: {},
      error: null } };
  });

  // ── 保存 ──────────────────────────────────────────────────────────
  stamp = () => ({ updatedAt: new Date().toISOString().slice(0, 16).replace('T', ' '), updatedBy: this.state.master.operator || '' });

  saveStayDrawer = () => {
    const d = this.state.drawer;
    if (!d) return;
    const name = String(d.name || '').trim();
    if (!name) { this.setState({ drawer: { ...d, error: '氏名を入力してください。' } }); return; }
    if (!d.startDate || !d.endDate || !d.startTime || !d.endTime) {
      this.setState({ drawer: { ...d, error: '入所日時と退所予定日時を入力してください。' } }); return;
    }
    const startAt = d.startDate + 'T' + d.startTime;
    const endAt = d.endDate + 'T' + d.endTime;
    if (endAt <= startAt) {
      this.setState({ drawer: { ...d, error: '退所予定日時は入所日時より後にしてください。' } }); return;
    }
    const master = this.state.master;
    const stayId = d.stayId || uid('s');
    const matched = master.users.find((u) => u.name === name);
    const shared = {
      stayId, name, userId: d.userId || (matched ? matched.id : null),
      category: d.category,
      meals: { b: false, l: false, d: false },
      mealsByDate: buildMealsByDate(d.startDate, d.endDate, d.mealsByDate, null),
      pickup: d.pickup, status: d.status || '確定', notes: d.notes || '', room: d.room,
      pickupMethod: d.pickupMethod || '', pickupLocation: d.pickupLocation || '', externalService: d.externalService || '',
      ...this.stamp(),
    };
    const checkin = { ...shared, id: stayId + '-in', type: '入所', time: d.startTime };
    const checkout = { ...shared, id: stayId + '-out', type: '退所', time: d.endTime };
    let nextData = removeRecords(this.state.data, (r) => r.stayId === stayId);
    nextData = applyRecord(nextData, d.startDate, checkin, true);
    nextData = applyRecord(nextData, d.endDate, checkout, true);

    if (!isCancelled(shared.status)) {
      const roomIds = new Set(master.rooms.map((r) => r.id));
      const stays = activeStaysOf(deriveStays(nextData, roomIds, master.rooms[0].id));
      const target = stays.find((s) => s.stayId === stayId);
      if (target) {
        const conflict = stays.find((s) => s !== target && s.room === target.room && staysOverlap(s, target));
        if (conflict) {
          const roomName = (master.rooms.find((r) => r.id === target.room) || {}).name || '';
          const overlapStart = stayStartAt(target) > stayStartAt(conflict) ? stayStartAt(target) : stayStartAt(conflict);
          const overlapEnd = stayEndAt(target) < stayEndAt(conflict) ? stayEndAt(target) : stayEndAt(conflict);
          const endLabel = overlapEnd.startsWith('9999-') ? '退所未定' : overlapEnd.replace(/-/g, '/').replace('T', ' ');
          this.setState({ drawer: { ...d, error: roomName + 'は ' + overlapStart.replace(/-/g, '/').replace('T', ' ') + '〜' + endLabel + ' に、' + conflict.name + ' さんの予約と重複します。' } });
          return;
        }
      }
    }
    const users = master.users.some((u) => u.name === name)
      ? master.users
      : [...master.users, { id: shared.userId || uid('u'), name }];
    this.setState({
      data: nextData,
      master: { ...master, users },
      currentMonth: monthKeyOf(d.startDate),
      drawer: null,
      storageNotice: isCancelled(d.status) ? '予約をキャンセル記録として保存しました。' : '保存しました。',
    });
  };

  saveDrawer = () => {
    const d = this.state.drawer;
    if (!d) return;
    if (d.isStayEditor) { this.saveStayDrawer(); return; }
    const master = this.state.master;
    const name = String(d.name || '').trim();
    if (!name) {
      this.setState({ drawer: { ...d, error: '氏名を入力してください。' } });
      return;
    }
    const record = {
      id: d.id || uid('r'),
      type: d.type, time: d.time, name, userId: d.userId || null, category: d.category,
      meals: { b: d.mealB, l: d.mealL, d: d.mealD },
      mealsByDate: d.mealsByDate || {},
      pickup: d.pickup, notes: d.notes, room: d.room,
      status: d.status || '確定', pickupMethod: d.pickupMethod || '', pickupLocation: d.pickupLocation || '', externalService: d.externalService || '',
      ...this.stamp(),
    };
    // 日付を変更した既存記録は、元の日付から削除して新しい日付へ移動する。
    let baseData = this.state.data;
    const originalDateStr = d.originalDateStr || d.dateStr;
    if (!d.isNew && originalDateStr !== d.dateStr) {
      const originalMonthKey = monthKeyOf(originalDateStr);
      const movedData = { ...this.state.data };
      const originalMonth = { ...(movedData[originalMonthKey] || {}) };
      originalMonth[originalDateStr] = (originalMonth[originalDateStr] || []).filter((x) => x.id !== d.id);
      movedData[originalMonthKey] = originalMonth;
      baseData = movedData;
    }
    const nextData = applyRecord(baseData, d.dateStr, record, d.isNew);

    // Validate against the stays the save would produce, so a clash with the
    // middle of somebody else's stay is caught too, not just with a record.
    const roomIds = new Set(master.rooms.map((r) => r.id));
    const fallbackRoom = master.rooms[0].id;
    const stays = activeStaysOf(deriveStays(nextData, roomIds, fallbackRoom));
    const targetRoom = roomIds.has(d.room) ? d.room : fallbackRoom;
    const roomName = (master.rooms.find((r) => r.id === targetRoom) || {}).name || '';

    // 今回保存する入所・退所が属する滞在だけを検査する。
    // 入所を先に登録する通常の操作では、退所未定のために将来予約まで即時ブロックしない。
    const targetTime = String(d.time || (d.type === '入所' ? '00:00' : '23:59')).padStart(5, '0');
    const targetStays = stays.filter((s) => {
      if (s.name !== name || s.room !== targetRoom) return false;
      return d.type === '入所'
        ? s.start === d.dateStr && String(s.startTime || '00:00').padStart(5, '0') === targetTime
        : s.end === d.dateStr && String(s.endTime || '23:59').padStart(5, '0') === targetTime;
    });
    const targets = new Set(targetStays);
    for (let i = 0; i < stays.length; i++) {
      for (let j = i + 1; j < stays.length; j++) {
        const a = stays[i], b = stays[j];
        if (!targets.has(a) && !targets.has(b)) continue;
        if (a.room !== b.room || !staysOverlap(a, b)) continue;
        const targetStay = targets.has(a) ? a : b;
        const otherStay = targetStay === a ? b : a;
        // 新規の退所未定入所は、入所時点ですでに使用中の場合だけ止める。
        // 後日の予約は、退所記録を追加した時点で完成した区間として再検査する。
        if (d.isNew && d.type === '入所' && targetStay.end == null && stayStartAt(otherStay) > stayStartAt(targetStay)) continue;
        const overlapStart = stayStartAt(a) > stayStartAt(b) ? stayStartAt(a) : stayStartAt(b);
        const overlapEnd = stayEndAt(a) < stayEndAt(b) ? stayEndAt(a) : stayEndAt(b);
        const nameOfRoom = (master.rooms.find((r) => r.id === a.room) || {}).name || roomName;
        const startLabel = overlapStart.replace(/-/g, '/').replace('T', ' ');
        const endLabel = overlapEnd.startsWith('9999-') ? '退所未定' : overlapEnd.replace(/-/g, '/').replace('T', ' ');
        this.setState({ drawer: { ...d, error: nameOfRoom + 'は ' + startLabel + '〜' + endLabel + ' に、' + a.name + ' さんと ' + b.name + ' さんの利用時間が重複します。退所・入所時刻を確認してください。' } });
        return;
      }
    }

    this.setState({ data: nextData, drawer: null });
  };

  deleteDrawer = () => {
    const d = this.state.drawer;
    if (!d) return;
    if (d.isStayEditor && d.stayId) {
      if (!window.confirm('この利用の入所・退所記録を削除します。よろしいですか？')) return;
      this.setState({ data: removeRecords(this.state.data, (r) => r.stayId === d.stayId), drawer: null, storageNotice: '利用記録を削除しました。' });
      return;
    }
    const deleteDateStr = d.originalDateStr || d.dateStr;
    const mk = monthKeyOf(deleteDateStr);
    const data = { ...this.state.data };
    const monthData = { ...(data[mk] || {}) };
    monthData[deleteDateStr] = (monthData[deleteDateStr] || []).filter((x) => x.id !== d.id);
    data[mk] = monthData;
    this.setState({ data, drawer: null });
  };

  // ── タイムラインのドラッグ ────────────────────────────────────────
  // 予約バーの本体をドラッグ=日程移動、左右の取っ手をドラッグ=期間の伸縮、
  // 空きトラックのドラッグ=その期間で新規作成。
  startBarDrag = (mode, stay) => (e) => {
    if (e.button != null && e.button !== 0) return;
    // バー上の操作はトラック側の「新規作成ドラッグ」を起動させない。
    // 取っ手の上での操作は、バー本体の「移動ドラッグ」も起動させない。
    if (e.stopPropagation) e.stopPropagation();
    // 旧形式(stayId なし)の滞在は入所・退所が別レコードとして独立しているため、
    // まとめて動かせない。ドラッグは受け付けずクリック編集だけにする。
    if (!stay.stayId) return;
    const track = e.target && e.target.closest ? e.target.closest('[data-track]') : null;
    if (!track) return;
    this.drag = {
      mode, stayId: stay.stayId, roomId: stay.room,
      origStart: stay.start, origEnd: stay.end || stay.start,
      startX: e.clientX, trackLeft: track.getBoundingClientRect().left,
      gridStart: track.getAttribute('data-grid-start'),
      delta: 0, moved: false,
    };
  };

  startTrackDrag = (roomId, gridStart) => (e) => {
    if (e.button != null && e.button !== 0) return;
    const track = e.target && e.target.closest ? e.target.closest('[data-track]') : null;
    if (!track) return;
    const left = track.getBoundingClientRect().left;
    const dayIndex = Math.max(0, Math.floor((e.clientX - left) / DAY_W));
    this.drag = {
      mode: 'create', roomId, gridStart,
      anchorIndex: dayIndex, curIndex: dayIndex,
      startX: e.clientX, trackLeft: left, moved: false,
    };
    this.setState({ dragPreview: { mode: 'create', roomId, start: shiftDate(gridStart, dayIndex), end: shiftDate(gridStart, dayIndex + 1) } });
  };

  onPointerMove(e) {
    const drag = this.drag;
    if (!drag) return;
    if (drag.mode === 'create') {
      const idx = Math.max(0, Math.floor((e.clientX - drag.trackLeft) / DAY_W));
      if (idx === drag.curIndex) return;
      drag.curIndex = idx;
      drag.moved = true;
      const a = Math.min(drag.anchorIndex, idx);
      const b = Math.max(drag.anchorIndex, idx);
      this.setState({ dragPreview: { mode: 'create', roomId: drag.roomId, start: shiftDate(drag.gridStart, a), end: shiftDate(drag.gridStart, b + 1) } });
      return;
    }
    const delta = Math.round((e.clientX - drag.startX) / DAY_W);
    if (delta === drag.delta) return;
    drag.delta = delta;
    drag.moved = drag.moved || delta !== 0;
    let start = drag.origStart;
    let end = drag.origEnd;
    if (drag.mode === 'move') { start = shiftDate(start, delta); end = shiftDate(end, delta); }
    else if (drag.mode === 'start') { start = shiftDate(start, delta); if (start > end) start = end; }
    else if (drag.mode === 'end') { end = shiftDate(end, delta); if (end < start) end = start; }
    this.setState({ dragPreview: { mode: drag.mode, stayId: drag.stayId, roomId: drag.roomId, start, end } });
  }

  onPointerUp() {
    const drag = this.drag;
    this.drag = null;
    if (!drag) return;
    const preview = this.state.dragPreview;
    this.setState({ dragPreview: null });
    if (drag.moved) {
      // pointerup の直後に click が続く。ドラッグで動かした場合まで編集画面を
      // 開いてしまわないよう、この1回だけクリックを無視する。
      this.suppressClick = true;
      window.setTimeout(() => { this.suppressClick = false; }, 0);
    }
    if (!preview || !drag.moved) {
      // 動かさなかった＝ただのクリック。作成ドラッグなら1泊で新規作成する。
      if (drag.mode === 'create' && preview) {
        this.openAdd(preview.start, drag.roomId, { startDate: preview.start, endDate: preview.end });
      }
      return;
    }
    if (drag.mode === 'create') {
      this.openAdd(preview.start, drag.roomId, { startDate: preview.start, endDate: preview.end });
      return;
    }
    this.applyStayDates(drag.stayId, preview.start, preview.end);
  }

  // ドラッグ結果を記録へ反映する。重複したら元に戻して理由を伝える。
  applyStayDates = (stayId, startDate, endDate) => {
    const master = this.state.master;
    const entries = recordEntries(this.state.data).filter((e) => e.rec.stayId === stayId);
    const checkin = entries.find((e) => e.rec.type === '入所');
    const checkout = entries.find((e) => e.rec.type === '退所');
    if (!checkin) return;
    const stampVals = this.stamp();
    const mealsByDate = buildMealsByDate(startDate, endDate, checkin.rec.mealsByDate, checkin.rec.meals);
    const nextCheckin = { ...checkin.rec, mealsByDate, ...stampVals };
    const nextCheckout = checkout ? { ...checkout.rec, mealsByDate, ...stampVals } : null;
    let nextData = removeRecords(this.state.data, (r) => r.stayId === stayId);
    nextData = applyRecord(nextData, startDate, nextCheckin, true);
    if (nextCheckout) nextData = applyRecord(nextData, endDate, nextCheckout, true);

    if (!isCancelled(nextCheckin.status)) {
      const roomIds = new Set(master.rooms.map((r) => r.id));
      const stays = activeStaysOf(deriveStays(nextData, roomIds, master.rooms[0].id));
      const target = stays.find((s) => s.stayId === stayId);
      const conflict = target && stays.find((s) => s !== target && s.room === target.room && staysOverlap(s, target));
      if (conflict) {
        this.setState({ storageNotice: conflict.name + ' さんの予約と重複するため、移動を取り消しました。' });
        return;
      }
    }
    this.setState({ data: nextData, storageNotice: '日程を ' + startDate.replace(/-/g, '/') + '〜' + endDate.replace(/-/g, '/') + ' に変更しました。' });
  };

  // ── 描画 ──────────────────────────────────────────────────────────
  renderVals() {
    const primaryAccent = this.props.primaryAccent ?? '#2f6690';
    const density = this.props.density ?? 'comfortable';
    const pickupHighlight = this.props.pickupHighlight ?? true;
    const cellMinHeight = density === 'compact' ? '104px' : '150px';
    const master = this.state.master;
    const view = this.state.view;
    const categoryByLabel = {};
    master.categories.forEach((c) => { categoryByLabel[c.label] = c; });
    const pickupByLabel = {};
    master.pickupOptions.forEach((p) => { pickupByLabel[p.label] = p; });
    const roomNameOf = (id) => (master.rooms.find((r) => r.id === id) || {}).name || '';

    const currentMonth = this.state.currentMonth;
    const currentMonthLabel = monthLabelOf(currentMonth);
    const monthData = this.state.data[currentMonth] || {};
    const recordsOn = (dateStr) => (this.state.data[monthKeyOf(dateStr)] || {})[dateStr] || [];

    const roomIds = new Set(master.rooms.map((r) => r.id));
    const fallbackRoom = master.rooms[0].id;
    // A record pointing at a room that no longer exists still has to show up
    // somewhere, so it lands in the first lane rather than vanishing.
    const laneIdFor = (rec) => (rec.room && roomIds.has(rec.room) ? rec.room : fallbackRoom);
    const allStays = deriveStays(this.state.data, roomIds, fallbackRoom);
    const activeStays = activeStaysOf(allStays);
    const conflicts = conflictInfo(activeStays);
    // KPI の母数。仮予約・調整中を含めるかは切り替えられる。
    const kpiStays = this.state.kpiIncludeTentative ? activeStays : activeStays.filter((s) => !isTentative(s.status));

    const daysInMonthCount = daysInMonthOf(currentMonth);
    const monthStartStr = monthStartOf(currentMonth);
    const monthEndStr = monthEndOf(currentMonth);

    const monthStats = rangeStats(kpiStays, monthStartStr, monthEndStr);
    const recordCount = Object.values(monthData).reduce((acc, list) => acc + list.length, 0);
    const totalCount = monthStats.stays;
    const roomNights = daysInMonthCount * master.rooms.length;
    // 稼働率は「延べ宿泊数 ÷ 室日」。退所日を含む利用日ベースだと、同日入替のある
    // 月で 100% を超えてしまうため。
    const utilizationRate = roomNights ? Math.round((monthStats.bedNights / roomNights) * 1000) / 10 : 0;

    const monthEvents = [];
    Object.keys(monthData).forEach((dateStr) => (monthData[dateStr] || []).forEach((r) => monthEvents.push(r)));
    // 1滞在＝入所+退所の2レコードなので、レコード数で数えると件数が倍になる。滞在単位で数える。
    const pickupCount = activeStays.filter((s) => {
      const pk = pickupByLabel[s.pickup];
      return pk && pk.highlight && stayDays(s, monthStartStr, monthEndStr).length > 0;
    }).length;
    const cancelledCount = monthEvents.filter((r) => isCancelled(r.status) && r.type === '入所').length;
    const tentativeCount = activeStays.filter((s) => isTentative(s.status) && stayDays(s, monthStartStr, monthEndStr).length).length;

    const reportRange = this.state.reportRange;
    const rangeValid = !!reportRange.start && !!reportRange.end && reportRange.end >= reportRange.start;
    const customRangeStats = rangeValid ? rangeStats(kpiStays, reportRange.start, reportRange.end) : { days: 0, users: 0, bedNights: 0, stays: 0 };
    const isCurrentActualMonth = currentMonth === THIS_MONTH;
    const cumEndDay = isCurrentActualMonth ? NOW.getDate() : daysInMonthCount;
    const cumEndStr = currentMonth + '-' + pad2(Math.min(cumEndDay, daysInMonthCount));
    const cumulative = rangeStats(kpiStays, monthStartStr, cumEndStr);

    // 当月に1日以上かかる滞在を「1回」として利用者別に集計する。
    const userUsageStats = {};
    kpiStays.forEach((s) => {
      const days = stayDays(s, monthStartStr, monthEndStr);
      if (!days.length) return;
      const row = (userUsageStats[s.name] = userUsageStats[s.name] || { count: 0, days: 0 });
      row.count += 1;
      row.days += days.length;
    });
    const userUsageRows = Object.keys(userUsageStats)
      .map((name) => ({ name, count: userUsageStats[name].count, days: userUsageStats[name].days }))
      .sort((a, b) => b.days - a.days || b.count - a.count || a.name.localeCompare(b.name, 'ja'));
    const userUsageTotalCount = userUsageRows.reduce((sum, row) => sum + row.count, 0);
    const userUsageTotalDays = userUsageRows.reduce((sum, row) => sum + row.days, 0);

    const userKpi = { label: '利用者数', value: monthStats.users, unit: '名', sub: currentMonthLabel, color: 'oklch(0.5 0.14 250)' };
    const kpiCards = [
      { label: '稼働日数', value: monthStats.days, unit: '日', sub: '/ ' + daysInMonthCount + '日中', color: 'oklch(0.55 0.13 150)' },
      { label: '稼働率（宿泊ベース）', value: utilizationRate, unit: '%', sub: '延べ' + monthStats.bedNights + '泊 ÷ ' + roomNights + '室日', color: primaryAccent },
      { label: '送迎あり', value: pickupCount, unit: '件', sub: '仮予約 ' + tentativeCount + '件 / キャンセル ' + cancelledCount + '件', color: 'oklch(0.6 0.13 70)' },
    ];

    // ── 月グリッド ──────────────────────────────────────────────
    const [year, monthNum] = currentMonth.split('-').map(Number);
    const firstOfMonth = new Date(year, monthNum - 1, 1);
    const startWeekday = firstOfMonth.getDay();
    const gridStart = new Date(year, monthNum - 1, 1 - startWeekday);
    const gridStartStr = fmtDate(gridStart);
    const gridEnd = new Date(gridStart); gridEnd.setDate(gridStart.getDate() + 41);
    const occupancy = occupancyMap(activeStays, gridStartStr, fmtDate(gridEnd));

    const statusStyleOf = (status) => {
      if (isCancelled(status)) return { dashed: true, faded: true };
      if (isTentative(status)) return { dashed: true, faded: false };
      return { dashed: false, faded: false };
    };

    const calendarCells = [];
    for (let i = 0; i < 42; i++) {
      const d = new Date(gridStart);
      d.setDate(gridStart.getDate() + i);
      const dateStr = fmtDate(d);
      const inMonth = d.getMonth() === monthNum - 1;
      const weekday = d.getDay();
      const isToday = dateStr === TODAY;
      const buildChip = (r) => {
        const cat = categoryByLabel[r.category] || { color: '#666666' };
        const cancelled = isCancelled(r.status);
        const st = statusStyleOf(r.status);
        const catColor = cancelled ? 'oklch(0.58 0.01 75)' : cat.color;
        const catBg = cancelled ? 'oklch(0.94 0.005 75)' : 'color-mix(in oklab, ' + catColor + ' 12%, white)';
        const pk = pickupByLabel[r.pickup];
        const detail = [r.notes, r.pickupMethod, r.pickupLocation, r.externalService].filter(Boolean).join(' / ');
        return {
          id: r.id,
          typeLabel: cancelled ? '取消 ' + r.type : ((r.status && r.status !== '確定') ? r.status + ' ' + r.type : r.type),
          time: r.time || '時間未定',
          name: r.name,
          color: catColor,
          bg: catBg,
          // 仮予約は破線、確定は実線。遠目に「確定床」と「仮床」が見分けられる。
          borderStyle: st.dashed ? 'dashed' : 'solid',
          showPickup: !!pk && pk.highlight,
          pickupLabel: r.pickup,
          pickupBg: pickupHighlight ? 'oklch(0.75 0.13 70)' : 'oklch(0.94 0.02 70)',
          pickupColor: pickupHighlight ? 'oklch(0.32 0.1 70)' : 'oklch(0.5 0.05 70)',
          hasNotes: !!detail,
          // 開かなくても読めるように title 属性でツールチップを出す。
          tooltip: detail || (r.name + ' ' + r.type),
          onEdit: () => this.openEdit(dateStr, r.id),
        };
      };
      const dayRecords = recordsOn(dateStr);
      const occByRoom = occupancy[dateStr] || {};
      const multiRoom = master.rooms.length > 1;
      const lanes = master.rooms.map((room) => {
        const roomRecords = dayRecords
          .filter((r) => laneIdFor(r) === room.id)
          .sort((a, b) => {
            // 同日入替は、退所を上・入所を下に固定して流れを読みやすくする。
            const typeOrderA = a.type === '退所' ? 0 : 1;
            const typeOrderB = b.type === '退所' ? 0 : 1;
            if (typeOrderA !== typeOrderB) return typeOrderA - typeOrderB;
            return String(a.time || '').localeCompare(String(b.time || ''));
          });
        const reservations = roomRecords.map(buildChip);
        // Days in the middle of a stay carry no record of their own; without
        // this the room reads as vacant and invites a double booking.
        const stayingChips = (occByRoom[room.id] || [])
          .filter((s) => !roomRecords.some((r) => r.name === s.name))
          .map((s) => {
            const cat = categoryByLabel[s.category] || { color: '#666666' };
            const meals = mealsOfStay(s)[dateStr] || { b: false, l: false, d: false };
            return {
              name: s.name,
              color: cat.color,
              bg: 'color-mix(in oklab, ' + cat.color + ' 7%, white)',
              borderStyle: isTentative(s.status) ? 'dotted' : 'dashed',
              mealLabel: [meals.b ? '朝' : '', meals.l ? '昼' : '', meals.d ? '夕' : ''].filter(Boolean).join('・') || '食事なし',
              tooltip: s.name + ' / ' + s.category + ' / ' + s.start + '〜' + (s.end || '退所未定'),
              onOpen: () => this.openEditStay(s),
            };
          });
        const occupied = reservations.length > 0 || stayingChips.length > 0;
        return {
          roomId: room.id,
          roomName: room.name,
          showLabel: multiRoom,
          reservations,
          stayingChips,
          hasReservations: occupied,
          isEmpty: !occupied,
          hasConflict: !!conflicts.dateRoom[dateStr + '|' + room.id],
          onAdd: () => this.openAdd(dateStr, room.id),
        };
      });
      const dayMeals = { b: 0, l: 0, d: 0 };
      (Object.keys(occByRoom)).forEach((rid) => (occByRoom[rid] || []).forEach((s) => {
        const m = mealsOfStay(s)[dateStr];
        if (!m) return;
        if (m.b) dayMeals.b++;
        if (m.l) dayMeals.l++;
        if (m.d) dayMeals.d++;
      }));
      calendarCells.push({
        key: dateStr + '-' + i,
        dateStr,
        dayNum: d.getDate(),
        inMonth,
        isToday,
        todayBadge: isToday ? '今日' : '',
        cellBg: isToday ? 'color-mix(in oklab, ' + primaryAccent + ' 6%, white)' : (inMonth ? 'oklch(1 0 0)' : 'oklch(0.96 0.005 75)'),
        cellBorder: isToday ? primaryAccent : 'oklch(0.91 0.008 75)',
        cellBorderWidth: isToday ? '2px' : '1px',
        cellOpacity: inMonth ? 1 : 0.6,
        // 狭い画面では月グリッドを1列に落とす。そのとき前後月のセルまで並べると
        // 42行になって読めないので、当月以外は隠す。
        outClass: inMonth ? '' : 'ss-cell--out',
        dateColor: !inMonth ? 'oklch(0.7 0.005 75)' : weekday === 0 ? 'oklch(0.55 0.15 25)' : weekday === 6 ? 'oklch(0.5 0.13 250)' : 'oklch(0.3 0.01 75)',
        hasMeals: dayMeals.b + dayMeals.l + dayMeals.d > 0,
        mealSummary: '朝' + dayMeals.b + '・昼' + dayMeals.l + '・夕' + dayMeals.d,
        lanes,
      });
    }

    const weekdayHeaders = WEEKDAY_LABELS.map((label, i) => ({
      label,
      color: i === 0 ? 'oklch(0.55 0.15 25)' : i === 6 ? 'oklch(0.5 0.13 250)' : 'oklch(0.5 0.01 75)',
    }));

    // ── タイムライン（部屋 × 日） ────────────────────────────────
    const preview = this.state.dragPreview;
    const timelineDays = [];
    for (let i = 0; i < daysInMonthCount; i++) {
      const ds = shiftDate(monthStartStr, i);
      const wd = parseDate(ds).getDay();
      timelineDays.push({
        dateStr: ds,
        dayNum: i + 1,
        weekday: WEEKDAY_LABELS[wd],
        isToday: ds === TODAY,
        color: wd === 0 ? 'oklch(0.55 0.15 25)' : wd === 6 ? 'oklch(0.5 0.13 250)' : 'oklch(0.5 0.01 75)',
        bg: ds === TODAY ? 'color-mix(in oklab, ' + primaryAccent + ' 14%, white)' : (wd === 0 || wd === 6 ? 'oklch(0.975 0.006 75)' : 'transparent'),
      });
    }
    const timelineWidth = (daysInMonthCount * DAY_W) + 'px';
    const timelineRows = master.rooms.map((room) => {
      const bars = activeStays
        .filter((s) => s.room === room.id)
        .map((s) => {
          const usesPreview = preview && preview.stayId && preview.stayId === s.stayId;
          const start = usesPreview ? preview.start : s.start;
          const end = usesPreview ? preview.end : (s.end || monthEndStr);
          if (end < monthStartStr || start > monthEndStr) return null;
          const clampedStart = start < monthStartStr ? monthStartStr : start;
          const clampedEnd = end > monthEndStr ? monthEndStr : end;
          const offset = daysBetween(monthStartStr, clampedStart);
          const span = daysBetween(clampedStart, clampedEnd) + 1;
          const cat = categoryByLabel[s.category] || { color: '#666666' };
          const nights = stayNights(s, RANGE_MIN, RANGE_MAX).length;
          const openBar = () => { if (!this.suppressClick) this.openEditStay(s); };
          const noop = () => {};
          return {
            key: (s.stayId || s.name + s.start) + '-' + room.id,
            left: (offset * DAY_W) + 'px',
            width: (span * DAY_W - 4) + 'px',
            bg: 'color-mix(in oklab, ' + cat.color + ' ' + (isTentative(s.status) ? '10' : '20') + '%, white)',
            border: cat.color,
            borderStyle: isTentative(s.status) ? 'dashed' : 'solid',
            color: cat.color,
            label: s.name,
            sub: (s.startTimeRaw || '--:--') + '→' + (s.endTimeRaw || '--:--') + (isTentative(s.status) ? ' ' + s.status : ''),
            tooltip: s.name + ' / ' + s.category + ' / ' + s.start + '〜' + (s.end || '退所未定') + ' / ' + nights + '泊'
              + (s.notes ? ' / ' + s.notes : ''),
            truncStart: start < monthStartStr,
            truncEnd: end > monthEndStr,
            draggable: !!s.stayId,
            onOpen: openBar,
            onMoveDown: s.stayId ? this.startBarDrag('move', s) : noop,
            onStartDown: s.stayId ? this.startBarDrag('start', s) : noop,
            onEndDown: s.stayId ? this.startBarDrag('end', s) : noop,
          };
        })
        .filter(Boolean)
        .sort((a, b) => parseFloat(a.left) - parseFloat(b.left));
      const previewBar = preview && preview.mode === 'create' && preview.roomId === room.id
        ? {
          left: (daysBetween(monthStartStr, preview.start) * DAY_W) + 'px',
          width: ((daysBetween(preview.start, preview.end) + 1) * DAY_W - 4) + 'px',
          label: preview.start.slice(5).replace('-', '/') + '〜' + preview.end.slice(5).replace('-', '/'),
        }
        : null;
      return {
        roomId: room.id,
        roomName: room.name,
        gridStart: monthStartStr,
        bars,
        hasPreview: !!previewBar,
        preview: previewBar || { left: '0px', width: '0px', label: '' },
        onTrackDown: this.startTrackDrag(room.id, monthStartStr),
      };
    });

    // ── 当日ビュー ────────────────────────────────────────────────
    const focusDate = this.state.focusDate;
    const focusStays = activeStays.filter((s) => stayDays(s, focusDate, focusDate).length > 0);
    const focusMeals = { b: 0, l: 0, d: 0 };
    focusStays.forEach((s) => {
      const m = mealsOfStay(s)[focusDate];
      if (!m) return;
      if (m.b) focusMeals.b++;
      if (m.l) focusMeals.l++;
      if (m.d) focusMeals.d++;
    });
    const toDayRow = (s) => {
      const cat = categoryByLabel[s.category] || { color: '#666666' };
      const m = mealsOfStay(s)[focusDate] || { b: false, l: false, d: false };
      return {
        name: s.name,
        room: roomNameOf(s.room),
        category: s.category,
        color: cat.color,
        time: s.start === focusDate ? (s.startTimeRaw || '時間未定') : (s.end === focusDate ? (s.endTimeRaw || '時間未定') : ''),
        meals: [m.b ? '朝' : '', m.l ? '昼' : '', m.d ? '夕' : ''].filter(Boolean).join('・') || '食事なし',
        pickup: s.pickup,
        pickupDetail: [s.pickupMethod, s.pickupLocation].filter(Boolean).join(' / '),
        showPickupDetail: !!(s.pickupMethod || s.pickupLocation),
        status: s.status,
        showStatus: s.status !== '確定',
        notes: s.notes,
        showNotes: !!s.notes,
        onOpen: () => this.openEditStay(s),
      };
    };
    const arrivals = focusStays.filter((s) => s.start === focusDate).map(toDayRow);
    const departures = focusStays.filter((s) => s.end === focusDate).map(toDayRow);
    const staying = focusStays.filter((s) => s.start !== focusDate && s.end !== focusDate).map(toDayRow);
    const focusPickups = focusStays
      .filter((s) => { const pk = pickupByLabel[s.pickup]; return pk && pk.highlight; })
      .map(toDayRow);

    // ── 一覧（印刷向け） ──────────────────────────────────────────
    const listRows = activeStays
      .filter((s) => stayDays(s, monthStartStr, monthEndStr).length)
      .sort((a, b) => (a.start + (a.startTime || '')).localeCompare(b.start + (b.startTime || '')))
      .map((s) => {
        const cat = categoryByLabel[s.category] || { color: '#666666' };
        const meals = mealsOfStay(s);
        let b = 0, l = 0, dd = 0;
        stayDays(s, monthStartStr, monthEndStr).forEach((day) => {
          const m = meals[day];
          if (!m) return;
          if (m.b) b++;
          if (m.l) l++;
          if (m.d) dd++;
        });
        return {
          name: s.name,
          category: s.category,
          color: cat.color,
          room: roomNameOf(s.room),
          range: String(s.start).replace(/-/g, '/') + ' ' + (s.startTimeRaw || '時刻未記入') + ' 〜 ' + (s.end ? String(s.end).replace(/-/g, '/') : '退所未定') + ' ' + (s.endTimeRaw || '時刻未記入'),
          days: stayDays(s, monthStartStr, monthEndStr).length,
          nights: stayNights(s, monthStartStr, monthEndStr).length,
          meals: '朝' + b + ' 昼' + l + ' 夕' + dd,
          status: s.status,
          pickup: s.pickup,
          notes: s.notes,
          onOpen: () => this.openEditStay(s),
        };
      });

    // ── 空室照会 ──────────────────────────────────────────────────
    const vac = this.state.vacancyRange;
    const vacancyValid = !!vac.start && !!vac.end && vac.end >= vac.start;
    const vacancyRows = vacancyValid
      ? vacancyFor(activeStays, master.rooms, vac.start, vac.end, master.defaultCheckinTime, master.defaultCheckoutTime)
        .map((v) => ({
          roomName: v.roomName,
          free: v.free,
          busy: !v.free,
          statusLabel: v.free ? '空き' : '予約あり',
          detail: v.free
            ? (daysBetween(vac.start, vac.end)) + '泊で受け入れ可能'
            : v.blockers.map((b) => b.name + '（' + b.range + '）').join(' / '),
          onCreate: () => this.openAdd(vac.start, v.roomId, { startDate: vac.start, endDate: vac.end }),
        }))
      : [];
    const vacancyFreeCount = vacancyRows.filter((v) => v.free).length;

    // ── ドロワー ──────────────────────────────────────────────────
    const drawer = this.state.drawer;
    const drawerOpen = !!drawer;

    const catOpts = master.categories.map((c) => {
      const active = drawer && drawer.category === c.label;
      return {
        label: c.label,
        select: () => this.setDrawerField({ category: c.label }),
        border: active ? c.color : 'oklch(0.87 0.01 75)',
        bg: active ? 'color-mix(in oklab, ' + c.color + ' 12%, white)' : 'oklch(1 0 0)',
        color: active ? c.color : 'oklch(0.5 0.01 75)',
      };
    });

    const legacyMealOpts = [
      { key: 'mealB', label: '朝' },
      { key: 'mealL', label: '昼' },
      { key: 'mealD', label: '夕' },
    ].map((m) => {
      const active = drawer && drawer[m.key];
      return {
        label: m.label,
        toggle: () => this.setDrawerField({ [m.key]: !(drawer && drawer[m.key]) }),
        border: active ? primaryAccent : 'oklch(0.87 0.01 75)',
        bg: active ? 'oklch(0.94 0.03 250)' : 'oklch(1 0 0)',
        color: active ? primaryAccent : 'oklch(0.5 0.01 75)',
      };
    });

    // 日別の食事表。連泊でも初日は夕食から、最終日は朝食まで、が表現できる。
    const mealRows = [];
    const mealTotals = { b: 0, l: 0, d: 0 };
    if (drawer && drawer.isStayEditor && drawer.startDate && drawer.endDate && drawer.endDate >= drawer.startDate) {
      const table = drawer.mealsByDate || {};
      eachDate(drawer.startDate, drawer.endDate, (day) => {
        const m = table[day] || { b: false, l: false, d: false };
        if (m.b) mealTotals.b += 1;
        if (m.l) mealTotals.l += 1;
        if (m.d) mealTotals.d += 1;
        mealRows.push({
          dateStr: day,
          label: dateLabelOf(day),
          cells: [
            { key: 'b', label: '朝', on: !!m.b },
            { key: 'l', label: '昼', on: !!m.l },
            { key: 'd', label: '夕', on: !!m.d },
          ].map((c) => ({
            label: c.label,
            toggle: () => this.toggleDrawerMeal(day, c.key),
            bg: c.on ? 'oklch(0.9 0.06 75)' : 'oklch(0.97 0.003 75)',
            color: c.on ? 'oklch(0.38 0.09 75)' : 'oklch(0.72 0.005 75)',
            border: c.on ? 'oklch(0.72 0.09 75)' : 'oklch(0.9 0.005 75)',
          })),
        });
      });
    }
    const multiRoomDrawer = master.rooms.length > 1;
    const roomOpts = master.rooms.map((room) => {
      const active = drawer && drawer.room === room.id;
      return {
        label: room.name,
        select: () => this.setDrawerField({ room: room.id }),
        border: active ? primaryAccent : 'oklch(0.87 0.01 75)',
        bg: active ? 'color-mix(in oklab, ' + primaryAccent + ' 12%, white)' : 'oklch(1 0 0)',
        color: active ? primaryAccent : 'oklch(0.5 0.01 75)',
      };
    });

    const pickupOpts = master.pickupOptions.map((p) => {
      const active = drawer && drawer.pickup === p.label;
      return {
        label: p.label,
        select: () => this.setDrawerField({ pickup: p.label }),
        border: active ? 'oklch(0.65 0.13 70)' : 'oklch(0.87 0.01 75)',
        bg: active ? 'oklch(0.95 0.05 70)' : 'oklch(1 0 0)',
        color: active ? 'oklch(0.4 0.1 70)' : 'oklch(0.5 0.01 75)',
      };
    });

    const statusOpts = STATUS_LIST.map((label) => {
      const active = drawer && (drawer.status || '確定') === label;
      const danger = label === 'キャンセル';
      return {
        label, select: () => this.setDrawerField({ status: label }),
        border: active ? (danger ? 'oklch(0.65 0.14 25)' : primaryAccent) : 'oklch(0.87 0.01 75)',
        bg: active ? (danger ? 'oklch(0.95 0.05 25)' : 'oklch(0.94 0.03 250)') : 'white',
        color: active ? (danger ? 'oklch(0.5 0.16 25)' : primaryAccent) : 'oklch(0.5 0.01 75)',
      };
    });

    const mealPresets = [
      { label: '標準（初日夕〜最終日朝）', apply: () => this.applyMealStandard() },
      { label: '全日3食', apply: () => this.applyMealPreset({ b: true, l: true, d: true }) },
      { label: '全日食事なし', apply: () => this.applyMealPreset({ b: false, l: false, d: false }) },
    ];

    const viewTabs = [
      { key: 'month', label: '月カレンダー' },
      { key: 'timeline', label: 'タイムライン' },
      { key: 'today', label: '当日の動き' },
      { key: 'list', label: '一覧・印刷' },
    ].map((t) => ({
      label: t.label,
      select: () => this.setView(t.key),
      bg: view === t.key ? primaryAccent : 'white',
      color: view === t.key ? 'white' : 'oklch(0.42 0.01 75)',
      border: view === t.key ? primaryAccent : 'oklch(0.87 0.01 75)',
      weight: view === t.key ? '700' : '500',
    }));

    const drawerNights = drawer && drawer.isStayEditor && drawer.endDate >= drawer.startDate
      ? daysBetween(drawer.startDate, drawer.endDate) : 0;

    return {
      primaryAccent,
      cellMinHeight,
      timelineWidth,
      currentLabel: currentMonthLabel,
      currentMonth,
      totalCount,
      recordCount,
      userKpi,
      kpiCards,
      kpiIncludeTentative: this.state.kpiIncludeTentative,
      kpiScopeLabel: this.state.kpiIncludeTentative ? '仮予約を含む' : '確定のみ',
      toggleKpiTentative: this.toggleKpiTentative,

      isMonthView: view === 'month',
      isTimelineView: view === 'timeline',
      isTodayView: view === 'today',
      isListView: view === 'list',
      viewTabs,

      prevMonth: this.prevMonth,
      nextMonth: this.nextMonth,
      goToday: this.goToday,
      onMonthInputChange: this.onMonthInputChange,

      conflictCount: conflicts.count,
      hasConflicts: conflicts.count > 0,
      conflictMessage: conflicts.count > 0
        ? '同じ部屋で期間が重なっている予約が ' + conflicts.count + ' 組あります。該当セルに「重複」と表示しています。'
        : '',

      userUsageOpen: this.state.userUsageOpen,
      userUsageRows,
      userUsageTotalCount,
      userUsageTotalDays,
      userUsageHasRows: userUsageRows.length > 0,
      userUsageEmpty: userUsageRows.length === 0,
      toggleUserUsage: this.toggleUserUsage,
      userUsageToggleLabel: this.state.userUsageOpen ? 'クリックして閉じる' : 'クリックして内訳を表示',

      vacancyOpen: this.state.vacancyOpen,
      toggleVacancy: this.toggleVacancy,
      vacancyToggleLabel: this.state.vacancyOpen ? '空室照会を閉じる' : '空室照会を開く',
      vacancyStart: vac.start,
      vacancyEnd: vac.end,
      onVacancyStartChange: this.updateVacancyStart,
      onVacancyEndChange: this.updateVacancyEnd,
      vacancyRows,
      vacancyValid,
      vacancyInvalid: !vacancyValid,
      vacancySummary: vacancyValid
        ? vac.start.replace(/-/g, '/') + '〜' + vac.end.replace(/-/g, '/') + '（' + daysBetween(vac.start, vac.end) + '泊）は ' + vacancyFreeCount + ' / ' + master.rooms.length + ' 室が空いています'
        : '開始日が終了日より後になっています。',

      reportRangeStart: reportRange.start,
      reportRangeEnd: reportRange.end,
      onReportStartChange: this.updateReportStart,
      onReportEndChange: this.updateReportEnd,
      setReportThisMonth: this.setReportThisMonth,
      reportRangeInvalid: !rangeValid,
      reportCards: [
        {
          label: '指定期間（' + String(reportRange.start).replace(/-/g, '/') + '〜' + String(reportRange.end).replace(/-/g, '/') + '）',
          users: customRangeStats.users, days: customRangeStats.days, nights: customRangeStats.bedNights,
          bg: 'oklch(0.98 0.005 75)', border: 'oklch(0.85 0.01 75)',
        },
        {
          label: isCurrentActualMonth ? '当月累計（本日 ' + cumEndDay + '日時点）' : '当月累計（' + daysInMonthCount + '日時点）',
          users: cumulative.users, days: cumulative.days, nights: cumulative.bedNights,
          bg: 'color-mix(in oklab, ' + primaryAccent + ' 8%, white)', border: primaryAccent,
        },
      ],
      exportDetailCsv: this.exportDetailCsv,
      exportSummaryCsv: this.exportSummaryCsv,
      exportMealCsv: this.exportMealCsv,

      weekdayHeaders,
      calendarCells,
      timelineDays,
      timelineRows,

      focusDate,
      focusDateLabel: dateLabelOf(focusDate),
      focusIsToday: focusDate === TODAY,
      onFocusDateChange: this.onFocusDateChange,
      focusPrev: () => this.shiftFocusDate(-1),
      focusNext: () => this.shiftFocusDate(1),
      focusToday: () => this.setState({ focusDate: TODAY }),
      arrivals, departures, staying, focusPickups,
      hasArrivals: arrivals.length > 0, noArrivals: arrivals.length === 0,
      hasDepartures: departures.length > 0, noDepartures: departures.length === 0,
      hasStaying: staying.length > 0, noStaying: staying.length === 0,
      hasFocusPickups: focusPickups.length > 0, noFocusPickups: focusPickups.length === 0,
      focusMealB: focusMeals.b, focusMealL: focusMeals.l, focusMealD: focusMeals.d,
      focusMealTotal: focusMeals.b + focusMeals.l + focusMeals.d,
      focusOccupied: focusStays.length,
      focusCapacity: master.rooms.length,

      listRows,
      listEmpty: listRows.length === 0,
      listHasRows: listRows.length > 0,

      drawerOpen,
      drawerDateLabel: drawer ? String(drawer.dateStr).replace(/-/g, '/') : '',
      drawerDateStr: drawer ? drawer.dateStr : TODAY,
      drawerTime: drawer ? drawer.time : '10:00',
      drawerIsStayEditor: drawer ? !!drawer.isStayEditor : false,
      drawerIsLegacyEditor: drawer ? !drawer.isStayEditor : false,
      drawerStartDate: drawer ? drawer.startDate : TODAY,
      drawerStartTime: drawer ? drawer.startTime : master.defaultCheckinTime,
      drawerEndDate: drawer ? drawer.endDate : shiftDate(TODAY, 1),
      drawerEndTime: drawer ? drawer.endTime : master.defaultCheckoutTime,
      drawerNightsLabel: drawerNights + '泊' + (drawerNights + 1) + '日',
      drawerName: drawer ? drawer.name : '',
      drawerNotes: drawer ? drawer.notes : '',
      drawerPickupMethod: drawer ? drawer.pickupMethod : '',
      drawerPickupLocation: drawer ? drawer.pickupLocation : '',
      drawerExternalService: drawer ? drawer.externalService : '',
      drawerDetailsOpen: drawer ? !!drawer.detailsOpen : false,
      drawerDetailsLabel: drawer && drawer.detailsOpen ? '詳細項目を閉じる' : '送迎場所・外部サービスなどの詳細を表示',
      drawerMealsOpen: drawer ? !!drawer.mealsOpen : false,
      drawerMealsLabel: drawer && drawer.mealsOpen ? '日別の食事を閉じる' : '日別の食事を編集',
      drawerMealSummary: '朝' + mealTotals.b + ' ／ 昼' + mealTotals.l + ' ／ 夕' + mealTotals.d,
      mealRows,
      mealPresets,
      drawerIsExisting: drawer ? !drawer.isNew : false,
      drawerError: drawer ? drawer.error : null,
      drawerUpdatedAt: drawer && drawer.updatedAt ? '最終更新 ' + drawer.updatedAt + (drawer.updatedBy ? '（' + drawer.updatedBy + '）' : '') : '',
      drawerHasHistory: !!(drawer && drawer.updatedAt),
      onDateChange: (e) => this.setDrawerField({ dateStr: e.target.value }),
      onTimeChange: (e) => this.setDrawerField({ time: e.target.value }),
      onStartDateChange: (e) => this.setDrawerField({ startDate: e.target.value, dateStr: e.target.value }),
      onStartTimeChange: (e) => this.setDrawerField({ startTime: e.target.value }),
      onEndDateChange: (e) => this.setDrawerField({ endDate: e.target.value }),
      onEndTimeChange: (e) => this.setDrawerField({ endTime: e.target.value }),
      onNameChange: this.onNameChange,
      onNotesChange: (e) => this.setDrawerField({ notes: e.target.value }),
      onPickupMethodChange: (e) => this.setDrawerField({ pickupMethod: e.target.value }),
      onPickupLocationChange: (e) => this.setDrawerField({ pickupLocation: e.target.value }),
      onExternalServiceChange: (e) => this.setDrawerField({ externalService: e.target.value }),
      toggleDrawerDetails: () => this.setDrawerField({ detailsOpen: !(drawer && drawer.detailsOpen) }),
      toggleDrawerMeals: this.toggleDrawerMeals,
      extendStay: () => this.extendStay(1),
      shortenStay: () => this.extendStay(-1),
      applyPrevious: this.applyPrevious,
      duplicateNextWeek: this.duplicateNextWeek,
      setTypeCheckin: () => this.setDrawerField({ type: '入所' }),
      setTypeCheckout: () => this.setDrawerField({ type: '退所' }),
      typeCheckinBorder: drawer && drawer.type === '入所' ? primaryAccent : 'oklch(0.87 0.01 75)',
      typeCheckinBg: drawer && drawer.type === '入所' ? 'oklch(0.94 0.03 250)' : 'oklch(1 0 0)',
      typeCheckinColor: drawer && drawer.type === '入所' ? primaryAccent : 'oklch(0.5 0.01 75)',
      typeCheckoutBorder: drawer && drawer.type === '退所' ? primaryAccent : 'oklch(0.87 0.01 75)',
      typeCheckoutBg: drawer && drawer.type === '退所' ? 'oklch(0.94 0.03 250)' : 'oklch(1 0 0)',
      typeCheckoutColor: drawer && drawer.type === '退所' ? primaryAccent : 'oklch(0.5 0.01 75)',
      categoryOptions: catOpts,
      mealOptions: legacyMealOpts,
      pickupOptions: pickupOpts,
      statusOptions: statusOpts,
      roomOptions: roomOpts,
      multiRoomDrawer,
      printCalendar: () => window.print(),
      closeDrawer: this.closeDrawer,
      saveDrawer: this.saveDrawer,
      deleteDrawer: this.deleteDrawer,

      facilityName: master.facilityName,
      buildingLabel: master.buildingLabel,
      onFacilityNameChange: (e) => this.updateMaster({ facilityName: e.target.value }),
      onBuildingLabelChange: (e) => this.updateMaster({ buildingLabel: e.target.value }),
      capacity: master.rooms.length,
      defaultCheckinTime: master.defaultCheckinTime,
      defaultCheckoutTime: master.defaultCheckoutTime,
      onDefaultCheckinChange: (e) => this.updateMaster({ defaultCheckinTime: e.target.value }),
      onDefaultCheckoutChange: (e) => this.updateMaster({ defaultCheckoutTime: e.target.value }),
      operator: master.operator,
      onOperatorChange: (e) => this.updateMaster({ operator: e.target.value }),

      // FAX帳票の送信元として刷り込む施設連絡先
      facilityPostalCode: master.postalCode,
      facilityAddress: master.address,
      facilityTel: master.tel,
      facilityFax: master.fax,
      facilityContactPerson: master.contactPerson,
      onPostalCodeChange: (e) => this.updateMaster({ postalCode: e.target.value }),
      onAddressChange: (e) => this.updateMaster({ address: e.target.value }),
      onTelChange: (e) => this.updateMaster({ tel: e.target.value }),
      onFaxChange: (e) => this.updateMaster({ fax: e.target.value }),
      onContactPersonChange: (e) => this.updateMaster({ contactPerson: e.target.value }),

      faxRecipientRows: master.faxRecipients.map((r, idx) => ({
        idx, name: r.name, contact: r.contact, fax: r.fax, tel: r.tel, note: r.note,
        title: r.name || '（事業所名未入力）',
        onNameChange: (e) => this.updateFaxRecipient(idx, { name: e.target.value }),
        onContactChange: (e) => this.updateFaxRecipient(idx, { contact: e.target.value }),
        onFaxChange: (e) => this.updateFaxRecipient(idx, { fax: e.target.value }),
        onTelChange: (e) => this.updateFaxRecipient(idx, { tel: e.target.value }),
        onNoteChange: (e) => this.updateFaxRecipient(idx, { note: e.target.value }),
        onRemove: () => this.removeFaxRecipient(idx),
      })),
      addFaxRecipient: this.addFaxRecipient,
      faxRecipientsEmpty: master.faxRecipients.length === 0,

      boilerplateBelongings: master.boilerplate.belongings,
      boilerplateNotices: master.boilerplate.notices,
      onBelongingsChange: (e) => this.updateBoilerplate({ belongings: e.target.value }),
      onNoticesChange: (e) => this.updateBoilerplate({ notices: e.target.value }),

      openSettings: this.openSettings,
      closeSettings: this.closeSettings,
      settingsOpen: this.state.settingsOpen,
      resetAll: this.resetAll,
      exportBackup: this.exportBackup,
      importBackup: this.importBackup,
      markImportMerge: this.markImportMerge,
      restoreAutoBackup: this.restoreAutoBackup,
      storageNotice: this.state.storageNotice,
      resetLabel: '初期データに戻す（確認・自動バックアップあり）',
      resetColor: 'oklch(0.6 0.01 75)',

      userRows: master.users.map((u, idx) => ({
        idx, value: u.name,
        onChange: (e) => this.updateUser(idx, e.target.value),
        onRemove: () => this.removeUser(idx),
      })),
      addUser: this.addUser,

      categoryRows: master.categories.map((c, idx) => ({
        idx, label: c.label, color: c.color,
        onLabelChange: (e) => this.updateCategory(idx, { label: e.target.value }),
        onColorChange: (e) => this.updateCategory(idx, { color: e.target.value }),
        onRemove: () => this.removeCategory(idx),
      })),
      addCategory: this.addCategory,

      pickupRows: master.pickupOptions.map((p, idx) => ({
        idx, label: p.label, highlight: p.highlight,
        onLabelChange: (e) => this.updatePickupOption(idx, { label: e.target.value }),
        onToggleHighlight: () => this.updatePickupOption(idx, { highlight: !p.highlight }),
        onRemove: () => this.removePickupOption(idx),
      })),
      addPickupOption: this.addPickupOption,

      roomRows: master.rooms.map((r, idx) => ({
        idx, name: r.name,
        onChange: (e) => this.updateRoom(idx, e.target.value),
        onRemove: () => this.removeRoom(idx),
      })),
      addRoom: this.addRoom,

      legendCategories: master.categories,
      userNameOptions: master.users.map((u) => u.name),
    };
  }
}
