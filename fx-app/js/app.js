'use strict';

// ========== Tab Navigation ==========
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const target = btn.dataset.tab;
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tab-' + target).classList.add('active');
  });
});

// ========== LocalStorage Helpers ==========
function loadJSON(key, fallback) {
  try { return JSON.parse(localStorage.getItem(key)) || fallback; } catch { return fallback; }
}

function saveJSON(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
}

// ========== Today's Date ==========
function formatDate(d) {
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日（${'日月火水木金土'[d.getDay()]}）`;
}

const today = new Date();
const todayKey = today.toISOString().slice(0, 10);
document.getElementById('today-date').textContent = formatDate(today);

// Default record date to today
const recDateEl = document.getElementById('rec-date');
if (recDateEl) recDateEl.value = todayKey;

// ========== Daily Checklist ==========
const CHECKLIST_KEYS = ['morning', 'noon', 'evening', 'night', 'after_trade', 'summary'];

function loadChecklist() {
  const saved = loadJSON('checklist_' + todayKey, {});
  CHECKLIST_KEYS.forEach(key => {
    const el = document.querySelector(`input[data-key="${key}"]`);
    if (el) el.checked = !!saved[key];
  });
  updateProgress();
}

function saveChecklist() {
  const data = {};
  CHECKLIST_KEYS.forEach(key => {
    const el = document.querySelector(`input[data-key="${key}"]`);
    if (el) data[key] = el.checked;
  });
  saveJSON('checklist_' + todayKey, data);
  updateProgress();
}

function updateProgress() {
  let done = 0;
  CHECKLIST_KEYS.forEach(key => {
    const el = document.querySelector(`input[data-key="${key}"]`);
    if (el && el.checked) done++;
  });
  const pct = Math.round((done / CHECKLIST_KEYS.length) * 100);
  document.getElementById('checklist-progress').style.width = pct + '%';
  document.getElementById('progress-text').textContent = `${done} / ${CHECKLIST_KEYS.length} 完了`;
}

document.querySelectorAll('.checklist-item input[type="checkbox"]').forEach(el => {
  el.addEventListener('change', saveChecklist);
});

loadChecklist();

// ========== Rules ==========
const RULE_FIELDS = [
  'rule-max-leverage', 'rule-max-loss', 'rule-pair',
  'rule-max-trades', 'rule-require-sl', 'rule-no-trade-indicators', 'rule-capital'
];

function loadRules() {
  const saved = loadJSON('user_rules', {});
  RULE_FIELDS.forEach(id => {
    const el = document.getElementById(id);
    if (el && saved[id] !== undefined) el.value = saved[id];
  });
}

function getRules() {
  const rules = {};
  RULE_FIELDS.forEach(id => {
    const el = document.getElementById(id);
    if (el) rules[id] = el.value;
  });
  return rules;
}

document.getElementById('rules-form').addEventListener('submit', e => {
  e.preventDefault();
  saveJSON('user_rules', getRules());
  const msg = document.getElementById('rules-saved-msg');
  msg.classList.remove('hidden');
  setTimeout(() => msg.classList.add('hidden'), 2000);
});

loadRules();

// ========== Pre-Trade Check ==========
document.getElementById('check-submit').addEventListener('click', () => {
  const rules = loadJSON('user_rules', {});
  const pair = document.getElementById('check-pair').value;
  const directionEl = document.querySelector('input[name="check-direction"]:checked');
  const reason = document.getElementById('check-reason').value.trim();
  const sl = parseFloat(document.getElementById('check-sl').value);
  const lot = parseFloat(document.getElementById('check-lot').value);
  const indicator = document.getElementById('check-indicator').value;
  const emotion = document.getElementById('check-emotion').value;

  const warnings = [];
  let dangerCount = 0;

  // 損切り未設定
  if (!sl || sl <= 0) {
    warnings.push({ level: 'danger', msg: '損切り位置が未設定です。取引前に損切りを決めることはこのアプリの最重要ルールです。' });
    dangerCount++;
  } else {
    warnings.push({ level: 'ok', msg: `損切り ${sl} pips を設定しています。` });
  }

  // 取引理由
  if (!reason || reason.length < 5) {
    warnings.push({ level: 'caution', msg: '取引理由が入力されていません。感情ではなく根拠を言語化してから入ることが大事です。' });
  } else {
    warnings.push({ level: 'ok', msg: '取引理由が記入されています。' });
  }

  // 重要指標前
  if (indicator === 'before') {
    warnings.push({ level: 'danger', msg: '重要指標の発表前です。初心者は指標発表前後は値動きが荒れるため、観察を推奨します。' });
    dangerCount++;
  }

  // 感情チェック
  if (emotion !== 'calm') {
    const emotionLabel = { rush: '焦り', greed: '欲', fear: '怖さ' }[emotion] || emotion;
    warnings.push({ level: 'caution', msg: `感情：「${emotionLabel}」を感じているようです。感情的な取引はルール違反につながりやすいです。一度深呼吸してください。` });
  } else {
    warnings.push({ level: 'ok', msg: '感情：冷静な状態です。' });
  }

  // 通貨ペアチェック
  const allowedPair = rules['rule-pair'];
  if (allowedPair && pair !== allowedPair) {
    warnings.push({ level: 'caution', msg: `あなたのルール設定では ${allowedPair} を取引対象にしています。${pair} は設定外です。` });
  }

  // ロット超過（資金と最大損失%からざっくりチェック）
  const capital = parseFloat(rules['rule-capital']) || 0;
  const maxLossPct = parseFloat(rules['rule-max-loss']) || 2;
  if (capital > 0 && sl > 0 && lot > 0) {
    // ドル円想定：1万通貨×1pips = 100円
    const slLoss = lot * 10000 * (sl * 0.01);
    const lossPct = (slLoss / capital) * 100;
    if (lossPct > maxLossPct) {
      warnings.push({ level: 'danger', msg: `損切り時の損失が資金の ${lossPct.toFixed(1)}% です。あなたのルール（${maxLossPct}%）を超えています。ロットを減らしてください。` });
      dangerCount++;
    } else {
      warnings.push({ level: 'ok', msg: `損失率 ${lossPct.toFixed(1)}%：ルール内（${maxLossPct}%以内）です。` });
    }
  }

  // レバレッジチェック
  const maxLev = parseFloat(rules['rule-max-leverage']) || 3;
  if (capital > 0 && lot > 0 && document.getElementById('calc-price').value) {
    const price = parseFloat(document.getElementById('calc-price').value);
    const tradeAmount = lot * 10000 * price;
    const leverage = tradeAmount / capital;
    if (leverage > maxLev) {
      warnings.push({ level: 'danger', msg: `実質レバレッジが ${leverage.toFixed(1)} 倍です。あなたのルール（${maxLev}倍）を超えています。` });
      dangerCount++;
    }
  }

  // 損切り必須ルール
  if (rules['rule-require-sl'] === 'yes' && (!sl || sl <= 0)) {
    // already covered above
  }

  // 表示
  const warnEl = document.getElementById('check-warnings');
  warnEl.innerHTML = warnings.map(w =>
    `<div class="warning-item ${w.level}">${w.msg}</div>`
  ).join('');

  const verdictEl = document.getElementById('check-verdict');
  if (dangerCount === 0) {
    verdictEl.innerHTML = '<div class="verdict-ok">✓ ルール上の重大な問題は見当たりません。ただし取引の推奨ではありません。</div>';
  } else if (dangerCount === 1) {
    verdictEl.innerHTML = '<div class="verdict-caution">⚠ 注意点が1つあります。もう一度確認してから判断してください。</div>';
  } else {
    verdictEl.innerHTML = '<div class="verdict-stop">✕ 複数のルール違反があります。今日は観察に徹することを検討してください。</div>';
  }

  document.getElementById('check-result').classList.remove('hidden');
});

// ========== Lot / Leverage Calc ==========
document.getElementById('calc-submit').addEventListener('click', () => {
  const capital = parseFloat(document.getElementById('calc-capital').value);
  const lot = parseFloat(document.getElementById('calc-lot').value);
  const price = parseFloat(document.getElementById('calc-price').value);
  const slPips = parseFloat(document.getElementById('calc-sl-pips').value);

  if (!capital || !lot || !price) {
    alert('資金・ロット・価格を入力してください。');
    return;
  }

  // ドル円前提：1万通貨 = 10,000通貨
  const units = lot * 10000;
  const tradeAmount = units * price;
  const leverage = tradeAmount / capital;
  const perYen = units;  // ドル円：1円動くと units 円の損益
  const slLoss = slPips > 0 ? units * (slPips * 0.01) : null;  // 1pips = 0.01円
  const lossPct = slLoss ? (slLoss / capital) * 100 : null;

  document.getElementById('res-trade-amount').textContent = `¥${tradeAmount.toLocaleString()}`;
  document.getElementById('res-leverage').textContent = `${leverage.toFixed(2)} 倍`;
  document.getElementById('res-per-yen').textContent = `±¥${perYen.toLocaleString()}`;
  document.getElementById('res-sl-loss').textContent = slLoss ? `¥${slLoss.toLocaleString()}` : '—';
  document.getElementById('res-loss-pct').textContent = lossPct ? `${lossPct.toFixed(2)}%` : '—';

  const warnEl = document.getElementById('calc-warnings');
  const warns = [];
  if (leverage > 10) warns.push('<div class="warning-item danger">レバレッジが10倍を超えています。初心者は1〜3倍を推奨します。</div>');
  else if (leverage > 3) warns.push('<div class="warning-item caution">レバレッジが3倍を超えています。初心者の目安は3倍以下です。</div>');
  if (lossPct && lossPct > 5) warns.push('<div class="warning-item danger">損切り時の損失が資金の5%を超えています。ロットを減らすことを検討してください。</div>');
  else if (lossPct && lossPct > 2) warns.push('<div class="warning-item caution">損切り時の損失が資金の2%を超えています。初心者の目安は1〜2%です。</div>');
  warnEl.innerHTML = warns.join('');

  document.getElementById('calc-result').classList.remove('hidden');
});

// ========== Trade Records ==========
function loadRecords() {
  return loadJSON('trade_records', []);
}

function saveRecords(records) {
  saveJSON('trade_records', records);
}

function renderRecords() {
  const records = loadRecords();
  const container = document.getElementById('records-list');
  if (records.length === 0) {
    container.innerHTML = '<p class="no-records">まだ記録がありません。<br>取引後はかならず記録を残しましょう。</p>';
    return;
  }

  const resultLabel = { win: '勝ち', loss: '負け', break_even: '建値', open: '未決済' };
  const emotionLabel = { calm: '冷静', rush: '焦り', greed: '欲', fear: '怖さ' };
  const dirLabel = { buy: '買い', sell: '売り' };

  container.innerHTML = records.slice().reverse().map((r, i) => {
    const idx = records.length - 1 - i;
    return `
    <div class="record-card ${r.result}">
      <div class="record-header">
        <span class="record-date">${r.date}</span>
        <div style="display:flex;gap:8px;align-items:center;">
          <span class="record-badge badge-${r.result}">${resultLabel[r.result] || r.result}</span>
          <button class="delete-btn" data-idx="${idx}" title="削除">✕</button>
        </div>
      </div>
      <div class="record-pair">${r.pair}　${dirLabel[r.direction] || r.direction}</div>
      <div class="record-reason">${r.reason || '（理由なし）'}</div>
      <div class="record-meta">
        <span class="tag ${r.sl_set === 'yes' ? 'green' : 'red'}">損切り: ${r.sl_set === 'yes' ? 'あり' : 'なし'}</span>
        <span class="tag">感情: ${emotionLabel[r.emotion] || r.emotion}</span>
        <span class="tag ${r.followed_rules === 'yes' ? 'green' : 'red'}">ルール: ${r.followed_rules === 'yes' ? '遵守' : '違反'}</span>
      </div>
      ${r.reflection ? `<p style="margin-top:8px;font-size:0.8rem;color:#555;">${r.reflection}</p>` : ''}
    </div>`;
  }).join('');

  // Delete buttons
  container.querySelectorAll('.delete-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      if (!confirm('この記録を削除しますか？')) return;
      const records = loadRecords();
      records.splice(parseInt(btn.dataset.idx), 1);
      saveRecords(records);
      renderRecords();
    });
  });
}

// Record form toggle
document.getElementById('open-record-form').addEventListener('click', () => {
  document.getElementById('record-form-wrap').classList.remove('hidden');
});

document.getElementById('cancel-record').addEventListener('click', () => {
  document.getElementById('record-form-wrap').classList.add('hidden');
});

document.getElementById('save-record').addEventListener('click', () => {
  const directionEl = document.querySelector('input[name="rec-direction"]:checked');
  const record = {
    date: document.getElementById('rec-date').value || todayKey,
    pair: document.getElementById('rec-pair').value,
    direction: directionEl ? directionEl.value : '',
    reason: document.getElementById('rec-reason').value.trim(),
    sl_set: document.getElementById('rec-sl-set').value,
    result: document.getElementById('rec-result').value,
    emotion: document.getElementById('rec-emotion').value,
    followed_rules: document.getElementById('rec-followed').value,
    reflection: document.getElementById('rec-reflection').value.trim(),
    created_at: new Date().toISOString()
  };

  if (!record.direction) { alert('取引方向を選択してください。'); return; }

  const records = loadRecords();
  records.push(record);
  saveRecords(records);
  renderRecords();

  // Reset form
  document.getElementById('rec-reason').value = '';
  document.getElementById('rec-reflection').value = '';
  document.querySelector('input[name="rec-direction"]:checked') && (document.querySelector('input[name="rec-direction"]:checked').checked = false);
  document.getElementById('rec-sl-set').value = 'yes';
  document.getElementById('rec-result').value = 'win';
  document.getElementById('rec-emotion').value = 'calm';
  document.getElementById('rec-followed').value = 'yes';
  document.getElementById('record-form-wrap').classList.add('hidden');
});

renderRecords();
