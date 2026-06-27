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

// 通貨ペア切替時にUI更新
document.getElementById('calc-pair').addEventListener('change', function() {
  const isEURUSD = this.value === 'EURUSD';
  const isJPYpair = this.value === 'USDJPY' || this.value === 'GBPJPY';
  const usdRateWrap = document.getElementById('calc-usdjpy-rate-wrap');
  const priceHint = document.getElementById('calc-price-hint');
  const pipsHint = document.getElementById('calc-pips-hint');
  const priceLabel = document.getElementById('calc-price-label');

  if (isEURUSD) {
    usdRateWrap.style.display = '';
    priceLabel.textContent = '現在の価格（ドル）';
    priceHint.textContent = 'EURUSDの場合：例 1.0850（ドル）';
    pipsHint.textContent = 'EURUSD：1pips = 0.0001ドル';
  } else {
    usdRateWrap.style.display = 'none';
    priceLabel.textContent = '現在の価格（円）';
    priceHint.textContent = this.value === 'GBPJPY'
      ? 'ポンド円の場合：例 190.00（円）'
      : 'ドル円の場合：例 150.00（円）';
    pipsHint.textContent = 'ドル円・ポンド円：1pips = 0.01円';
  }
});

document.getElementById('calc-submit').addEventListener('click', () => {
  const pair = document.getElementById('calc-pair').value;
  const capital = parseFloat(document.getElementById('calc-capital').value);
  const lot = parseFloat(document.getElementById('calc-lot').value);
  const price = parseFloat(document.getElementById('calc-price').value);
  const slPips = parseFloat(document.getElementById('calc-sl-pips').value);

  if (!capital || !lot || !price) {
    alert('資金・ロット・価格を入力してください。');
    return;
  }

  const units = lot * 10000;
  let tradeAmountJPY, perUnitJPY, slLossJPY, perUnitLabel;

  if (pair === 'USDJPY' || pair === 'GBPJPY') {
    // 円建てペア：価格が円レート
    // 1pips = 0.01円
    tradeAmountJPY = units * price;
    perUnitJPY = units;           // 1円動いたときの損益（円）
    slLossJPY = slPips > 0 ? units * slPips * 0.01 : null;
    perUnitLabel = '1円動いたときの損益';
  } else {
    // EURUSD：価格がドルレート、損益はドル建て → 円換算にドル円レートが必要
    const usdJpy = parseFloat(document.getElementById('calc-usdjpy-rate').value);
    if (!usdJpy) {
      alert('EURUSDはドル円レートも入力してください。');
      return;
    }
    tradeAmountJPY = units * price * usdJpy;
    // 1pips = 0.0001ドル → 円換算
    const perPipUSD = units * 0.0001;
    const perPipJPY = perPipUSD * usdJpy;
    perUnitJPY = perPipJPY;
    slLossJPY = slPips > 0 ? perPipJPY * slPips : null;
    perUnitLabel = '1pips動いたときの損益（円換算）';
  }

  const leverage = tradeAmountJPY / capital;
  const lossPct = slLossJPY ? (slLossJPY / capital) * 100 : null;

  document.getElementById('res-trade-amount').textContent = `¥${Math.round(tradeAmountJPY).toLocaleString()}`;
  document.getElementById('res-leverage').textContent = `${leverage.toFixed(2)} 倍`;
  document.getElementById('res-per-unit-label').textContent = perUnitLabel;
  document.getElementById('res-per-yen').textContent = `±¥${Math.round(perUnitJPY).toLocaleString()}`;
  document.getElementById('res-sl-loss').textContent = slLossJPY ? `¥${Math.round(slLossJPY).toLocaleString()}` : '—';
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

// ========== AI Helper ==========
const AI_SYSTEM_PROMPT = `あなたはFX初心者の学習補助AIです。
以下の制約を必ず守ってください：
・売買推奨・価格予想・利益保証は一切しません
・ニュース整理・リスク確認・取引記録の振り返りのみ行います
・損切り未設定・過大ロット・感情的取引には具体的に警告します
・無理な取引より観察を優先する助言をします
・出力は日本語のみ、300字以内で簡潔に
・回答の末尾に必ず「※これは学習補助コメントです。売買の推奨ではありません。」を付けてください`;

async function callClaudeAPI(userMessage) {
  const apiKey = localStorage.getItem('claude_api_key');
  if (!apiKey) {
    throw new Error('APIキーが設定されていません。「AI振り返り」タブでAPIキーを保存してください。');
  }

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-calls': 'true'
    },
    body: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 512,
      system: AI_SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userMessage }]
    })
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `APIエラー: ${res.status}`);
  }

  const data = await res.json();
  return data.content[0].text;
}

// ========== API Key Management ==========
const apiKeyInput = document.getElementById('api-key-input');
const savedKey = localStorage.getItem('claude_api_key');
if (savedKey) apiKeyInput.value = savedKey;

document.getElementById('toggle-key-visibility').addEventListener('click', function() {
  const isPassword = apiKeyInput.type === 'password';
  apiKeyInput.type = isPassword ? 'text' : 'password';
  this.textContent = isPassword ? '隠す' : '表示';
});

document.getElementById('save-api-key').addEventListener('click', () => {
  const key = apiKeyInput.value.trim();
  if (!key) { alert('APIキーを入力してください。'); return; }
  localStorage.setItem('claude_api_key', key);
  const msg = document.getElementById('apikey-saved');
  msg.classList.remove('hidden');
  setTimeout(() => msg.classList.add('hidden'), 2000);
});

// ========== Pre-trade AI Comment ==========
document.getElementById('check-ai-btn').addEventListener('click', async function() {
  const warnings = document.getElementById('check-warnings').innerText;
  const verdict = document.getElementById('check-verdict').innerText;
  const reason = document.getElementById('check-reason').value.trim();
  const emotion = document.getElementById('check-emotion').value;
  const sl = document.getElementById('check-sl').value;
  const lot = document.getElementById('check-lot').value;
  const indicator = document.getElementById('check-indicator').value;

  const emotionLabel = { calm: '冷静', rush: '焦り', greed: '欲', fear: '怖さ' };
  const indicatorLabel = { none: 'なし', before: 'あり（発表前）', after: 'あり（発表済み）' };

  const msg = `取引前チェック結果を振り返ってください。
取引理由: ${reason || '未入力'}
損切り: ${sl ? sl + ' pips' : '未設定'}
ロット: ${lot ? lot + ' 万通貨' : '未入力'}
今の感情: ${emotionLabel[emotion]}
重要指標: ${indicatorLabel[indicator]}
ルール照合結果: ${warnings}
総合判定: ${verdict}

初心者への学習的なコメントをしてください。売買推奨はしないでください。`;

  this.disabled = true;
  this.textContent = '送信中...';
  const resultEl = document.getElementById('check-ai-result');
  resultEl.classList.remove('hidden');
  resultEl.textContent = '分析中...';

  try {
    const reply = await callClaudeAPI(msg);
    resultEl.textContent = reply;
  } catch (e) {
    resultEl.textContent = 'エラー: ' + e.message;
  } finally {
    this.disabled = false;
    this.textContent = 'AIにコメントをもらう';
  }
});

// ========== Review Tab ==========
function getWeeklyStats(records) {
  const now = new Date();
  const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  const weekly = records.filter(r => new Date(r.date) >= weekAgo);

  const total = weekly.length;
  const wins = weekly.filter(r => r.result === 'win').length;
  const followed = weekly.filter(r => r.followed_rules === 'yes').length;
  const slSet = weekly.filter(r => r.sl_set === 'yes').length;
  const emotional = weekly.filter(r => r.emotion !== 'calm').length;

  return { total, wins, followed, slSet, emotional };
}

function renderWeeklyStats() {
  const records = loadRecords();
  const s = getWeeklyStats(records);
  const el = document.getElementById('weekly-stats');
  if (s.total === 0) {
    el.innerHTML = '<p style="color:#aaa;font-size:0.85rem;">直近7日間の取引記録がありません。</p>';
    return;
  }
  const followRate = s.total > 0 ? Math.round((s.followed / s.total) * 100) : 0;
  const slRate = s.total > 0 ? Math.round((s.slSet / s.total) * 100) : 0;
  el.innerHTML = `
    <div class="stat-grid">
      <div class="stat-item"><div class="stat-value">${s.total}</div><div class="stat-label">取引回数</div></div>
      <div class="stat-item highlight"><div class="stat-value">${followRate}%</div><div class="stat-label">ルール遵守率</div></div>
      <div class="stat-item"><div class="stat-value">${slRate}%</div><div class="stat-label">損切り設定率</div></div>
      <div class="stat-item"><div class="stat-value">${s.emotional}</div><div class="stat-label">感情的取引回数</div></div>
    </div>`;
}

// AI Review
document.getElementById('review-submit').addEventListener('click', async function() {
  const period = document.getElementById('review-period').value;
  const allRecords = loadRecords();
  const now = new Date();

  let targetRecords;
  if (period === 'today') {
    targetRecords = allRecords.filter(r => r.date === todayKey);
  } else if (period === 'week') {
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    targetRecords = allRecords.filter(r => new Date(r.date) >= weekAgo);
  } else {
    targetRecords = allRecords;
  }

  if (targetRecords.length === 0) {
    alert('対象期間に取引記録がありません。');
    return;
  }

  const resultLabel = { win: '勝ち', loss: '負け', break_even: '建値', open: '未決済' };
  const emotionLabel = { calm: '冷静', rush: '焦り', greed: '欲', fear: '怖さ' };
  const dirLabel = { buy: '買い', sell: '売り' };

  const recordText = targetRecords.map(r =>
    `[${r.date}] ${r.pair} ${dirLabel[r.direction] || ''} / 理由: ${r.reason || 'なし'} / 損切り: ${r.sl_set === 'yes' ? 'あり' : 'なし'} / 感情: ${emotionLabel[r.emotion] || ''} / 結果: ${resultLabel[r.result] || ''} / ルール: ${r.followed_rules === 'yes' ? '遵守' : '違反'} / メモ: ${r.reflection || 'なし'}`
  ).join('\n');

  const msg = `以下のFX取引記録を振り返ってください。

${recordText}

以下の観点で学習的なコメントをしてください：
1. 守れたルール
2. 守れなかったルール・改善点
3. 感情的な取引パターンがあればその指摘
4. 次回取引で気をつけること（3点以内）

売買推奨・価格予想はしないでください。`;

  this.disabled = true;
  const resultEl = document.getElementById('review-result');
  const loadingEl = document.getElementById('review-loading');
  const contentEl = document.getElementById('review-content');
  resultEl.classList.remove('hidden');
  loadingEl.classList.remove('hidden');
  contentEl.textContent = '';

  try {
    const reply = await callClaudeAPI(msg);
    contentEl.textContent = reply;
  } catch (e) {
    contentEl.textContent = 'エラー: ' + e.message;
  } finally {
    loadingEl.classList.add('hidden');
    this.disabled = false;
  }
});

// Render weekly stats when switching to review tab
document.querySelectorAll('.tab-btn').forEach(btn => {
  if (btn.dataset.tab === 'review') {
    btn.addEventListener('click', renderWeeklyStats);
  }
});
