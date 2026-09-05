const DATA_PATH = './dashboard-data.json';
const $ = id => document.getElementById(id);
const norm = value => String(value ?? '').trim();
const displayValue = value => {
  const text = norm(value);
  return text.startsWith('=') ? '' : text;
};
const num = value => {
  const parsed = Number(displayValue(value));
  return Number.isFinite(parsed) ? parsed : null;
};
const dateKey = value => {
  if (!value) return '';
  const text = String(value).trim().replace(/[/.]/g, '-');
  const match = text.match(/(20\d{2})-(\d{1,2})-(\d{1,2})/);
  return match ? `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}` : '';
};
const taipeiToday = () => new Intl.DateTimeFormat('en-CA', {
  timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit',
}).format(new Date());
const shiftDate = (date, delta) => {
  const value = new Date(`${date}T00:00:00Z`);
  value.setUTCDate(value.getUTCDate() + delta);
  return value.toISOString().slice(0, 10);
};
const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({
  '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
}[char]));

const state = {
  data: null,
  foodRows: [],
  checkRows: [],
  workRows: [],
  selected: '',
  minDate: '',
  maxDate: '',
  startDate: '',
};

function findSheet(data, keywords) {
  return Object.keys(data.sheets || {}).find(name => keywords.some(keyword => name.includes(keyword)));
}

function firstField(row, names) {
  for (const name of names) {
    const key = Object.keys(row || {}).find(candidate => candidate.trim() === name || candidate.includes(name));
    if (!key) continue;
    const value = displayValue(row[key]);
    if (value !== '') return value;
  }
  return '';
}

function rowsForDate(rows, date) {
  return rows.filter(row => dateKey(row['日期'] || row.Date) === date);
}

function sumField(rows, names) {
  return rows.reduce((sum, row) => sum + (num(firstField(row, names)) ?? 0), 0);
}

function setMetric(id, value, fallback = '—') {
  $(id).textContent = value === null || value === undefined || value === '' ? fallback : value;
}

function dateList(rows) {
  return rows.map(row => dateKey(row['日期'] || row.Date)).filter(Boolean);
}

function formatLongDate(date) {
  return new Intl.DateTimeFormat('zh-TW', {
    timeZone: 'UTC', month: 'long', day: 'numeric', weekday: 'long',
  }).format(new Date(`${date}T00:00:00Z`));
}

function hasFormulaField(row, name) {
  const key = Object.keys(row || {}).find(candidate => candidate.includes(name));
  return Boolean(key && norm(row[key]).startsWith('='));
}

function hasMeaningfulCheckData(row) {
  return Boolean(row && Object.entries(row).some(([key, value]) => !key.includes('日期') && displayValue(value) !== ''));
}

function niceMax(value) {
  if (value <= 0) return 1;
  const power = 10 ** Math.floor(Math.log10(value));
  const normalized = value / power;
  return (normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10) * power;
}

function renderMeals(rows) {
  const box = $('meals');
  box.innerHTML = '';
  if (!rows.length) {
    box.innerHTML = '<div class="row"><div class="left"><b>沒有飲食紀錄</b><p>這一天尚未記錄食物。</p></div></div>';
    $('mealCount').textContent = '0 筆';
    return;
  }

  const body = rows.map(row => {
    const food = firstField(row, ['餐點內容', '餐點', '內容', '食物']) || '—';
    const protein = firstField(row, ['蛋白質']);
    const carbs = firstField(row, ['碳水化合物', '碳水']);
    return `<tr><td>${esc(food)}</td><td>${protein !== '' ? `${esc(protein)} g` : '—'}</td><td>${carbs !== '' ? `${esc(carbs)} g` : '—'}</td></tr>`;
  }).join('');

  box.innerHTML = `<div class="table-wrap"><table class="data-table food-table"><thead><tr><th>食物</th><th>蛋白質</th><th>碳水</th></tr></thead><tbody>${body}</tbody></table></div>`;
  $('mealCount').textContent = `${rows.length} 筆`;
}

function renderChecks(row, workouts) {
  const box = $('checks');
  box.innerHTML = '';
  if (!row) {
    box.innerHTML = '<div class="row"><div class="left"><b>沒有每日檢核</b><p>這一天尚未填寫。</p></div></div>';
    $('checkState').textContent = '待補';
    return;
  }

  const items = [
    ['零下注', ['零下注']],
    ['查看盤口', ['查看盤口']],
    ['衝動分數', ['衝動 0–10', '衝動分數', '衝動']],
    ['完成重訓', ['完成重訓']],
    ['散步分鐘', ['散步分鐘']],
    ['睡眠小時', ['睡眠小時']],
  ];

  let filled = 0;
  items.forEach(([label, names]) => {
    let value = firstField(row, names);
    if (label === '完成重訓' && value === '') {
      if (workouts.length) value = '是';
      else if (hasFormulaField(row, '完成重訓') && firstField(row, ['零下注']) !== '') value = '否';
    }
    if (value !== '') filled += 1;
    box.insertAdjacentHTML('beforeend', `<div class="check"><span>${label}</span><b>${esc(value || '—')}</b></div>`);
  });
  $('checkState').textContent = `${filled}/${items.length}`;
}

function renderWorkouts(rows) {
  const box = $('workouts');
  box.innerHTML = '';
  if (!rows.length) {
    box.innerHTML = '<div class="row"><div class="left"><b>沒有重訓紀錄</b><p>這一天沒有已記錄的訓練組。</p></div></div>';
    $('workoutState').textContent = '0 組';
    return;
  }

  const body = rows.map(row => {
    const exercise = firstField(row, ['動作', '項目']) || '—';
    const set = firstField(row, ['組次', '組']) || '—';
    const weight = firstField(row, ['重量']);
    const reps = firstField(row, ['次數']);
    const rpe = firstField(row, ['RPE']);
    return `<tr><td>${esc(exercise)}</td><td>${esc(set)}</td><td>${weight !== '' ? `${esc(weight)} kg` : '—'}</td><td>${reps !== '' ? esc(reps) : '—'}</td><td>${rpe !== '' ? esc(rpe) : '—'}</td></tr>`;
  }).join('');

  box.innerHTML = `<div class="table-wrap"><table class="data-table workout-table"><thead><tr><th>動作</th><th>組次</th><th>重量</th><th>次數</th><th>RPE</th></tr></thead><tbody>${body}</tbody></table></div>`;
  $('workoutState').textContent = `${rows.length} 組`;
}

function renderTrend(selected) {
  const grouped = new Map();
  state.foodRows.forEach(row => {
    const date = dateKey(row['日期'] || row.Date);
    if (!date) return;
    const current = grouped.get(date) || { p: 0, c: 0 };
    current.p += num(firstField(row, ['蛋白質'])) || 0;
    current.c += num(firstField(row, ['熱量', '卡路里'])) || 0;
    grouped.set(date, current);
  });

  const days = Array.from({ length: 7 }, (_, index) => shiftDate(selected, index - 6));
  const points = days.map(date => ({ date, ...(grouped.get(date) || { p: 0, c: 0 }) }));
  const proteinMax = niceMax(Math.max(...points.map(point => point.p), 1));
  const calorieMax = niceMax(Math.max(...points.map(point => point.c), 1));

  const W = 680;
  const H = 250;
  const L = 54;
  const R = 58;
  const T = 22;
  const B = 42;
  const IW = W - L - R;
  const IH = H - T - B;
  const x = index => L + (points.length === 1 ? 0 : index / (points.length - 1)) * IW;
  const proteinY = value => T + IH - (value / proteinMax) * IH;
  const calorieY = value => T + IH - (value / calorieMax) * IH;

  let grid = '';
  for (let index = 0; index <= 4; index += 1) {
    const y = T + IH - (index / 4) * IH;
    grid += `<line class="chart-grid" x1="${L}" y1="${y}" x2="${W - R}" y2="${y}"/><text class="axis-text" x="${L - 8}" y="${y + 4}" text-anchor="end">${Math.round(proteinMax * index / 4)}</text><text class="axis-text" x="${W - R + 8}" y="${y + 4}" text-anchor="start">${Math.round(calorieMax * index / 4)}</text>`;
  }

  const proteinLine = points.map((point, index) => `${x(index)},${proteinY(point.p)}`).join(' ');
  const calorieLine = points.map((point, index) => `${x(index)},${calorieY(point.c)}`).join(' ');
  const xLabels = points.map((point, index) => `<text class="axis-text${point.date === selected ? ' axis-selected' : ''}" x="${x(index)}" y="${H - 16}" text-anchor="middle">${Number(point.date.slice(5, 7))}/${Number(point.date.slice(8, 10))}</text>`).join('');
  const proteinDots = points.map((point, index) => `<circle class="series-dot protein-series" cx="${x(index)}" cy="${proteinY(point.p)}" r="3.5"/>`).join('');
  const calorieDots = points.map((point, index) => `<circle class="series-dot calorie-series" cx="${x(index)}" cy="${calorieY(point.c)}" r="3.5"/>`).join('');
  const hoverTargets = points.map((point, index) => {
    const left = index === 0 ? L : (x(index - 1) + x(index)) / 2;
    const right = index === points.length - 1 ? W - R : (x(index) + x(index + 1)) / 2;
    return `<rect class="chart-hover-target" data-i="${index}" x="${left}" y="${T}" width="${right - left}" height="${IH}"/>`;
  }).join('');

  const box = $('trendBars');
  box.innerHTML = `<div class="trend-chart-wrap"><svg class="trend-svg" viewBox="0 0 ${W} ${H}" role="img" aria-label="最近七天蛋白質與熱量趨勢圖">${grid}<line class="chart-axis" x1="${L}" y1="${T}" x2="${L}" y2="${T + IH}"/><line class="chart-axis" x1="${W - R}" y1="${T}" x2="${W - R}" y2="${T + IH}"/><line class="chart-axis" x1="${L}" y1="${T + IH}" x2="${W - R}" y2="${T + IH}"/><text class="axis-title protein-axis" x="${L}" y="12">蛋白質 (g)</text><text class="axis-title calorie-axis" x="${W - R}" y="12" text-anchor="end">熱量 (kcal)</text><polyline class="series-line protein-series" points="${proteinLine}"/>${proteinDots}<polyline class="series-line calorie-series" points="${calorieLine}"/>${calorieDots}${xLabels}${hoverTargets}</svg><div class="chart-tooltip" hidden></div></div>`;

  const tooltip = box.querySelector('.chart-tooltip');
  const wrap = box.querySelector('.trend-chart-wrap');
  const svg = box.querySelector('.trend-svg');

  function showTooltip(index) {
    const point = points[index];
    tooltip.innerHTML = `<strong>${Number(point.date.slice(5, 7))}/${Number(point.date.slice(8, 10))}</strong><span>蛋白質 ${point.p} g</span><span>熱量 ${point.c} kcal</span>`;
    tooltip.hidden = false;

    const wrapRect = wrap.getBoundingClientRect();
    const svgRect = svg.getBoundingClientRect();
    const anchorX = svgRect.left - wrapRect.left + (x(index) / W) * svgRect.width;
    const anchorYValue = Math.min(proteinY(point.p), calorieY(point.c));
    const anchorY = svgRect.top - wrapRect.top + (anchorYValue / H) * svgRect.height;
    const left = Math.min(
      Math.max(anchorX - tooltip.offsetWidth / 2, 6),
      wrapRect.width - tooltip.offsetWidth - 6,
    );
    const top = Math.max(6, anchorY - tooltip.offsetHeight - 10);

    tooltip.style.left = `${left}px`;
    tooltip.style.top = `${top}px`;
  }

  box.querySelectorAll('.chart-hover-target').forEach(target => {
    target.addEventListener('pointerenter', () => showTooltip(Number(target.dataset.i)));
    target.addEventListener('pointerleave', () => { tooltip.hidden = true; });
  });
}

function render(date) {
  state.selected = date;
  $('datePicker').value = date;
  $('selectedDateLabel').textContent = formatLongDate(date);

  const meals = rowsForDate(state.foodRows, date);
  const checks = rowsForDate(state.checkRows, date);
  const workouts = rowsForDate(state.workRows, date);
  const check = checks[0];

  setMetric('protein', sumField(meals, ['蛋白質']) || 0);
  setMetric('calories', sumField(meals, ['熱量', '卡路里']) || 0);
  setMetric('sleep', check ? firstField(check, ['睡眠小時']) : '');
  setMetric('walk', check ? firstField(check, ['散步分鐘']) : '');
  setMetric('zeroBet', check ? firstField(check, ['零下注']) : '');
  setMetric('urge', check ? firstField(check, ['衝動 0–10', '衝動分數', '衝動']) : '');

  renderMeals(meals);
  renderChecks(check, workouts);
  renderWorkouts(workouts);
  renderTrend(date);

  const day = Math.max(1, Math.min(30, Math.floor((Date.parse(`${date}T00:00:00Z`) - Date.parse(`${state.startDate}T00:00:00Z`)) / 86400000) + 1));
  const percentage = Math.round(day / 30 * 100);
  $('dayLabel').textContent = `Day ${day} / 30`;
  $('progressPct').textContent = `${percentage}%`;
  $('progressBar').style.width = `${percentage}%`;

  const hasData = meals.length || workouts.length || hasMeaningfulCheckData(check);
  $('statusBanner').textContent = hasData ? '這一天的資料已載入。' : '這一天目前沒有已填寫的紀錄。';
  $('statusBanner').classList.remove('bad');
  $('prevDay').disabled = date <= state.minDate;
  $('nextDay').disabled = date >= state.maxDate;
}

function bindNavigation() {
  $('prevDay').addEventListener('click', () => render(shiftDate(state.selected, -1)));
  $('nextDay').addEventListener('click', () => render(shiftDate(state.selected, 1)));
  $('datePicker').addEventListener('change', event => { if (event.target.value) render(event.target.value); });
  $('todayBtn').addEventListener('click', () => {
    const today = taipeiToday();
    render(today < state.minDate ? state.minDate : today > state.maxDate ? state.maxDate : today);
  });
}

async function main() {
  try {
    const response = await fetch(`${DATA_PATH}?v=${Date.now()}`);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    state.data = await response.json();

    const foodName = findSheet(state.data, ['飲食']);
    const checkName = findSheet(state.data, ['每日檢核', '檢核']);
    const workName = findSheet(state.data, ['重訓']);
    state.foodRows = foodName ? state.data.sheets[foodName].rows : [];
    state.checkRows = checkName ? state.data.sheets[checkName].rows : [];
    state.workRows = workName ? state.data.sheets[workName].rows : [];

    const checkDates = [...new Set(dateList(state.checkRows))].sort();
    const allDates = [...new Set([...dateList(state.foodRows), ...checkDates, ...dateList(state.workRows)])].sort();
    const today = taipeiToday();
    state.startDate = checkDates[0] || allDates[0] || today;
    state.minDate = checkDates[0] || allDates[0] || today;
    state.maxDate = allDates.at(-1) || checkDates.at(-1) || today;
    $('datePicker').min = state.minDate;
    $('datePicker').max = state.maxDate;

    bindNavigation();
    render(today < state.minDate ? state.minDate : today > state.maxDate ? state.maxDate : today);
    $('updatedAt').textContent = `更新 ${new Date(state.data.generated_at).toLocaleString('zh-TW', {
      timeZone: 'Asia/Taipei', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit',
    })}`;
  } catch (error) {
    console.error(error);
    $('statusBanner').textContent = `讀取資料失敗：${error.message}`;
    $('statusBanner').classList.add('bad');
  }
}

main();
