const DATA_PATH='./dashboard-data.json';
const $=id=>document.getElementById(id);
const norm=v=>String(v??'').trim();
const displayValue=v=>{const s=norm(v);return s.startsWith('=')?'':s};
const num=v=>{const n=Number(displayValue(v));return Number.isFinite(n)?n:null};
const dateKey=v=>{if(!v)return'';const s=String(v).trim().replace(/[/.]/g,'-');const m=s.match(/(20\d{2})-(\d{1,2})-(\d{1,2})/);return m?`${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`:''};
const taipeiToday=()=>new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Taipei',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
const shiftDate=(d,delta)=>{const x=new Date(`${d}T00:00:00Z`);x.setUTCDate(x.getUTCDate()+delta);return x.toISOString().slice(0,10)};
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const state={data:null,foodRows:[],checkRows:[],workRows:[],selected:'',minDate:'',maxDate:'',startDate:''};
function findSheet(data,keys){return Object.keys(data.sheets||{}).find(n=>keys.some(k=>n.includes(k)))}
function firstField(r,names){for(const n of names){const k=Object.keys(r||{}).find(x=>x.trim()===n||x.includes(n));if(k){const value=displayValue(r[k]);if(value!=='')return value}}return''}
function rowsForDate(rows,date){return rows.filter(r=>dateKey(r['日期']||r.Date)===date)}
function sumField(rows,names){return rows.reduce((sum,r)=>sum+(num(firstField(r,names))??0),0)}
function setMetric(id,v,fallback='—'){$(id).textContent=v===null||v===undefined||v===''?fallback:v}
function dateList(rows){return rows.map(r=>dateKey(r['日期']||r.Date)).filter(Boolean)}
function formatLongDate(date){return new Intl.DateTimeFormat('zh-TW',{timeZone:'UTC',month:'long',day:'numeric',weekday:'long'}).format(new Date(`${date}T00:00:00Z`))}
function hasFormulaField(r,name){const key=Object.keys(r||{}).find(x=>x.includes(name));return Boolean(key&&norm(r[key]).startsWith('='))}
function hasMeaningfulCheckData(r){return Boolean(r&&Object.entries(r).some(([k,v])=>!k.includes('日期')&&displayValue(v)!==''))}
function niceMax(value){if(value<=0)return 1;const p=10**Math.floor(Math.log10(value));const n=value/p;return(n<=1?1:n<=2?2:n<=5?5:10)*p}

function renderMeals(rows){
  const box=$('meals');box.innerHTML='';
  if(!rows.length){box.innerHTML='<div class="row"><div class="left"><b>沒有飲食紀錄</b><p>這一天尚未記錄食物。</p></div></div>';$('mealCount').textContent='0 筆';return}
  const body=rows.map(r=>{
    const food=firstField(r,['餐點內容','餐點','內容','食物'])||'—';
    const protein=firstField(r,['蛋白質']);
    const carbs=firstField(r,['碳水化合物','碳水']);
    return `<tr><td>${esc(food)}</td><td>${protein!==''?`${esc(protein)} g`:'—'}</td><td>${carbs!==''?`${esc(carbs)} g`:'—'}</td></tr>`;
  }).join('');
  box.innerHTML=`<div class="table-wrap"><table class="data-table food-table"><thead><tr><th>食物</th><th>蛋白質</th><th>碳水</th></tr></thead><tbody>${body}</tbody></table></div>`;
  $('mealCount').textContent=`${rows.length} 筆`;
}

function renderChecks(r,workouts){
  const box=$('checks');box.innerHTML='';
  if(!r){box.innerHTML='<div class="row"><div class="left"><b>沒有每日檢核</b><p>這一天尚未填寫。</p></div></div>';$('checkState').textContent='待補';return}
  const items=[['零下注',['零下注']],['查看盤口',['查看盤口']],['衝動分數',['衝動 0–10','衝動分數','衝動']],['完成重訓',['完成重訓']],['散步分鐘',['散步分鐘']],['睡眠小時',['睡眠小時']]];
  let filled=0;
  items.forEach(([label,names])=>{let v=firstField(r,names);if(label==='完成重訓'&&v===''){if(workouts.length)v='是';else if(hasFormulaField(r,'完成重訓')&&firstField(r,['零下注'])!=='')v='否'}if(v!=='')filled++;box.insertAdjacentHTML('beforeend',`<div class="check"><span>${label}</span><b>${esc(v||'—')}</b></div>`)});
  $('checkState').textContent=`${filled}/${items.length}`;
}

function renderWorkouts(rows){
  const box=$('workouts');box.innerHTML='';
  if(!rows.length){box.innerHTML='<div class="row"><div class="left"><b>沒有重訓紀錄</b><p>這一天沒有已記錄的訓練組。</p></div></div>';$('workoutState').textContent='0 組';return}
  const body=rows.map(r=>{const ex=firstField(r,['動作','項目'])||'—';const set=firstField(r,['組次','組'])||'—';const w=firstField(r,['重量']);const reps=firstField(r,['次數']);const rpe=firstField(r,['RPE']);return `<tr><td>${esc(ex)}</td><td>${esc(set)}</td><td>${w!==''?`${esc(w)} kg`:'—'}</td><td>${reps!==''?esc(reps):'—'}</td><td>${rpe!==''?esc(rpe):'—'}</td></tr>`}).join('');
  box.innerHTML=`<div class="table-wrap"><table class="data-table workout-table"><thead><tr><th>動作</th><th>組次</th><th>重量</th><th>次數</th><th>RPE</th></tr></thead><tbody>${body}</tbody></table></div>`;
  $('workoutState').textContent=`${rows.length} 組`;
}

function renderTrend(selected){
  const grouped=new Map();
  state.foodRows.forEach(r=>{const d=dateKey(r['日期']||r.Date);if(!d)return;const g=grouped.get(d)||{p:0,c:0};g.p+=num(firstField(r,['蛋白質']))||0;g.c+=num(firstField(r,['熱量','卡路里']))||0;grouped.set(d,g)});
  const days=Array.from({length:7},(_,i)=>shiftDate(selected,i-6));
  const points=days.map(d=>({date:d,...(grouped.get(d)||{p:0,c:0})}));
  const pMax=niceMax(Math.max(...points.map(x=>x.p),1));
  const cMax=niceMax(Math.max(...points.map(x=>x.c),1));
  const W=680,H=250,L=54,R=58,T=22,B=42,IW=W-L-R,IH=H-T-B;
  const x=i=>L+(points.length===1?0:i/(points.length-1))*IW;
  const yp=v=>T+IH-(v/pMax)*IH;
  const yc=v=>T+IH-(v/cMax)*IH;
  let grid='';
  for(let i=0;i<=4;i++){const y=T+IH-(i/4)*IH;grid+=`<line class="chart-grid" x1="${L}" y1="${y}" x2="${W-R}" y2="${y}"/><text class="axis-text" x="${L-8}" y="${y+4}" text-anchor="end">${Math.round(pMax*i/4)}</text><text class="axis-text" x="${W-R+8}" y="${y+4}" text-anchor="start">${Math.round(cMax*i/4)}</text>`}
  const pLine=points.map((v,i)=>`${x(i)},${yp(v.p)}`).join(' ');
  const cLine=points.map((v,i)=>`${x(i)},${yc(v.c)}`).join(' ');
  const xLabels=points.map((v,i)=>`<text class="axis-text${v.date===selected?' axis-selected':''}" x="${x(i)}" y="${H-16}" text-anchor="middle">${Number(v.date.slice(5,7))}/${Number(v.date.slice(8,10))}</text>`).join('');
  const pDots=points.map((v,i)=>`<circle class="series-dot protein-series" cx="${x(i)}" cy="${yp(v.p)}" r="3.5"/>`).join('');
  const cDots=points.map((v,i)=>`<circle class="series-dot calorie-series" cx="${x(i)}" cy="${yc(v.c)}" r="3.5"/>`).join('');
  const hover=points.map((v,i)=>{const left=i===0?L:(x(i-1)+x(i))/2;const right=i===points.length-1?W-R:(x(i)+x(i+1))/2;return `<rect class="chart-hover-target" data-i="${i}" x="${left}" y="${T}" width="${right-left}" height="${IH}"/>`}).join('');
  const box=$('trendBars');
  box.innerHTML=`<div class="trend-chart-wrap"><svg class="trend-svg" viewBox="0 0 ${W} ${H}" role="img" aria-label="最近七天蛋白質與熱量趨勢圖">${grid}<line class="chart-axis" x1="${L}" y1="${T}" x2="${L}" y2="${T+IH}"/><line class="chart-axis" x1="${W-R}" y1="${T}" x2="${W-R}" y2="${T+IH}"/><line class="chart-axis" x1="${L}" y1="${T+IH}" x2="${W-R}" y2="${T+IH}"/><text class="axis-title protein-axis" x="${L}" y="12">蛋白質 (g)</text><text class="axis-title calorie-axis" x="${W-R}" y="12" text-anchor="end">熱量 (kcal)</text><polyline class="series-line protein-series" points="${pLine}"/>${pDots}<polyline class="series-line calorie-series" points="${cLine}"/>${cDots}${xLabels}${hover}</svg><div class="chart-tooltip" hidden></div></div>`;
  const tooltip=box.querySelector('.chart-tooltip');
  const wrap=box.querySelector('.trend-chart-wrap');
  box.querySelectorAll('.chart-hover-target').forEach(target=>{
    const move=e=>{const rect=wrap.getBoundingClientRect();let left=e.clientX-rect.left+12;let top=e.clientY-rect.top-18;tooltip.style.left=`${Math.min(left,rect.width-130)}px`;tooltip.style.top=`${Math.max(6,top)}px`};
    const show=e=>{const p=points[Number(target.dataset.i)];tooltip.innerHTML=`<strong>${Number(p.date.slice(5,7))}/${Number(p.date.slice(8,10))}</strong><span>蛋白質 ${p.p} g</span><span>熱量 ${p.c} kcal</span>`;tooltip.hidden=false;move(e)};
    target.addEventListener('pointerenter',show);target.addEventListener('pointermove',move);target.addEventListener('pointerleave',()=>{tooltip.hidden=true});
  });
}

function render(date){
  state.selected=date;$('datePicker').value=date;$('selectedDateLabel').textContent=formatLongDate(date);
  const meals=rowsForDate(state.foodRows,date),checks=rowsForDate(state.checkRows,date),workouts=rowsForDate(state.workRows,date),check=checks[0];
  setMetric('protein',sumField(meals,['蛋白質'])||0);setMetric('calories',sumField(meals,['熱量','卡路里'])||0);setMetric('sleep',check?firstField(check,['睡眠小時']):'');setMetric('walk',check?firstField(check,['散步分鐘']):'');setMetric('zeroBet',check?firstField(check,['零下注']):'');setMetric('urge',check?firstField(check,['衝動 0–10','衝動分數','衝動']):'');
  renderMeals(meals);renderChecks(check,workouts);renderWorkouts(workouts);renderTrend(date);
  const day=Math.max(1,Math.min(30,Math.floor((Date.parse(`${date}T00:00:00Z`)-Date.parse(`${state.startDate}T00:00:00Z`))/86400000)+1));const pct=Math.round(day/30*100);$('dayLabel').textContent=`Day ${day} / 30`;$('progressPct').textContent=`${pct}%`;$('progressBar').style.width=`${pct}%`;
  const hasData=meals.length||workouts.length||hasMeaningfulCheckData(check);$('statusBanner').textContent=hasData?'這一天的資料已載入。':'這一天目前沒有已填寫的紀錄。';$('statusBanner').classList.remove('bad');$('prevDay').disabled=date<=state.minDate;$('nextDay').disabled=date>=state.maxDate;
}
function bindNavigation(){$('prevDay').addEventListener('click',()=>render(shiftDate(state.selected,-1)));$('nextDay').addEventListener('click',()=>render(shiftDate(state.selected,1)));$('datePicker').addEventListener('change',e=>{if(e.target.value)render(e.target.value)});$('todayBtn').addEventListener('click',()=>{const today=taipeiToday();render(today<state.minDate?state.minDate:today>state.maxDate?state.maxDate:today)})}
async function main(){try{const res=await fetch(`${DATA_PATH}?v=${Date.now()}`);if(!res.ok)throw new Error(`HTTP ${res.status}`);state.data=await res.json();const foodName=findSheet(state.data,['飲食']),checkName=findSheet(state.data,['每日檢核','檢核']),workName=findSheet(state.data,['重訓']);state.foodRows=foodName?state.data.sheets[foodName].rows:[];state.checkRows=checkName?state.data.sheets[checkName].rows:[];state.workRows=workName?state.data.sheets[workName].rows:[];const checkDates=[...new Set(dateList(state.checkRows))].sort();const allDates=[...new Set([...dateList(state.foodRows),...checkDates,...dateList(state.workRows)])].sort();const today=taipeiToday();state.startDate=checkDates[0]||allDates[0]||today;state.minDate=checkDates[0]||allDates[0]||today;state.maxDate=checkDates.at(-1)||allDates.at(-1)||today;$('datePicker').min=state.minDate;$('datePicker').max=state.maxDate;bindNavigation();render(today<state.minDate?state.minDate:today>state.maxDate?state.maxDate:today);$('updatedAt').textContent=`更新 ${new Date(state.data.generated_at).toLocaleString('zh-TW',{timeZone:'Asia/Taipei',month:'numeric',day:'numeric',hour:'2-digit',minute:'2-digit'})}`}catch(e){console.error(e);$('statusBanner').textContent=`讀取資料失敗：${e.message}`;$('statusBanner').classList.add('bad')}}main();
