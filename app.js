const DATA_PATH='./dashboard-data.json';
const $=id=>document.getElementById(id);
const norm=v=>String(v??'').trim();
const displayValue=v=>{const s=norm(v);return s.startsWith('=')?'':s};
const num=v=>{const n=Number(displayValue(v));return Number.isFinite(n)?n:null};
const dateKey=v=>{if(!v)return'';const s=String(v).trim().replace(/[/.]/g,'-');const m=s.match(/(20\d{2})-(\d{1,2})-(\d{1,2})/);return m?`${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`:''};
const taipeiToday=()=>new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Taipei',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
const shiftDate=(d,delta)=>{const x=new Date(`${d}T00:00:00Z`);x.setUTCDate(x.getUTCDate()+delta);return x.toISOString().slice(0,10)};
const clamp=(v,min,max)=>Math.min(max,Math.max(min,v));
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

function renderMeals(rows){
  const box=$('meals');box.innerHTML='';
  if(!rows.length){box.innerHTML='<div class="row"><div class="left"><b>沒有飲食紀錄</b><p>這一天尚未記錄餐點。</p></div></div>';$('mealCount').textContent='0 筆';return}
  rows.forEach(r=>{
    const meal=firstField(r,['餐別','餐次'])||'飲食';
    const food=firstField(r,['餐點內容','餐點','內容','食物'])||'未命名餐點';
    const p=firstField(r,['蛋白質']);
    const c=firstField(r,['熱量','卡路里']);
    const source=firstField(r,['餐廳／來源','來源']);
    const note=firstField(r,['備註']);
    const meta=[source,note].filter(Boolean).join(' · ')||'—';
    box.insertAdjacentHTML('beforeend',`<div class="row"><div class="left"><b>${esc(meal)} · ${esc(food)}</b><p>${esc(meta)}</p></div><div class="value">${p?`${esc(p)} g`:''}${p&&c?'<br>':''}${c?`${esc(c)} kcal`:''}</div></div>`);
  });
  $('mealCount').textContent=`${rows.length} 筆`;
}

function renderChecks(r,workouts){
  const box=$('checks');box.innerHTML='';
  if(!r){box.innerHTML='<div class="row"><div class="left"><b>沒有每日檢核</b><p>這一天尚未填寫。</p></div></div>';$('checkState').textContent='待補';return}
  const items=[['零下注',['零下注']],['查看盤口',['查看盤口']],['衝動分數',['衝動 0–10','衝動分數','衝動']],['完成重訓',['完成重訓']],['散步分鐘',['散步分鐘']],['睡眠小時',['睡眠小時']]];
  let filled=0;
  items.forEach(([label,names])=>{
    let v=firstField(r,names);
    if(label==='完成重訓'&&v===''){
      if(workouts.length)v='是';
      else if(hasFormulaField(r,'完成重訓')&&firstField(r,['零下注'])!=='')v='否';
    }
    if(v!=='')filled++;
    box.insertAdjacentHTML('beforeend',`<div class="check"><span>${label}</span><b>${esc(v||'—')}</b></div>`);
  });
  $('checkState').textContent=`${filled}/${items.length}`;
}

function renderWorkouts(rows){
  const box=$('workouts');box.innerHTML='';
  if(!rows.length){box.innerHTML='<div class="row"><div class="left"><b>沒有重訓紀錄</b><p>這一天沒有已記錄的訓練組。</p></div></div>';$('workoutState').textContent='0 組';return}
  rows.forEach(r=>{
    const ex=firstField(r,['動作','項目'])||'動作';
    const set=firstField(r,['組次','組']);
    const w=firstField(r,['重量']);
    const reps=firstField(r,['次數']);
    const rpe=firstField(r,['RPE']);
    box.insertAdjacentHTML('beforeend',`<div class="row"><div class="left"><b>${esc(ex)}${set?` · 第 ${esc(set)} 組`:''}</b><p>${rpe?`RPE ${esc(rpe)}`:'RPE 未記錄'}</p></div><div class="value">${w?`${esc(w)} kg`:''}${reps?` × ${esc(reps)}`:''}</div></div>`);
  });
  $('workoutState').textContent=`${rows.length} 組`;
}

function renderTrend(selected){
  const grouped=new Map();
  state.foodRows.forEach(r=>{const d=dateKey(r['日期']||r.Date);if(!d)return;const g=grouped.get(d)||{p:0,c:0};g.p+=num(firstField(r,['蛋白質']))||0;g.c+=num(firstField(r,['熱量','卡路里']))||0;grouped.set(d,g)});
  const days=Array.from({length:7},(_,i)=>shiftDate(selected,i-6));
  const points=days.map(d=>({date:d,...(grouped.get(d)||{p:0,c:0})}));
  const maxP=Math.max(1,...points.map(x=>x.p));
  const maxC=Math.max(1,...points.map(x=>x.c));
  $('trendBars').innerHTML=points.map(x=>{
    const ph=x.p?clamp(x.p/maxP*100,4,100):2;
    const ch=x.c?clamp(x.c/maxC*100,4,100):2;
    const label=`${Number(x.date.slice(5,7))}/${Number(x.date.slice(8,10))}`;
    return `<div class="trend-day${x.date===selected?' selected':''}" title="${esc(x.date)} · 蛋白質 ${x.p}g · 熱量 ${x.c}kcal"><div class="trend-columns"><div class="trend-bar protein" style="height:${ph}%"></div><div class="trend-bar calorie" style="height:${ch}%"></div></div><div class="trend-label">${label}</div></div>`;
  }).join('');
}

function render(date){
  state.selected=date;
  $('datePicker').value=date;
  $('selectedDateLabel').textContent=formatLongDate(date);

  const meals=rowsForDate(state.foodRows,date);
  const checks=rowsForDate(state.checkRows,date);
  const workouts=rowsForDate(state.workRows,date);
  const check=checks[0];

  setMetric('protein',sumField(meals,['蛋白質'])||0);
  setMetric('calories',sumField(meals,['熱量','卡路里'])||0);
  setMetric('sleep',check?firstField(check,['睡眠小時']):'');
  setMetric('walk',check?firstField(check,['散步分鐘']):'');
  setMetric('zeroBet',check?firstField(check,['零下注']):'');
  setMetric('urge',check?firstField(check,['衝動 0–10','衝動分數','衝動']):'');

  renderMeals(meals);renderChecks(check,workouts);renderWorkouts(workouts);renderTrend(date);

  const day=Math.max(1,Math.min(30,Math.floor((Date.parse(`${date}T00:00:00Z`)-Date.parse(`${state.startDate}T00:00:00Z`))/86400000)+1));
  const pct=Math.round(day/30*100);
  $('dayLabel').textContent=`Day ${day} / 30`;
  $('progressPct').textContent=`${pct}%`;
  $('progressBar').style.width=`${pct}%`;

  const hasData=meals.length||workouts.length||hasMeaningfulCheckData(check);
  $('statusBanner').textContent=hasData?'這一天的資料已載入。':'這一天目前沒有已填寫的紀錄。';
  $('statusBanner').classList.remove('bad');
  $('prevDay').disabled=date<=state.minDate;
  $('nextDay').disabled=date>=state.maxDate;
}

function bindNavigation(){
  $('prevDay').addEventListener('click',()=>render(shiftDate(state.selected,-1)));
  $('nextDay').addEventListener('click',()=>render(shiftDate(state.selected,1)));
  $('datePicker').addEventListener('change',e=>{if(e.target.value)render(e.target.value)});
  $('todayBtn').addEventListener('click',()=>{const today=taipeiToday();render(today<state.minDate?state.minDate:today>state.maxDate?state.maxDate:today)});
}

async function main(){
  try{
    const res=await fetch(`${DATA_PATH}?v=${Date.now()}`);
    if(!res.ok)throw new Error(`HTTP ${res.status}`);
    state.data=await res.json();
    const foodName=findSheet(state.data,['飲食']);
    const checkName=findSheet(state.data,['每日檢核','檢核']);
    const workName=findSheet(state.data,['重訓']);
    state.foodRows=foodName?state.data.sheets[foodName].rows:[];
    state.checkRows=checkName?state.data.sheets[checkName].rows:[];
    state.workRows=workName?state.data.sheets[workName].rows:[];

    const checkDates=[...new Set(dateList(state.checkRows))].sort();
    const allDates=[...new Set([...dateList(state.foodRows),...checkDates,...dateList(state.workRows)])].sort();
    const today=taipeiToday();
    state.startDate=checkDates[0]||allDates[0]||today;
    state.minDate=checkDates[0]||allDates[0]||today;
    state.maxDate=checkDates.at(-1)||allDates.at(-1)||today;
    $('datePicker').min=state.minDate;$('datePicker').max=state.maxDate;
    bindNavigation();
    render(today<state.minDate?state.minDate:today>state.maxDate?state.maxDate:today);
    $('updatedAt').textContent=`更新 ${new Date(state.data.generated_at).toLocaleString('zh-TW',{timeZone:'Asia/Taipei',month:'numeric',day:'numeric',hour:'2-digit',minute:'2-digit'})}`;
  }catch(e){
    console.error(e);
    $('statusBanner').textContent=`讀取資料失敗：${e.message}`;
    $('statusBanner').classList.add('bad');
  }
}
main();
