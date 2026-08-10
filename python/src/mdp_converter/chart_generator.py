"""JavaScript runtime for safe AST evaluation and chart rendering."""

EVAL_JS = r"""
function esc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));}
function fid(name){return 'f_'+Array.from(name).map(c=>c.charCodeAt(0).toString(36)).join('_');}
function mid(name){return 'm_'+Array.from(name).map(c=>c.charCodeAt(0).toString(36)).join('_');}
function envKey(name){return String(name||'').toLowerCase().replace(/ё/g,'е').replace(/[^0-9a-zа-я]+/g,'');}

function lookupEnv(name, env){
  if(Object.prototype.hasOwnProperty.call(env,name)) return env[name];
  const variants=[name,name.replace(/\s+/g,'_'),name.replace(/_/g,' '),envKey(name)];
  for(const v of variants){ if(Object.prototype.hasOwnProperty.call(env,v)) return env[v]; }
  return 0;
}

function evalAst(node, env){
  if(!node) return NaN;
  switch(node.type){
    case 'num': return Number(node.value);
    case 'var': { const v=lookupEnv(node.name, env); return typeof v==='boolean' ? (v?1:0) : Number(v||0); }
    case 'un':
      if(node.op==='-') return -evalAst(node.arg, env);
      if(node.op==='not') return evalAst(node.arg, env) ? 0 : 1;
      return NaN;
    case 'bin': {
      const a=evalAst(node.left,env), b=evalAst(node.right,env);
      if(node.op==='+') return a+b;
      if(node.op==='-') return a-b;
      if(node.op==='*') return a*b;
      if(node.op==='/') return b? a/b : NaN;
      if(node.op==='^') return Math.pow(a,b);
      return NaN;
    }
    case 'cmp': {
      const a=evalAst(node.left,env), b=evalAst(node.right,env);
      if(node.op==='=='||node.op==='=') return a===b?1:0;
      if(node.op==='<>'||node.op==='!=') return a!==b?1:0;
      if(node.op==='<') return a<b?1:0;
      if(node.op==='<=') return a<=b?1:0;
      if(node.op==='>') return a>b?1:0;
      if(node.op==='>=') return a>=b?1:0;
      return 0;
    }
    case 'logic': {
      const vals=node.args.map(x=>evalAst(x,env));
      if(node.op==='and') return vals.every(Boolean)?1:0;
      if(node.op==='or') return vals.some(Boolean)?1:0;
      return 0;
    }
    case 'func': {
      const args=node.args.map(x=>evalAst(x,env));
      if(node.name==='abs') return Math.abs(args[0]);
      if(node.name==='min') return Math.min(...args);
      if(node.name==='max') return Math.max(...args);
      return NaN;
    }
    case 'if':
      return evalAst(evalAst(node.cond,env)?node.then:node.else, env);
    default: return NaN;
  }
}

function activeAst(node, env){
  if(!node) return null;
  if(node.type==='if'){
    return evalAst(node.cond, env) ? activeAst(node.then, env) : activeAst(node.else, env);
  }
  if(node.type==='un') return {...node,arg:activeAst(node.arg,env)};
  if(node.type==='bin'||node.type==='cmp') return {...node,left:activeAst(node.left,env),right:activeAst(node.right,env)};
  if(node.type==='logic'||node.type==='func') return {...node,args:(node.args||[]).map(x=>activeAst(x,env))};
  return {...node};
}

function formatAst(node){
  if(!node) return '';
  switch(node.type){
    case 'num': {
      const v=Number(node.value);
      return Number.isInteger(v)? String(v) : String(Math.round(v*1000)/1000);
    }
    case 'var': return node.name;
    case 'un': return node.op==='-' ? '-'+formatAst(node.arg) : 'NOT '+formatAst(node.arg);
    case 'bin': {
      const sym = node.op==='*'?'×': node.op==='/'?'÷': node.op;
      return formatAst(node.left)+' '+sym+' '+formatAst(node.right);
    }
    case 'cmp': return formatAst(node.left)+' '+node.op+' '+formatAst(node.right);
    case 'logic': return node.args.map(formatAst).join(node.op==='and'?' AND ':' OR ');
    case 'func': return node.name.toUpperCase()+'('+node.args.map(formatAst).join(', ')+')';
    case 'if': return 'IF(...)';
    default: return '';
  }
}

function isValidResult(v){ return Number.isFinite(v) && Math.abs(v-999)>1e-9; }

function splitCriteria(text){
  const s=String(text||'').trim();
  if(!s) return [];
  return s.split(/(?=(?:^|\s)\d+\)\s*)/).map(x=>x.trim()).filter(Boolean).map((x,i)=>{
    const m=x.match(/^(\d+)\)\s*(.*)$/s);
    const raw=(m?m[2]:x).trim();
    return {n:m?Number(m[1]):i+1, t:raw.replace(/\[\s*(?:пл|pl)\s*\]/gi,'').trim(), planning:/\[\s*(?:пл|pl)\s*\]/i.test(raw)};
  });
}

function modeValues(){
  const env={};
  (DATA.mode_params||[]).forEach(p=>{
    const el=document.getElementById(mid(p.name));
    if(!el){ env[p.name]=0; return; }
    if(p.kind==='bool') env[p.name]=el.checked?1:0;
    else env[p.name]=Number(el.value||0);
  });
  return env;
}

function factorValues(){
  const env=modeValues();
  (DATA.factors||[]).forEach(f=>{
    const el=document.getElementById(fid(f.name));
    const v=Number(el?.value||0);
    env[f.name]=v;
    env[f.name.replace(/\s+/g,'_')]=v;
    env[f.name.replace(/_/g,' ')]=v;
    env[envKey(f.name)]=v;
  });
  return env;
}

function selectedPaSeason(){
  const el=document.getElementById(mid('pa_season'));
  return el?String(el.value||'1'):'1';
}

function activePaGroup(row){
  const variants=row?.mdp_pa_variants||[];
  if(!variants.length){
    return {items:row?.mdp_pa_items||[], crit:row?.crit_mdp_pa||'', raw:row?.mdp_pa||''};
  }
  const season=selectedPaSeason();
  const found=variants.find(v=>String(v.season)===season)||variants[0];
  return {items:found?.mdp_pa_items||[], crit:found?.crit_mdp_pa||'', raw:found?.mdp_pa||''};
}

const FORMULA_VAR_RESERVED=new Set(['if','min','max','abs','and','or','not','true','false']);

function collectVariablesFromAst(node){
  const names=new Set();
  const walk=(current)=>{
    if(!current) return;
    if(current.type==='var'){ names.add(current.name); return; }
    if(current.type==='un') walk(current.arg);
    else if(current.type==='bin'||current.type==='cmp'){ walk(current.left); walk(current.right); }
    else if(current.type==='logic'||current.type==='func') (current.args||[]).forEach(walk);
    else if(current.type==='if'){ walk(current.cond); walk(current.then); walk(current.else); }
  };
  walk(node);
  return names;
}

function extractVariablesFromText(text){
  const names=new Set();
  const re=/[A-Za-zА-Яа-яЁё_][A-Za-zА-Яа-яЁё0-9_]*/g;
  let match;
  while((match=re.exec(String(text||'')))){
    const token=match[0];
    if(!FORMULA_VAR_RESERVED.has(token.toLowerCase())) names.add(token);
  }
  return names;
}

function collectSchemeFormulaVariables(scheme, group){
  const names=new Set();
  (scheme?.rows||[]).forEach(row=>{
    let items=[];
    let rawText='';
    if(group==='mdp_pa'){
      const pa=activePaGroup(row);
      items=pa.items||[];
      rawText=pa.raw||row.mdp_pa||'';
    } else {
      items=row.mdp_items||[];
      rawText=row.mdp||'';
    }
    extractVariablesFromText(rawText).forEach(name=>names.add(name));
    (items||[]).forEach(item=>{
      if(item?.ast) collectVariablesFromAst(item.ast).forEach(name=>names.add(name));
      else extractVariablesFromText(item?.raw||'').forEach(name=>names.add(name));
    });
  });
  return names;
}

function formatCopyFactorDefinitions(usedNames){
  if(!usedNames?.size) return [];
  const usedKeys=new Set([...usedNames].map(envKey));
  const lines=[];
  (DATA.factor_definitions||[]).forEach(def=>{
    const name=String(def?.name||'').trim();
    if(!name) return;
    const defKey=envKey(name);
    let matched=usedKeys.has(defKey);
    if(!matched){
      for(const usedName of usedNames){
        if(usedName===name){ matched=true; break; }
      }
    }
    if(!matched) return;
    const description=String(def?.description||'').trim();
    lines.push(`${name}: ${description}${description.endsWith(';')?'':';'}`);
  });
  return lines;
}

function calcItems(items, env){
  return (items||[]).map(it=>{
    if(!it.is_computable||!it.ast){
      return {...it, display: it.raw, value: NaN, activeAst: null};
    }
    const active=activeAst(it.ast, env);
    const val=evalAst(active, env);
    return {...it, display: formatAst(active), value: val, activeAst: active};
  });
}

function minItem(items){
  const nums=items.filter(x=>isValidResult(x.value));
  if(!nums.length) return null;
  return nums.reduce((a,b)=> b.value<a.value?b:a);
}

function formulaBlock(items, env){
  const calc=calcItems(items, env);
  const min=minItem(calc);
  const showValues=Boolean(document.getElementById('calcToggle')?.checked);
  return calc.map(it=>{
    const isMin=showValues&&min&&it.number===min.number&&isValidResult(it.value);
    if(!it.is_computable){
      return `<div class="formula"><b>${it.number})</b> ${esc(String(it.raw||'').replace(/\[\s*(?:пл|pl)\s*\]/gi,'').trim())}</div>`;
    }
    const value=showValues&&Number.isFinite(it.value)?`<span class="inline-value"> = ${it.value.toFixed(1)} МВт</span>`:'';
    return `<div class="formula ${isMin?'minimum':''}"><b>${it.number})</b> ${esc(it.display)}${value}${isMin?'<span class="min-badge">МДП</span>':''}</div>`;
  }).join('');
}

function withMinimumPrefix(html, rawText){
  if(!html) return html;
  const raw=String(rawText||'');
  if(/минимальное\s+из/i.test(raw)) return html;
  return `<div class="minimum-from">Минимальный из:</div>${html}`;
}

function copyNeedsMinimumPrefix(rawText, itemCount){
  if(itemCount<1) return false;
  return !/минимальное\s+из/i.test(String(rawText||''));
}

function criteriaBlock(items, env, critText, highlightMin=true){
  const calc=calcItems(items, env);
  const min=minItem(calc);
  const showValues=Boolean(document.getElementById('calcToggle')?.checked);
  const crits=splitCriteria(critText);
  if(!crits.length) return textLines(critText);
  return crits.map(c=>{
    const isMin=highlightMin&&showValues&&min&&c.n===min.number&&isValidResult(min.value);
    const item=(items||[]).find(x=>Number(x.number)===Number(c.n));
    const planning=c.planning||/\[\s*(?:пл|pl)\s*\]/i.test(String(item?.raw||''));
    return `<div class="crit ${isMin?'minimum':''}"><b>${c.n})</b> ${esc(c.t)}${planning?'<span class="planning-badge">Для планирования</span>':''}</div>`;
  }).join('');
}

function textLines(s){ return esc(String(s||'')).replace(/(?=(?:^|\s)\d+\)\s*)/g,'<br>'); }
"""

CHART_JS = r"""
let CHART_CACHE=null;
function initChart(){
  if(!OPTIONS.chart) return;
  const rs=document.getElementById('chartScheme');
  rs.innerHTML=DATA.schemes.map((s,i)=>`<option value="${i}">${esc(s.number)}. ${esc(s.name.split('\n')[0])}</option>`).join('');
  const gf=document.getElementById('chartGroup');
  gf.innerHTML='';
  if(DATA.has_mdp) gf.innerHTML+='<option value="mdp">МДП без ПА</option>';
  if(DATA.has_mdp_pa) gf.innerHTML+='<option value="mdp_pa">МДП с ПА</option>';
  document.getElementById('chartFactor').innerHTML=(DATA.factors||[]).map(f=>`<option>${esc(f.name)}</option>`).join('');
  function fillTemps(){
    const s=DATA.schemes[Number(rs.value)||0];
    document.getElementById('chartRow').innerHTML=(s?.rows||[]).map((r,i)=>`<option value="${i}">${esc(r.temperature||('строка '+(i+1)))}</option>`).join('');
  }
  rs.onchange=()=>{fillTemps();drawChart();};
  fillTemps();
  document.querySelectorAll('.chart-controls select,.chart-controls input').forEach(e=>e.addEventListener('input',drawChart));
  const canvas=document.getElementById('chart');
  canvas?.addEventListener('mousemove',onChartMove);
  canvas?.addEventListener('mouseleave',()=>{const tip=document.getElementById('chartTooltip');if(tip)tip.style.display='none';drawChart();});
  drawChart();
}

function drawChart(){
  if(!OPTIONS.chart) return;
  const si=Number(document.getElementById('chartScheme').value||0);
  const ri=Number(document.getElementById('chartRow').value||0);
  const group=document.getElementById('chartGroup').value;
  const factor=document.getElementById('chartFactor').value;
  const x0=Number(document.getElementById('chartMin').value||-100);
  const x1=Number(document.getElementById('chartMax').value||100);
  const step=Math.max(0.1, Number(document.getElementById('chartStep').value||10));
  const row=DATA.schemes[si]?.rows?.[ri];
  if(!row||x1<=x0) return showChartMessage('Укажите корректный диапазон влияющего фактора');
  const pa=group==='mdp_pa'?activePaGroup(row):{items:row.mdp_items||[]};
  const items=pa.items||[];
  const formulas=items.filter(x=>x.is_computable&&x.ast);
  const base=factorValues();
  const series=formulas.map(z=>({name:'Критерий '+z.number, pts:[]}));
  const minPts=[];
  for(let x=x0;x<=x1+1e-9;x+=step){
    const env={...base,[factor]:x,[factor.replace(/\s+/g,'_')]:x,[factor.replace(/_/g,' ')]:x};
    const vals=formulas.map(z=>evalAst(activeAst(z.ast, env), env));
    const valid=vals.map((v,i)=>({i,v})).filter(p=>isValidResult(p.v));
    valid.forEach(p=>series[p.i].pts.push({x,y:p.v}));
    if(valid.length) minPts.push({x, y: Math.min(...valid.map(p=>p.v))});
  }
  series.push({name:'Итоговый МДП (минимум)', pts:minPts, dash:true,minimum:true});
  paintChart(series, x0, x1, factor, formulas, base);
}

function showChartMessage(message){
  const c=document.getElementById('chart'); if(!c)return;
  const ctx=c.getContext('2d'); c.width=1200;c.height=420;ctx.clearRect(0,0,c.width,c.height);
  ctx.fillStyle='#64748b';ctx.font='15px sans-serif';ctx.fillText(message,28,44);CHART_CACHE=null;
}

function paintChart(series, x0, x1, xLabel, formulas, base){
  const c=document.getElementById('chart');
  const ctx=c.getContext('2d');
  const W=c.width=1200, H=c.height=420;
  ctx.clearRect(0,0,W,H);
  const ys=series.flatMap(s=>s.pts.map(p=>p.y));
  if(!ys.length) return showChartMessage('Недостаточно данных для построения графика');
  let ymin=Math.min(...ys), ymax=Math.max(...ys);
  if(ymin===ymax){ymin-=1;ymax+=1;}
  const pad=Math.max(2,(ymax-ymin)*.08);ymin-=pad;ymax+=pad;
  const L=72,R=28,T=25,B=52;
  const X=x=>L+(x-x0)/(x1-x0)*(W-L-R);
  const Y=y=>T+(ymax-y)/(ymax-ymin)*(H-T-B);
  ctx.font='12px sans-serif';ctx.textAlign='left';
  for(let i=0;i<=5;i++){
    const py=T+i*(H-T-B)/5, val=ymax-i*(ymax-ymin)/5;
    ctx.strokeStyle='#e2e8f0';ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(L,py);ctx.lineTo(W-R,py);ctx.stroke();
    ctx.fillStyle='#64748b';ctx.fillText(String(Math.round(val*10)/10),10,py+4);
  }
  for(let i=0;i<=5;i++){
    const px=L+i*(W-L-R)/5, val=x0+i*(x1-x0)/5;
    ctx.strokeStyle='#edf1f6';ctx.beginPath();ctx.moveTo(px,T);ctx.lineTo(px,H-B);ctx.stroke();
    ctx.fillStyle='#64748b';ctx.fillText(String(Math.round(val*10)/10),px-13,H-B+20);
  }
  ctx.strokeStyle='#475569';ctx.lineWidth=1.4;ctx.beginPath();ctx.moveTo(L,T);ctx.lineTo(L,H-B);ctx.lineTo(W-R,H-B);ctx.stroke();
  const palette=['#2563eb','#dc2626','#16a34a','#9333ea','#ea580c','#0891b2','#4f46e5','#be123c','#0f766e'];
  series.forEach((s,i)=>{
    s.color=s.minimum?'#0f172a':palette[i%palette.length];
    ctx.strokeStyle=s.color; ctx.lineWidth=s.dash?3:2.2;
    if(s.dash) ctx.setLineDash([8,6]); else ctx.setLineDash([]);
    ctx.beginPath();
    s.pts.forEach((p,j)=> j?ctx.lineTo(X(p.x),Y(p.y)):ctx.moveTo(X(p.x),Y(p.y)));
    ctx.stroke();
  });
  ctx.setLineDash([]);
  ctx.fillStyle='#334155';
  ctx.fillText(xLabel, L+(W-L-R)/2-35, H-10);
  ctx.save(); ctx.translate(15,T+(H-T-B)/2); ctx.rotate(-Math.PI/2);
  ctx.fillText('МДП, МВт',0,0); ctx.restore();
  const legend=document.getElementById('chartLegend');
  if(legend)legend.innerHTML=series.map(s=>`<span><span class="legend-line" style="background:${s.color}"></span>${esc(s.name)}</span>`).join('');
  CHART_CACHE={series,formulas,base,factor:xLabel,x0,x1,ymin,ymax,L,R,T,B,X,Y,W,H};
}

function onChartMove(event){
  if(!CHART_CACHE)return;
  const c=CHART_CACHE,canvas=document.getElementById('chart'),tip=document.getElementById('chartTooltip');
  if(!canvas||!tip)return;
  const rect=canvas.getBoundingClientRect();
  const mx=(event.clientX-rect.left)*(canvas.width/rect.width),my=(event.clientY-rect.top)*(canvas.height/rect.height);
  if(mx<c.L||mx>c.W-c.R||my<c.T||my>c.H-c.B){tip.style.display='none';drawChart();return;}
  const factorValue=Math.round((c.x0+(mx-c.L)/(c.W-c.L-c.R)*(c.x1-c.x0))*10)/10;
  drawChart(); if(!CHART_CACHE)return;
  const current=CHART_CACHE,ctx=canvas.getContext('2d'),px=current.X(factorValue);
  const env={...current.base,[current.factor]:factorValue,[current.factor.replace(/\s+/g,'_')]:factorValue,[current.factor.replace(/_/g,' ')]:factorValue};
  const values=current.formulas.map((z,i)=>({name:'Критерий '+z.number,color:current.series[i]?.color||'#2563eb',value:evalAst(activeAst(z.ast,env),env)})).filter(x=>isValidResult(x.value));
  if(values.length)values.push({name:'Итоговый МДП',color:'#0f172a',value:Math.min(...values.map(x=>x.value))});
  ctx.save();ctx.strokeStyle='#0f172a';ctx.globalAlpha=.35;ctx.setLineDash([5,4]);ctx.beginPath();ctx.moveTo(px,current.T);ctx.lineTo(px,current.H-current.B);ctx.stroke();ctx.setLineDash([]);ctx.globalAlpha=1;
  values.forEach(v=>{const py=current.Y(v.value);ctx.fillStyle=v.color;ctx.beginPath();ctx.arc(px,py,5,0,Math.PI*2);ctx.fill();});ctx.restore();
  tip.innerHTML=`<b>${esc(current.factor)} = ${factorValue}</b>`+values.map(v=>`<div class="tt-row"><span class="tt-dot" style="background:${v.color}"></span>${esc(v.name)}: <b>${v.value.toFixed(1)}</b> МВт</div>`).join('');
  const wrap=canvas.parentElement.getBoundingClientRect();let left=event.clientX-wrap.left+14,top=event.clientY-wrap.top+14;tip.style.display='block';
  if(left+tip.offsetWidth>wrap.width-8)left=event.clientX-wrap.left-tip.offsetWidth-14;
  if(top+tip.offsetHeight>wrap.height-8)top=event.clientY-wrap.top-tip.offsetHeight-14;
  tip.style.left=Math.max(8,left)+'px';tip.style.top=Math.max(8,top)+'px';
}
"""
