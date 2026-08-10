"""Standalone HTML generation."""
from __future__ import annotations

import html
import json
import re
from dataclasses import asdict
from pathlib import Path

from .chart_generator import CHART_JS, EVAL_JS
from .models import Model

CSS = """
:root{--bg:#f3f6fa;--card:#fff;--text:#172033;--muted:#667085;--border:#d9e2ee;--blue:#1f5bb5;--blue2:#edf5ff;--green:#047857;--green-bg:#d1fae5;--amber:#b45309;--amber-bg:#fef3c7}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);font-family:Inter,-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;color:var(--text);font-size:14px;line-height:1.48;-webkit-font-smoothing:antialiased}
.page{max-width:1560px;margin:0 auto;padding:18px 22px 32px}
.hero{background:linear-gradient(135deg,#163b78,#2563a9);color:#fff;border-radius:16px;padding:20px 24px;box-shadow:0 8px 24px rgba(30,58,138,.2);margin-bottom:14px}
.hero h1{margin:0;font-size:25px;line-height:1.25;letter-spacing:-.015em}
.top-grid{display:grid;grid-template-columns:1.1fr .9fr;gap:16px;margin-bottom:16px}.top-grid.single{grid-template-columns:1fr}
.card{background:var(--card);border:1px solid var(--border);border-radius:14px;box-shadow:0 3px 12px rgba(15,23,42,.055);overflow:hidden;margin-bottom:14px}
.top-grid .card{margin-bottom:0}
.card-h{padding:12px 16px;border-bottom:1px solid var(--border);font-weight:700;background:#fafbfd;font-size:14px}
.card-b{padding:14px 16px}
details.card>summary{list-style:none}
details.card>summary::-webkit-details-marker{display:none}
.info-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:10px;margin-top:12px}.info-item{border:1px solid var(--border);border-radius:11px;padding:10px 12px;background:#f8fbff}.info-label{font-size:10px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);font-weight:750;margin-bottom:4px}.nk-value{font-size:20px;color:var(--blue);font-weight:800}.nk-unit{margin-left:5px;font-weight:700;color:#475569}.weather-list{display:flex;gap:5px;flex-wrap:wrap}.weather-badge{display:inline-flex;padding:3px 8px;border-radius:999px;background:#e8f2ff;color:#245996;font-size:12px;font-weight:650}
.elements{margin:0;padding:0;list-style:none}.elements li{position:relative;margin:6px 0;padding-left:22px}.elements li::before{content:'⚡';position:absolute;left:0;top:0;color:#e59b12}
.mode-grid,.factor-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:10px;align-items:end}
.mode-grid label,.factor-grid label{display:block;font-size:12px;font-weight:650;margin-bottom:5px;color:#334155}
.mode-grid select,.mode-grid input,.factor-grid input,select,input{width:100%;padding:9px 10px;border:1px solid var(--border);border-radius:10px;background:#fff;font:inherit}
.bool-row{display:flex;gap:9px;align-items:center;padding:10px 12px;border:1px solid var(--border);border-radius:10px;background:#fff}.bool-row input{width:auto}.bool-row label{margin:0;flex:1}
.mode-actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:12px}.btn{border:1px solid var(--border);background:#fff;border-radius:10px;padding:8px 12px;cursor:pointer;font-weight:650}.btn:hover{border-color:var(--blue);color:var(--blue);background:var(--blue2)}.calc-switch{display:inline-flex!important;align-items:center;gap:8px;margin:0!important;padding:7px 11px;border:1px solid var(--border);border-radius:10px;background:#fff;color:#26364a!important;font-size:13px!important;cursor:pointer}.calc-switch input{width:auto!important;margin:0}
.mode-status{display:inline-flex;padding:5px 12px;border-radius:999px;background:var(--green-bg);color:var(--green);font-weight:700;margin-left:auto}
.toolbar{display:flex;gap:10px;align-items:center;margin:16px 0 10px;position:sticky;top:0;z-index:15;background:rgba(243,246,250,.94);backdrop-filter:blur(8px);padding:8px 0}.toolbar input{flex:1;min-width:220px;background:#fff}.count{color:var(--muted);white-space:nowrap;font-size:13px}
.repair{background:#fff;border:1px solid var(--border);border-radius:12px;margin:8px 0;overflow:hidden;box-shadow:0 2px 7px rgba(15,23,42,.035)}.repair-head{display:flex;gap:12px;align-items:center;padding:11px 14px;cursor:pointer}.repair-head:hover{background:var(--blue2)}.rid{font-weight:800;color:#496078;min-width:48px}.rname{font-weight:650;flex:1;white-space:pre-line;line-height:1.35}.uncontrolled-badge{display:inline-flex;align-items:center;white-space:nowrap;padding:3px 9px;border-radius:999px;background:#fee2e2;color:#b42318;border:1px solid #fecaca;font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.025em}.chev{color:var(--muted);transition:transform .2s}.repair.open .chev{transform:rotate(180deg)}.repair-body{display:none;border-top:1px solid var(--border);padding:12px 14px 14px}.repair.open .repair-body{display:block}
.row-actions{display:flex;gap:8px;margin-bottom:8px}.copy-btn{font-size:11px;padding:5px 9px;border-radius:9px;border:1px solid var(--border);background:#fff;cursor:pointer;font-weight:650}.copy-btn:hover{background:var(--blue2);color:var(--blue);border-color:var(--blue)}
.table-wrap{width:100%;max-width:100%;overflow:auto;border-radius:9px;border:1px solid #dfe6ef}
table{width:100%;max-width:100%;border-collapse:collapse;border:1px solid #dfe6ef;font-size:11.5px;table-layout:fixed}
th,td{border:1px solid #dfe6ef;padding:6px 7px;vertical-align:top;overflow-wrap:anywhere;word-break:normal}
th{position:sticky;top:0;z-index:2;background:#f6f8fb;color:#536174;font-size:10px;text-transform:uppercase;letter-spacing:.035em;font-weight:750;text-align:center;line-height:1.25}.tnv{font-weight:700;text-align:center;white-space:normal;color:#334155}.adp-cell{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:11px;font-weight:400;color:#1e293b;text-align:center;vertical-align:top}.scheme-adp-criteria{vertical-align:top}.criteria-start{border-left:1px solid #dfe6ef}.col-tnv{width:7%}.col-adp{width:8%}.col-adp-crit{width:17%}.col-formula{width:30%}.col-criterion{width:38%}.groups-2 .col-tnv{width:6%}.groups-2 .col-adp{width:7%}.groups-2 .col-adp-crit{width:13%}.groups-2 .col-formula{width:18%}.groups-2 .col-criterion{width:19%}
.formula{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:11px;padding:3px 4px;border-radius:7px;margin:1px 0;line-height:1.35}
.formula .value{font-size:13px;font-weight:800;color:#1456a0;margin-top:3px}
.adp-cell .formula .value{font-size:11px;font-weight:400;color:#1e293b;margin:0}
.inline-value{display:inline;white-space:nowrap;font-weight:800;color:#1456a0}
.formula.minimum,.crit.minimum{background:#dcfce7;box-shadow:inset 0 0 0 2px #16a34a;border-radius:8px}.crit.minimum{margin-left:-4px;padding-left:8px}
.minimum-from{font-size:11px;font-weight:700;color:#334155;margin:0 0 2px}
.min-badge{display:inline-block;background:#16a34a;color:#fff;border-radius:999px;padding:1px 7px;font-size:9px;font-weight:800;margin-left:5px;text-transform:uppercase}
.planning-badge{display:inline-block;margin-left:6px;padding:1px 6px;border-radius:999px;background:#fff3cd;color:#8a4b08;border:1px solid #f3d18a;font-size:8.5px;font-weight:750;white-space:nowrap;vertical-align:1px}.control-cell{font-size:10.5px;color:#334155;vertical-align:top}.col-control{width:13%}
.factor-defs{width:100%;table-layout:auto;font-size:12px}.factor-defs td:first-child{width:30%;font-weight:700;color:#334155}.factor-defs td{padding:7px 9px}.factor-defs tr:last-child td{border-bottom:0}
.note{margin-top:8px;color:#566070}
.empty{color:#8b95a5}
.chart-controls{display:grid;grid-template-columns:2fr 1fr 1fr 1.4fr repeat(3,.8fr);gap:8px;align-items:end}.chart-controls label{font-size:11px;color:var(--muted);font-weight:650}
.chart-wrap{overflow:auto;margin-top:10px;background:#fff;border:1px solid var(--border);border-radius:12px;padding:10px;position:relative}
canvas{width:100%;min-width:700px;height:340px}
.chart-tooltip{position:absolute;display:none;pointer-events:none;background:#111827;color:#fff;border-radius:9px;padding:8px 10px;font-size:12px;line-height:1.45;box-shadow:0 8px 24px rgba(15,23,42,.28);z-index:5}.chart-tooltip b{font-weight:800}.tt-row{white-space:nowrap}.tt-dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px}.chart-legend{display:flex;gap:10px 16px;flex-wrap:wrap;margin-top:8px;font-size:12px;color:#475569}.legend-line{display:inline-block;width:20px;height:4px;border-radius:3px;margin-right:6px;vertical-align:middle}
.chart-note{font-size:12px;color:var(--muted);margin-top:8px}
@media(max-width:900px){.page{padding:8px}.top-grid{grid-template-columns:1fr}.chart-controls{grid-template-columns:1fr 1fr}.mode-status{margin-left:0}table{font-size:10px;min-width:920px}th,td{padding:5px}.toolbar{top:0}}
@media print{
  @page{size:A4 landscape;margin:7mm}
  body{background:#fff}
  .page{max-width:none;padding:0}
  .hero,.top-grid,.toolbar,details.card,.chart-card,.row-actions{display:none!important}
  .repair{border:0;border-radius:0;margin:0;break-inside:avoid}
  .repair-head{padding:5px 0;background:#fff}
  .repair-body{display:block!important;padding:0}
  .table-wrap{overflow:visible}
  table{min-width:0;width:100%;font-size:7pt}
  th,td{padding:2px}
  .formula{padding:1px}
  .formula .value{font-size:7pt}
}
"""

RENDER_JS = r"""
function renderModeControls(){
  const box=document.getElementById('modeGrid');
  if(!box) return;
  box.innerHTML=(DATA.mode_params||[]).map(p=>{
    const id=mid(p.name);
    if(p.kind==='bool') return `<div class="bool-row"><input id="${id}" type="checkbox"><label for="${id}">${esc(parameterLabel(p.name))}</label></div>`;
    if(p.kind==='select') return `<div><label>${esc(parameterLabel(p.name))}</label><select id="${id}">${(p.options||[]).map(o=>`<option value="${esc(o.value)}"${String(o.value)===String(p.default)?' selected':''}>${esc(o.label)}</option>`).join('')}</select></div>`;
    return `<div><label>${esc(parameterLabel(p.name))}</label><input id="${id}" type="number" value="${esc(p.default||'0')}" step="any"></div>`;
  }).join('');
  box.querySelectorAll('select,input').forEach(e=>e.addEventListener('input',()=>{updateModeStatus();render();drawChart();}));
  updateModeStatus();
}

function renderFactors(){
  const box=document.getElementById('factorGrid');
  if(!box) return;
  box.innerHTML=(DATA.factors||[]).map(f=>`<div><label>${esc(parameterLabel(f.name))}</label><input id="${fid(f.name)}" type="number" value="${f.default||0}" step="any"></div>`).join('');
  box.querySelectorAll('input').forEach(e=>e.addEventListener('input',()=>{render();drawChart();}));
}

function render(){
  const q=(document.getElementById('q')?.value||'').toLowerCase();
  const env=factorValues();
  const openBefore=new Set(Array.from(document.querySelectorAll('.repair.open')).map(x=>x.dataset.si));
  let out='';
  let shown=0;
  DATA.schemes.forEach((s,si)=>{
    if(!(String(s.number)+' '+s.name).toLowerCase().includes(q)) return;
    shown++;
    const schemeRows=s.rows||[];
    const hasControlMdp=DATA.has_mdp&&schemeRows.some(r=>String(r.control_mdp||'').trim());
    const hasControlMdpPa=DATA.has_mdp_pa&&schemeRows.some(r=>String(r.control_mdp_pa||'').trim());
    const hasControlAdp=DATA.has_adp&&schemeRows.some(r=>String(r.control_adp||'').trim());
    const mergeControlMdp=hasControlMdp&&!controlValuesVary(schemeRows,'control_mdp');
    const mergeControlMdpPa=hasControlMdpPa&&!controlValuesVary(schemeRows,'control_mdp_pa');
    const mergeControlAdp=hasControlAdp&&!controlValuesVary(schemeRows,'control_adp');
    const schemeControlMdp=uniqueSchemeText(schemeRows,'control_mdp');
    const schemeControlMdpPa=uniqueSchemeText(schemeRows,'control_mdp_pa');
    const schemeControlAdp=uniqueSchemeText(schemeRows,'control_adp');
    const formulaCols=`${DATA.has_mdp?'<col class="col-formula">':''}${DATA.has_mdp_pa?'<col class="col-formula">':''}${DATA.has_adp?'<col class="col-adp">':''}`;
    const criteriaCols=`${DATA.has_mdp?'<col class="col-criterion">':''}${DATA.has_mdp_pa?'<col class="col-criterion">':''}${DATA.has_adp?'<col class="col-adp-crit">':''}`;
    const controlCols=`${hasControlMdp?'<col class="col-control">':''}${hasControlMdpPa?'<col class="col-control">':''}${hasControlAdp?'<col class="col-control">':''}`;
    const cols=`<colgroup><col class="col-tnv">${formulaCols}${criteriaCols}${controlCols}</colgroup>`;
    const formulaHead=`${DATA.has_mdp?'<th>МДП без ПА</th>':''}${DATA.has_mdp_pa?'<th>МДП с ПА</th>':''}${DATA.has_adp?'<th>АДП</th>':''}`;
    const criteriaHead=`${DATA.has_mdp?'<th class="criteria-start">Критерии МДП без ПА</th>':''}${DATA.has_mdp_pa?`<th${DATA.has_mdp?'':' class="criteria-start"'}>Критерии МДП с ПА</th>`:''}${DATA.has_adp?`<th${DATA.has_mdp||DATA.has_mdp_pa?'':' class="criteria-start"'}>Критерий АДП</th>`:''}`;
    const controlHead=`${hasControlMdp?'<th>Контроль доп. параметров МДП без ПА</th>':''}${hasControlMdpPa?'<th>Контроль доп. параметров МДП с ПА</th>':''}${hasControlAdp?'<th>Контроль доп. параметров АДП</th>':''}`;
    const head=`<tr><th>${esc(DATA.row_axis_label||'ТНВ')}</th>${formulaHead}${criteriaHead}${controlHead}</tr>`;
    const statusBadge=s.is_controlled===false?'<span class="uncontrolled-badge">Не контролируется</span>':'';
    if(!schemeRows.length){
      const open=openBefore.has(String(si))?' open':'';
      const emptyText=s.is_controlled===false?'Ремонтная схема не контролируется':(s.note||'Данные допустимых перетоков для этой схемы отсутствуют');
      out+=`<section class="repair${open}" data-si="${si}"><div class="repair-head" onclick="this.parentElement.classList.toggle('open')"><div class="rid">${esc(s.number)}</div><div class="rname">${esc(s.name)}</div>${statusBadge}<div class="chev">⌄</div></div><div class="repair-body"><div class="empty"><b>${esc(emptyText)}</b></div></div></section>`;
      return;
    }
    const adpCriteria=uniqueSchemeText(schemeRows,'crit_adp');
    const adpHtml=schemeAdpBlock(schemeRows,env);
    const rowCount=schemeRows.length;
    const body=schemeRows.map((r,ri)=>{
      let mdpHtml='', mdpPaHtml='';
      let critMdpHtml='', critMdpPaHtml='';
      if(DATA.has_mdp){
        mdpHtml=(r.mdp_items||[]).length?formulaBlock(r.mdp_items||[],env):textLines(r.mdp);
        if(mdpHtml) mdpHtml=withMinimumPrefix(mdpHtml,r.mdp);
        critMdpHtml=criteriaBlock(r.mdp_items||[],env,r.crit_mdp);
      }
      if(DATA.has_mdp_pa){
        const pa=activePaGroup(r);
        mdpPaHtml=(pa.items||[]).length?formulaBlock(pa.items||[],env):textLines(pa.raw||r.mdp_pa);
        if(mdpPaHtml) mdpPaHtml=withMinimumPrefix(mdpPaHtml,pa.raw||r.mdp_pa);
        critMdpPaHtml=criteriaBlock(pa.items||[],env,pa.crit||r.crit_mdp_pa);
      }
      return `<tr><td class="tnv">${textLines(r.temperature)}</td>`+
        `${DATA.has_mdp?`<td>${mdpHtml}</td>`:''}`+
        `${DATA.has_mdp_pa?`<td>${mdpPaHtml}</td>`:''}`+
        `${DATA.has_adp&&ri===0?`<td ${mergedCellAttrs(rowCount,'adp-cell')}>${adpHtml}</td>`:''}`+
        `${DATA.has_mdp?`<td class="criteria-start">${critMdpHtml}</td>`:''}`+
        `${DATA.has_mdp_pa?`<td${DATA.has_mdp?'':' class="criteria-start"'}>${critMdpPaHtml}</td>`:''}`+
        `${DATA.has_adp&&ri===0?`<td ${mergedCellAttrs(rowCount,'scheme-adp-criteria'+(DATA.has_mdp||DATA.has_mdp_pa?'':' criteria-start'))}>${criteriaBlock((schemeRows.find(x=>(x.adp_items||[]).length)||{}).adp_items||[],env,adpCriteria,false)}</td>`:''}`+
        `${hasControlMdp?(mergeControlMdp?(ri===0?`<td ${mergedCellAttrs(rowCount,'control-cell')}>${textLines(schemeControlMdp)}</td>`:''):`<td class="control-cell">${textLines(r.control_mdp)}</td>`):''}`+
        `${hasControlMdpPa?(mergeControlMdpPa?(ri===0?`<td ${mergedCellAttrs(rowCount,'control-cell')}>${textLines(schemeControlMdpPa)}</td>`:''):`<td class="control-cell">${textLines(r.control_mdp_pa)}</td>`):''}`+
        `${hasControlAdp?(mergeControlAdp?(ri===0?`<td ${mergedCellAttrs(rowCount,'control-cell')}>${textLines(schemeControlAdp)}</td>`:''):`<td class="control-cell">${textLines(r.control_adp)}</td>`):''}</tr>`;
    }).join('');
    const open=openBefore.has(String(si))||(!document.querySelector('.repair')&&shown===1)?' open':'';
    const copyButtons=`${DATA.has_mdp?`<button class="copy-btn" onclick="copyMdpSeries(${si},'mdp',this)">Копировать МДП без ПА</button>`:''}${DATA.has_mdp_pa?`<button class="copy-btn" onclick="copyMdpSeries(${si},'mdp_pa',this)">Копировать МДП с ПА</button>`:''}`;
    out+=`<section class="repair${open}" data-si="${si}"><div class="repair-head" onclick="this.parentElement.classList.toggle('open')"><div class="rid">${esc(s.number)}</div><div class="rname">${esc(s.name)}</div>${statusBadge}<div class="chev">⌄</div></div><div class="repair-body"><div class="row-actions">${copyButtons}</div><div class="table-wrap"><table class="groups-${Number(DATA.has_mdp)+Number(DATA.has_mdp_pa)}">${cols}<thead>${head}</thead><tbody>${body}</tbody></table></div>${s.note?`<div class="note"><b>Примечание:</b> ${esc(s.note)}</div>`:''}</div></section>`;
  });
  document.getElementById('items').innerHTML=out||'<div class="empty">Ничего не найдено</div>';
  const count=document.getElementById('count'); if(count) count.textContent=`Показано: ${shown} из ${DATA.schemes.length}`;
}

function parameterLabel(name){
  const raw=String(name||''); const low=raw.toLowerCase();
  if(low==='pa_season') return DATA.pa_season_label||'Группа стабилизации МДП с ПА';
  if(low==='ртд') return String(DATA.title||'').toLowerCase().includes('печорск')?'Состояние реактора на Печорской ГРЭС':'Состояние реактора';
  if(low.includes('количество_генераторов_в_работе_на_втэц_2')) return 'Количество генераторов в работе на ВТЭЦ-2';
  if(low.includes('кол_во_тг')||low.includes('количество_работающих_энергоблоков')) return 'Количество работающих блоков';
  return raw.replace(/_{3,}/g,' – ').replace(/_+[–—-]_+/g,' – ').replace(/_+/g,' ').replace(/\s*[–—]\s*/g,' – ').replace(/\s+/g,' ').trim();
}

function updateModeStatus(){
  const box=document.getElementById('modeStatus'); if(!box) return;
  const labels=(DATA.mode_params||[]).map(p=>{const el=document.getElementById(mid(p.name));if(!el)return '';if(p.kind==='bool')return `${parameterLabel(p.name)}: ${el.checked?'включено':'отключено'}`;if(p.kind==='select')return `${parameterLabel(p.name)}: ${el.options[el.selectedIndex]?.text||el.value}`;return `${parameterLabel(p.name)}: ${el.value}`;}).filter(Boolean);
  box.textContent=labels.join(', ')||'Параметры режима не требуются';
}

function expandAll(){document.querySelectorAll('.repair').forEach(x=>x.classList.add('open'));}
function collapseAll(){document.querySelectorAll('.repair').forEach(x=>x.classList.remove('open'));}

function padCopyColumn(value,targetStop){
  const text=String(value??'');
  const length=Array.from(text).length;
  return text+'\t'.repeat(Math.max(1,Math.ceil((targetStop-length)/8)));
}

async function writeClipboard(text,button){
  let copied=false;
  try{if(navigator.clipboard?.writeText){await navigator.clipboard.writeText(text);copied=true;}}catch(_err){}
  if(!copied){
    const area=document.createElement('textarea');area.value=text;area.style.position='fixed';area.style.opacity='0';document.body.appendChild(area);area.select();
    try{copied=document.execCommand('copy');}catch(_err){} area.remove();
  }
  if(button){const old=button.textContent;button.textContent=copied?'Скопировано':'Не удалось скопировать';setTimeout(()=>button.textContent=old,1600);}
}

function copyMdpSeries(schemeIndex,group,button){
  const scheme=DATA.schemes?.[schemeIndex];if(!scheme)return;
  const rawKey=group==='mdp_pa'?'mdp_pa':'mdp';
  const title=group==='mdp_pa'?'МДП с ПА':'МДП без ПА';
  const axisLabel=DATA.row_axis_label||'ТНВ';
  const temperatures=(scheme.rows||[]).map(row=>String(row.temperature||''));
  const maxLength=Math.max(Array.from(axisLabel).length,...temperatures.map(value=>Array.from(value).length));
  const targetStop=Math.ceil((maxLength+1)/8)*8;
  const lines=[padCopyColumn(axisLabel,targetStop)+title];
  (scheme.rows||[]).forEach(row=>{
    let items=[];
    let rawText='';
    if(group==='mdp_pa'){
      const pa=activePaGroup(row);
      items=pa.items||[];
      rawText=pa.raw||row.mdp_pa||'';
    } else {
      items=row.mdp_items||[];
      rawText=row[rawKey]||'';
    }
    const prepared=(items||[]).map(item=>({number:item.number||'',text:String(item.raw||'').trim()})).filter(item=>item.text);
    if(!prepared.length&&String(rawText||'').trim())prepared.push({number:'',text:String(rawText).trim()});
    const needsMinimum=copyNeedsMinimumPrefix(rawText, prepared.length);
    if(needsMinimum){
      lines.push(padCopyColumn(row.temperature,targetStop)+'Минимальный из:');
      prepared.forEach(item=>lines.push(padCopyColumn('',targetStop)+`${item.number?item.number+') ':''}${item.text}`));
    } else {
      prepared.forEach((item,index)=>lines.push(padCopyColumn(index===0?row.temperature:'',targetStop)+`${item.number?item.number+') ':''}${item.text}`));
    }
  });
  const footer=formatCopyFactorDefinitions(collectSchemeFormulaVariables(scheme,group));
  if(footer.length){ lines.push(''); lines.push('где:'); footer.forEach(line=>lines.push(line)); }
  writeClipboard(lines.join('\n'),button);
}

function uniqueSchemeText(rows,key){
  const seen=new Set(), values=[];
  (rows||[]).forEach(r=>{
    const value=String(r?.[key]||'').trim();
    if(value&&!seen.has(value)){seen.add(value);values.push(value);}
  });
  return values.join('\n');
}

function controlValuesVary(rows,key){
  const seen=new Set();
  for(const row of rows||[]){
    const value=String(row?.[key]||'').trim();
    if(!value) continue;
    if(seen.size&&!seen.has(value)) return true;
    seen.add(value);
  }
  return false;
}

function mergedCellAttrs(count, className){
  if(count<2) return `class="${className}"`;
  return `rowspan="${count}" class="${className}"`;
}

function schemeAdpBlock(rows,env){
  const row=(rows||[]).find(r=>(r.adp_items||[]).length)||null;
  if(!row) return textLines(uniqueSchemeText(rows,'adp'));
  const item=(row.adp_items||[])[0];
  if(!item?.is_computable||!item.ast) return textLines(item?.raw||row.adp);
  const active=activeAst(item.ast,env);
  return `<div class="formula">${esc(formatAst(active)||item.raw)}</div>`;
}

function boot(){
  renderModeControls();
  if(OPTIONS.calc){ renderFactors(); }
  document.getElementById('q')?.addEventListener('input', render);
  render();
  initChart();
}
boot();
"""


def _serialize_model(model: Model) -> dict:
    data = asdict(model)
    return data


def _factor_definitions_table(model: Model) -> str:
    if not model.factor_definitions:
        return ""
    rows = "".join(
        "<tr>"
        f"<td>{html.escape(factor.name)}</td>"
        f"<td>{html.escape(factor.description or 'Описание в исходном файле не задано')}</td>"
        "</tr>"
        for factor in model.factor_definitions
    )
    return (
        '<div class="table-wrap"><table class="factor-defs"><tbody>'
        f"{rows}"
        "</tbody></table></div>"
    )


def _factor_definition_block(model: Model) -> str:
    definitions_table = _factor_definitions_table(model)
    if not definitions_table:
        return ""
    return (
        '<details class="card"><summary class="card-h">Влияющие факторы</summary>'
        f'<div class="card-b">{definitions_table}</div></details>'
    )


def _factor_values_block(include_calculation: bool) -> str:
    if not include_calculation:
        return '<div id="factorGrid" style="display:none"></div>'
    return (
        '<details class="card"><summary class="card-h">Значения влияющих факторов</summary>'
        '<div class="card-b"><div id="factorGrid" class="factor-grid"></div></div></details>'
    )


def _factor_blocks(model: Model, include_calculation: bool) -> str:
    parts = [_factor_definition_block(model), _factor_values_block(include_calculation)]
    return "".join(parts) if any(parts) else '<div id="factorGrid" style="display:none"></div>'


def generate(model: Model, out: str | Path, include_calculation: bool = True, include_chart: bool = True) -> None:
    include_chart = bool(include_calculation and include_chart)
    data = json.dumps(_serialize_model(model), ensure_ascii=False).replace("</", "<\\/")
    opts = json.dumps({"calc": bool(include_calculation), "chart": include_chart})

    elems = "".join(f"<li>{html.escape(x)}</li>" for x in model.elements)
    title_match = re.search(r"[«\"]([^»\"]+)[»\"]", model.title)
    display_title = f"КС «{title_match.group(1)}»" if title_match else model.title

    info_items: list[str] = []
    if model.irregular_oscillation_mw is not None:
        value = f"{model.irregular_oscillation_mw:g}"
        info_items.append(
            '<div class="info-item"><div class="info-label">Нерегулярные колебания</div>'
            f'<span class="nk-value">{html.escape(value)}</span><span class="nk-unit">МВт</span></div>'
        )
    if model.weather_stations:
        stations = "".join(
            f'<span class="weather-badge">{html.escape(station)}</span>'
            for station in model.weather_stations
        )
        info_items.append(
            '<div class="info-item"><div class="info-label">Метеостанции</div>'
            f'<div class="weather-list">{stations}</div></div>'
        )
    info_grid = f'<div class="info-grid">{"".join(info_items)}</div>' if info_items else ""
    composition_block = (
        '<section class="card"><div class="card-h">Состав контролируемого сечения</div>'
        f'<div class="card-b"><ul class="elements">{elems}</ul>{info_grid}</div></section>'
    )

    mode_block = ""
    if include_calculation:
        mode_block = (
            '<section class="card"><div class="card-h">Режим расчёта</div>'
            '<div class="card-b"><div id="modeGrid" class="mode-grid"></div>'
            '<div class="mode-actions"><label class="calc-switch"><input id="calcToggle" type="checkbox" onchange="render()">Вычислять МДП</label>'
            '<button class="btn" onclick="expandAll()">Раскрыть все</button>'
            '<button class="btn" onclick="collapseAll()">Свернуть все</button>'
            '<span id="modeStatus" class="mode-status"></span></div></div></section>'
        )
    else:
        mode_block = (
            '<section class="card"><div class="card-h">Режим расчёта</div>'
            '<div class="card-b"><div id="modeGrid" class="mode-grid"></div>'
            '<div class="mode-actions"><span id="modeStatus" class="mode-status"></span>'
            '</div></div></section>'
            if model.mode_params
            else '<div id="modeGrid" style="display:none"></div>'
        )
    factor_block = _factor_blocks(model, include_calculation)

    chart_block = ""
    if include_chart:
        chart_block = (
            '<details class="card chart-card"><summary class="card-h">График зависимости МДП от влияющих факторов</summary>'
            '<div class="card-b"><div class="chart-controls">'
            '<label>Схема сети<select id="chartScheme"></select></label>'
            '<label>ТНВ<select id="chartRow"></select></label>'
            '<label>Группа МДП<select id="chartGroup"></select></label>'
            '<label>Влияющий фактор<select id="chartFactor"></select></label>'
            '<label>От<input id="chartMin" type="number" value="-100"></label>'
            '<label>До<input id="chartMax" type="number" value="100"></label>'
            '<label>Шаг<input id="chartStep" type="number" value="10" min="0.1" step="any"></label>'
            '</div><div class="chart-wrap"><canvas id="chart"></canvas>'
            '<div id="chartTooltip" class="chart-tooltip"></div>'
            '<div id="chartLegend" class="chart-legend"></div></div>'
            '<div class="chart-note">Перемещайте указатель по графику: вертикальный маркер покажет точные значения МДП в выбранной точке.</div>'
            '</div></details>'
        )

    doc = f"""<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(model.title)}</title>
<style>{CSS}</style>
</head>
<body>
<div class="page">
  <div class="hero"><h1>{html.escape(display_title)}</h1></div>
  <div class="top-grid{' single' if not include_calculation and not model.mode_params else ''}">{composition_block}{mode_block}</div>
  {factor_block}
  {chart_block}
  <div class="toolbar"><input id="q" placeholder="Поиск по номеру или названию ремонтной схемы…"><span id="count" class="count"></span></div>
  <div id="items"></div>
</div>
<script>
const DATA={data};
const OPTIONS={opts};
{EVAL_JS}
{CHART_JS}
{RENDER_JS}
</script>
</body>
</html>"""
    Path(out).write_text(doc, encoding="utf-8")
