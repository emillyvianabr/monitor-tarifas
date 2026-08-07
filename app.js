const DEFAULT_FILE = 'data/dados.xlsx';
let workbookData = { tariffs: [], googleFlights: [], measures: [], glossary: [] };
let charts = {};

const fmtBRL = n => Number.isFinite(n) ? n.toLocaleString('pt-BR',{style:'currency',currency:'BRL'}) : '—';
const fmtPct = n => Number.isFinite(n) ? `${n > 0 ? '+' : ''}${n.toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})}%` : '—';
const monthLabel = (year, month) => `${String(month).padStart(2,'0')}/${year}`;
const num = v => {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v;
  const parsed = Number(String(v).replace(/\./g,'').replace(',','.').replace(/[^0-9.-]/g,''));
  return Number.isFinite(parsed) ? parsed : null;
};

function excelDate(v){
  if (!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'number') {
    const p = XLSX.SSF.parse_date_code(v);
    return p ? new Date(p.y,p.m-1,p.d) : null;
  }
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function rowsFromSheet(wb, name, options={}){
  const ws = wb.Sheets[name];
  return ws ? XLSX.utils.sheet_to_json(ws,{defval:'', ...options}) : [];
}

function normalizeWorkbook(wb){
  const tariffs = rowsFromSheet(wb,'BASE_TARIFAS_ANAC').map(r => ({
    date: excelDate(r.Mes_Referencia),
    year: num(r.Ano), month: num(r.Mes_Num), monthName: r.Mes || '',
    origin: r.Estado_Origem || 'Não informado', airport: r.Aeroporto_Origem || '',
    destination: r.Aeroporto_Destino || 'POA', airline: r.Companhia || 'Não informado',
    fare: num(r['Tarifa_Media_R$']), source: r.Fonte || '', updated: excelDate(r.Data_Atualizacao), note:r.Observacao || ''
  })).filter(r => r.year && r.month && Number.isFinite(r.fare));

  const googleFlights = rowsFromSheet(wb,'BASE_GOOGLE_FLIGHTS').map(r => ({
    searchDate: excelDate(r.Data_Busca), origin:r.Estado_Origem || 'Não informado', airport:r.Aeroporto_Origem || '',
    airline:r.Companhia || 'Não informado', price:num(r['Preco_R$']), tripType:r.Tipo_Viagem || '',
    outbound:excelDate(r.Data_Ida), returnDate:excelDate(r.Data_Volta)
  })).filter(r => r.searchDate && Number.isFinite(r.price));

  const measures = rowsFromSheet(wb,'INCENTIVOS_MEDIDAS').filter(r => r.Medida_Fator).map(r => ({
    title:r.Medida_Fator, period:r.Inicio_Periodo, scope:r.Abrangencia, effect:r.Efeito_Esperado, detail:r.Detalhe
  }));

  const glossaryRows = rowsFromSheet(wb,'GLOSSARIO_METODOLOGIA',{range:2});
  const glossary = glossaryRows.filter(r => r.Item && r['Definição sugerida']).map(r => ({
    item:r.Item, definition:r['Definição sugerida'], dashboard:r['Como aparece no dashboard'], attention:r['Atenção ao atualizar']
  }));

  workbookData = { tariffs, googleFlights, measures, glossary };
}

async function loadDefault(){
  try {
    const res = await fetch(`${DEFAULT_FILE}?v=${Date.now()}`);
    if(!res.ok) throw new Error('Arquivo não encontrado');
    const buffer = await res.arrayBuffer();
    const wb = XLSX.read(buffer,{type:'array',cellDates:true});
    normalizeWorkbook(wb);
    document.getElementById('dataSourceLabel').textContent = 'Fonte: data/dados.xlsx';
    renderAll();
  } catch(err){
    document.getElementById('updatedAt').textContent = 'Não foi possível carregar data/dados.xlsx. Use “Carregar planilha”.';
    console.error(err);
  }
}

function populateSelect(id, values){
  const el = document.getElementById(id);
  const current = el.value;
  el.innerHTML = `<option value="TODOS">${id==='airlineFilter'?'Todas':'Todos'}</option>`;
  [...new Set(values.filter(Boolean))].sort().forEach(v => {
    const o=document.createElement('option'); o.value=v; o.textContent=v; el.appendChild(o);
  });
  if([...el.options].some(o=>o.value===current)) el.value=current;
}

function setupFilters(){
  const t = workbookData.tariffs;
  populateSelect('originFilter', t.map(r=>r.origin));
  populateSelect('airlineFilter', t.map(r=>r.airline));
  populateSelect('yearFilter', t.map(r=>String(r.year)));
  populateSelect('gfOriginFilter', workbookData.googleFlights.map(r=>r.origin));

  const yearEl = document.getElementById('yearFilter');
  if (!yearEl.dataset.initialized) {
    if ([...yearEl.options].some(o => o.value === '2026')) yearEl.value = '2026';
    yearEl.dataset.initialized = 'true';
  }
}

function filteredTariffs(){
  const origin=document.getElementById('originFilter').value;
  const airline=document.getElementById('airlineFilter').value;
  const year=document.getElementById('yearFilter').value;
  return workbookData.tariffs.filter(r =>
    (origin==='TODOS'||r.origin===origin) &&
    (airline==='TODOS'||r.airline===airline) &&
    (year==='TODOS'||String(r.year)===year)
  );
}

function avg(arr){ const v=arr.filter(Number.isFinite); return v.length?v.reduce((a,b)=>a+b,0)/v.length:null; }

function updateKPIs(data){
  const fares=data.map(r=>r.fare).filter(Number.isFinite);
  const mean=avg(fares);
  const minRow=data.reduce((a,b)=>!a||b.fare<a.fare?b:a,null);
  const maxRow=data.reduce((a,b)=>!a||b.fare>a.fare?b:a,null);
  document.getElementById('kpiAvg').textContent=fmtBRL(mean);
  document.getElementById('kpiMin').textContent=minRow?fmtBRL(minRow.fare):'—';
  document.getElementById('kpiMinNote').textContent=minRow?`${monthLabel(minRow.year,minRow.month)} • ${minRow.airline}`:'—';
  document.getElementById('kpiMax').textContent=maxRow?fmtBRL(maxRow.fare):'—';
  document.getElementById('kpiMaxNote').textContent=maxRow?`${monthLabel(maxRow.year,maxRow.month)} • ${maxRow.airline}`:'—';

  const byYM={}; data.forEach(r=>{ const k=`${r.year}-${r.month}`; (byYM[k] ||= []).push(r.fare); });
  const pairs=[];
  Object.keys(byYM).forEach(k=>{ const [y,m]=k.split('-').map(Number); const prev=`${y-1}-${m}`; if(byYM[prev]) pairs.push([avg(byYM[k]),avg(byYM[prev])]); });
  const yoy = pairs.length ? avg(pairs.map(([cur,prev])=>(cur/prev-1)*100)) : null;
  const yoyEl=document.getElementById('kpiYoY'); yoyEl.textContent=fmtPct(yoy); yoyEl.className = Number.isFinite(yoy) ? (yoy>0?'positive':'negative') : '';
  document.getElementById('kpiYoYNote').textContent = pairs.length ? `${pairs.length} mês(es) comparável(is)` : 'sem meses comparáveis';
}

function chartBaseOptions(){
  return {responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{usePointStyle:true,boxWidth:8,font:{size:11}}},tooltip:{mode:'index',intersect:false}},scales:{x:{grid:{display:false},ticks:{font:{size:10}}},y:{grid:{color:'#edf0f4'},ticks:{callback:v=>'R$ '+Number(v).toLocaleString('pt-BR'),font:{size:10}}}}};
}
function destroyChart(name){ if(charts[name]){ charts[name].destroy(); delete charts[name]; } }

function updateMonthlyChart(data){
  destroyChart('monthly');
  const months=[...new Set(data.map(r=>`${r.year}-${String(r.month).padStart(2,'0')}`))].sort();
  const airlines=[...new Set(data.map(r=>r.airline))].sort();
  const datasets=airlines.map(a=>({label:a,data:months.map(k=>{const [y,m]=k.split('-').map(Number);return avg(data.filter(r=>r.airline===a&&r.year===y&&r.month===m).map(r=>r.fare));}),tension:.28,spanGaps:true,borderWidth:2,pointRadius:2}));
  charts.monthly=new Chart(document.getElementById('monthlyChart'),{type:'line',data:{labels:months.map(k=>{const [y,m]=k.split('-');return `${m}/${y}`}),datasets},options:chartBaseOptions()});
}

function updateBarChart(canvasId,name,groups,data,key){
  destroyChart(name);
  charts[name]=new Chart(document.getElementById(canvasId),{type:'bar',data:{labels:groups,datasets:[{label:'Tarifa média',data:groups.map(g=>avg(data.filter(r=>r[key]===g).map(r=>r.fare))),borderRadius:7,maxBarThickness:54}]},options:{...chartBaseOptions(),plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmtBRL(c.raw)}}}}});
}

function updateGoogleFlights(){
  const origin=document.getElementById('gfOriginFilter').value;
  let data=workbookData.googleFlights.filter(r=>origin==='TODOS'||r.origin===origin);
  if(data.length){ const maxDate=new Date(Math.max(...data.map(r=>r.searchDate))); const cutoff=new Date(maxDate); cutoff.setDate(cutoff.getDate()-59); data=data.filter(r=>r.searchDate>=cutoff); }
  const empty=document.getElementById('gfEmpty');
  destroyChart('gf');
  if(!data.length){ empty.classList.remove('hidden'); document.getElementById('googleFlightsChart').style.display='none'; return; }
  empty.classList.add('hidden'); document.getElementById('googleFlightsChart').style.display='block';
  const dates=[...new Set(data.map(r=>r.searchDate.toISOString().slice(0,10)))].sort();
  const origins=[...new Set(data.map(r=>r.origin))].sort();
  const datasets=origins.map(o=>({label:o,data:dates.map(d=>avg(data.filter(r=>r.origin===o&&r.searchDate.toISOString().slice(0,10)===d).map(r=>r.price))),tension:.3,spanGaps:true,borderWidth:2,pointRadius:2}));
  charts.gf=new Chart(document.getElementById('googleFlightsChart'),{type:'line',data:{labels:dates.map(d=>new Date(d+'T12:00:00').toLocaleDateString('pt-BR')),datasets},options:chartBaseOptions()});
}

function renderMeasures(){
  const c=document.getElementById('measuresList');
  c.innerHTML=workbookData.measures.length?workbookData.measures.map(m=>`<div class="measure-item"><strong>${m.title}</strong><p>${m.detail||''}</p><div class="measure-meta"><span class="tag">${m.period||'Período não informado'}</span><span class="tag">${m.scope||'Abrangência não informada'}</span>${m.effect?`<span class="tag">${m.effect}</span>`:''}</div></div>`).join(''):'<div class="empty-state">Sem medidas cadastradas.</div>';
}

function renderMethodology(){
  const c=document.getElementById('methodologyList');
  const fallback=[
    {item:'Valor por trecho',definition:'Confirmar na fonte se o preço representa um único trecho ou a viagem completa.'},
    {item:'Taxas',definition:'Registrar se taxas aeroportuárias, impostos e demais encargos estão incluídos.'},
    {item:'Mês de referência',definition:'Diferenciar claramente mês da compra/emissão de mês da viagem ou referência estatística.'},
    {item:'Google Flights',definition:'O histórico é construído pela data da consulta. Não substituir registros antigos.'}
  ];
  const list=workbookData.glossary.length?workbookData.glossary.slice(0,7):fallback;
  c.innerHTML=list.map(g=>`<div class="method-item"><strong>${g.item}</strong><p>${g.definition||''}</p></div>`).join('');
}

function renderTable(data){
  document.getElementById('rowCount').textContent=`${data.length} registro${data.length===1?'':'s'}`;
  const tbody=document.getElementById('dataTableBody');
  tbody.innerHTML=data.slice().sort((a,b)=>a.year-b.year||a.month-b.month||a.airline.localeCompare(b.airline)).map(r=>`<tr><td>${monthLabel(r.year,r.month)}</td><td>${r.origin}</td><td>${r.airline}</td><td>${fmtBRL(r.fare)}</td><td>${r.source||'—'}</td></tr>`).join('');
}

function updateTimestamp(){
  const dates=workbookData.tariffs.map(r=>r.updated).filter(Boolean);
  const latest=dates.length?new Date(Math.max(...dates)):null;
  document.getElementById('updatedAt').textContent=latest?`Última atualização informada: ${latest.toLocaleDateString('pt-BR')}`:`${workbookData.tariffs.length} tarifas válidas carregadas`;
}


function updateYearComparison(){
  destroyChart('yearComparison');
  const origin=document.getElementById('originFilter').value;
  const airline=document.getElementById('airlineFilter').value;

  const data=workbookData.tariffs.filter(r =>
    (origin==='TODOS'||r.origin===origin) &&
    (airline==='TODOS'||r.airline===airline) &&
    (r.year===2025||r.year===2026)
  );

  const labels=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  const seriesFor = year => labels.map((_,i) =>
    avg(data.filter(r=>r.year===year && r.month===i+1).map(r=>r.fare))
  );

  charts.yearComparison=new Chart(document.getElementById('yearComparisonChart'),{
    type:'line',
    data:{
      labels,
      datasets:[
        {label:'2025',data:seriesFor(2025),tension:.28,spanGaps:true,borderWidth:2.5,pointRadius:3},
        {label:'2026',data:seriesFor(2026),tension:.28,spanGaps:true,borderWidth:2.5,pointRadius:3}
      ]
    },
    options:chartBaseOptions()
  });
}

function renderFiltered(){
  const data=filteredTariffs();
  updateKPIs(data); updateMonthlyChart(data); updateYearComparison();
  updateBarChart('originChart','origin',[...new Set(data.map(r=>r.origin))].sort(),data,'origin');
  updateBarChart('airlineChart','airline',[...new Set(data.map(r=>r.airline))].sort(),data,'airline');
  renderTable(data);
}

function renderAll(){
  setupFilters(); updateTimestamp(); renderFiltered(); updateGoogleFlights(); renderMeasures(); renderMethodology();
}

['originFilter','airlineFilter','yearFilter'].forEach(id=>document.getElementById(id).addEventListener('change',renderFiltered));
document.getElementById('gfOriginFilter').addEventListener('change',updateGoogleFlights);
document.getElementById('clearFilters').addEventListener('click',()=>{
  document.getElementById('originFilter').value='TODOS';
  document.getElementById('airlineFilter').value='TODOS';
  document.getElementById('yearFilter').value=[...document.getElementById('yearFilter').options].some(o=>o.value==='2026')?'2026':'TODOS';
  renderFiltered();
});
document.getElementById('resetBtn').addEventListener('click',loadDefault);
document.getElementById('fileInput').addEventListener('change',async e=>{
  const file=e.target.files[0]; if(!file) return;
  const buffer=await file.arrayBuffer(); const wb=XLSX.read(buffer,{type:'array',cellDates:true}); normalizeWorkbook(wb);
  document.getElementById('dataSourceLabel').textContent=`Fonte temporária: ${file.name}`; renderAll();
});

loadDefault();
