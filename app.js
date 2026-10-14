const DEFAULT_FILE = 'data/dados.xlsx';
let workbookData = { tariffs: [], googleFlights: [], measures: [], glossary: [] };
let dataLoaded = false;
let charts = {};

// Paleta Grupo Iter: índigo, teal, azul, rosa e slate
const ORIGIN_COLORS = {
  RJ: '#4B4F9E',
  SP: '#62BAB6',
  MG: '#1D6DB6',
  SC: '#EB4A75',
  PR: '#53778B'
};
const originColor = origin => ORIGIN_COLORS[origin] || '#7693A1';
const CHART_COLORS = ['#4B4F9E','#62BAB6','#1D6DB6','#EB4A75','#53778B','#7693A1'];
const UI = { ink:'#233744', muted:'#53778B', grid:'#E8EEF1', up:'#EB4A75', down:'#1F8A68' };
const PREFERRED_ORIGINS = ['RJ','SP','MG','SC','PR'];

/* ---------- Destinos ----------
   Cada destino filtra a planilha pelo código do aeroporto de destino (coluna Aeroporto_Destino).
   Um destino só mostra gráficos quando a planilha já tem linhas para ele; enquanto isso,
   aparece a mensagem "Estamos trabalhando nisso". Para liberar o Rio de Janeiro, basta
   incluir linhas com Aeroporto_Destino = GIG (ou SDU) nas abas BASE_TARIFAS_ANAC / BASE_GOOGLE_FLIGHTS. */
const DESTINATIONS = [
  { id:'POA', slug:'poa', label:'Porto Alegre', codes:['POA'],
    subtitle:'Porto Alegre · principais mercados emissores de Gramado e Canela: SP, RJ, MG, SC e PR.' },
  { id:'RIO', slug:'rio', label:'Rio de Janeiro', codes:['GIG','SDU','RIO'],
    subtitle:'Rio de Janeiro · principais mercados emissores do destino.',
    subtitleSoon:'Rio de Janeiro · análise em preparação.' }
];
let currentDestinationId = 'POA';
const currentDestination = () => DESTINATIONS.find(d => d.id === currentDestinationId) || DESTINATIONS[0];
const inDestination = (row, dest = currentDestination()) => dest.codes.includes(String(row.destination || '').toUpperCase());
const scopedTariffs = () => workbookData.tariffs.filter(r => inDestination(r));
const scopedFlights = () => workbookData.googleFlights.filter(r => inDestination(r));
const destinationHasData = dest => workbookData.tariffs.some(r => inDestination(r, dest)) || workbookData.googleFlights.some(r => inDestination(r, dest));
const normText = t => String(t ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
// Medidas podem trazer uma coluna opcional "Destino" (ex.: POA, GIG ou "Rio de Janeiro"); sem ela valem para todos.
function measureMatches(m, dest = currentDestination()){
  if (!m.destination) return true;
  return m.destination.split(/[,;/|]/).map(normText).some(t => !t || t === 'todos' || dest.codes.some(c => normText(c) === t) || normText(dest.label) === t);
}

Chart.defaults.font.family = "Stolzl, system-ui, -apple-system, 'Segoe UI', sans-serif";
Chart.defaults.color = UI.muted;

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
    destination: String(r.Aeroporto_Destino || 'POA').trim().toUpperCase(), airline: r.Companhia || 'Não informado',
    fare: num(r['Tarifa_Media_R$']), source: r.Fonte || '', updated: excelDate(r.Data_Atualizacao), note:r.Observacao || ''
  })).filter(r => r.year && r.month && Number.isFinite(r.fare));

  const googleFlights = rowsFromSheet(wb,'BASE_GOOGLE_FLIGHTS').map(r => ({
    searchDate: excelDate(r.Data_Busca), origin:r.Estado_Origem || 'Não informado', airport:r.Aeroporto_Origem || '',
    destination:String(r.Aeroporto_Destino || 'POA').trim().toUpperCase(),
    airline:r.Companhia || 'Não informado', price:num(r['Preco_R$']), tripType:r.Tipo_Viagem || '',
    outbound:excelDate(r.Data_Ida), returnDate:excelDate(r.Data_Volta),
    source:r.Fonte || 'Google Flights', note:r.Observacao || ''
  })).filter(r => r.searchDate);

  const measures = rowsFromSheet(wb,'INCENTIVOS_MEDIDAS').filter(r => r.Medida_Fator).map(r => ({
    title:r.Medida_Fator, period:r.Inicio_Periodo, scope:r.Abrangencia, destination:String(r.Destino || '').trim(), effect:r.Efeito_Esperado, detail:r.Detalhe, impact:r.Como_Impacta || ''
  }));

  const glossaryRows = rowsFromSheet(wb,'GLOSSARIO_METODOLOGIA',{range:2});
  const glossary = glossaryRows.filter(r => r.Item && r['Definição sugerida']).map(r => ({
    item:r.Item, definition:r['Definição sugerida'], dashboard:r['Como aparece no dashboard'], attention:r['Atenção ao atualizar']
  }));

  workbookData = { tariffs, googleFlights, measures, glossary };
  dataLoaded = true;
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
  const t = scopedTariffs();
  populateSelect('originFilter', t.map(r=>r.origin));
  populateSelect('airlineFilter', t.map(r=>r.airline));
  populateSelect('yearFilter', t.map(r=>String(r.year)));

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
  return scopedTariffs().filter(r =>
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

  // Variação anual usa sempre os mesmos meses disponíveis nos dois anos.
  // Ao trocar o filtro de ano, a direção da comparação é invertida.
  const origin=document.getElementById('originFilter').value;
  const airline=document.getElementById('airlineFilter').value;
  const selectedYear=document.getElementById('yearFilter').value;

  const comparisonData=scopedTariffs().filter(r =>
    (origin==='TODOS'||r.origin===origin) &&
    (airline==='TODOS'||r.airline===airline) &&
    (r.year===2025||r.year===2026)
  );

  const months2025=[...new Set(comparisonData.filter(r=>r.year===2025).map(r=>r.month))];
  const months2026=[...new Set(comparisonData.filter(r=>r.year===2026).map(r=>r.month))];
  const comparableMonths=months2025.filter(m=>months2026.includes(m)).sort((a,b)=>a-b);

  let currentYear = selectedYear==='2025' ? 2025 : 2026;
  let comparisonYear = currentYear===2026 ? 2025 : 2026;

  const currentAvg=avg(comparisonData
    .filter(r=>r.year===currentYear && comparableMonths.includes(r.month))
    .map(r=>r.fare));
  const comparisonAvg=avg(comparisonData
    .filter(r=>r.year===comparisonYear && comparableMonths.includes(r.month))
    .map(r=>r.fare));

  const yoy=(Number.isFinite(currentAvg)&&Number.isFinite(comparisonAvg)&&comparisonAvg!==0)
    ? (currentAvg/comparisonAvg-1)*100
    : null;

  const yoyEl=document.getElementById('kpiYoY');
  yoyEl.textContent=fmtPct(yoy);
  yoyEl.className = Number.isFinite(yoy) ? (yoy>0?'positive':yoy<0?'negative':'') : '';

  const lastMonth=comparableMonths.length?comparableMonths[comparableMonths.length-1]:null;
  const monthNamesShort=['','jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
  document.getElementById('kpiYoYNote').textContent =
    comparableMonths.length
      ? `${currentYear} vs ${comparisonYear} • jan–${monthNamesShort[lastMonth]}`
      : 'sem meses comparáveis';
}

function chartBaseOptions(){
  return {responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{usePointStyle:true,boxWidth:8,padding:16,font:{size:10}}},tooltip:{mode:'index',intersect:false}},scales:{x:{grid:{display:false},ticks:{font:{size:10},color:UI.muted}},y:{grid:{color:UI.grid},border:{display:false},ticks:{callback:v=>'R$ '+Number(v).toLocaleString('pt-BR'),font:{size:10},color:UI.muted}}}};
}
function destroyChart(name){ if(charts[name]){ charts[name].destroy(); delete charts[name]; } }

function updateMonthlyChart(data){
  destroyChart('monthly');
  const months=[...new Set(data.map(r=>`${r.year}-${String(r.month).padStart(2,'0')}`))].sort();
  const airlines=[...new Set(data.map(r=>r.airline))].sort();
  const datasets=airlines.map((a,index)=>({label:a,data:months.map(k=>{const [y,m]=k.split('-').map(Number);return avg(data.filter(r=>r.airline===a&&r.year===y&&r.month===m).map(r=>r.fare));}),tension:.32,spanGaps:true,borderWidth:2.5,pointRadius:3,borderColor:CHART_COLORS[index%CHART_COLORS.length],backgroundColor:CHART_COLORS[index%CHART_COLORS.length]}));
  charts.monthly=new Chart(document.getElementById('monthlyChart'),{type:'line',data:{labels:months.map(k=>{const [y,m]=k.split('-');return `${m}/${y}`}),datasets},options:chartBaseOptions()});
}

function updateBarChart(canvasId,name,groups,data,key){
  destroyChart(name);
  charts[name]=new Chart(document.getElementById(canvasId),{type:'bar',data:{labels:groups,datasets:[{label:'Tarifa média',data:groups.map(g=>avg(data.filter(r=>r[key]===g).map(r=>r.fare))),backgroundColor:groups.map((group,index)=>key==='origin'?originColor(group):CHART_COLORS[index%CHART_COLORS.length]),borderRadius:7,maxBarThickness:54}]},options:{...chartBaseOptions(),plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmtBRL(c.raw)}}}}});
}

function updateGoogleFlights(){
  const origin=document.getElementById('originFilter').value;
  let all=scopedFlights().filter(r=>origin==='TODOS'||r.origin===origin);

  const empty=document.getElementById('gfEmpty');
  destroyChart('gf');

  if(!all.length){
    empty.classList.remove('hidden');
    document.getElementById('googleFlightsChart').style.display='none';
    return;
  }

  empty.classList.add('hidden');
  document.getElementById('googleFlightsChart').style.display='block';

  const monthKey = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
  const monthLabelGF = key => {
    const [y,m]=key.split('-').map(Number);
    return new Date(y,m-1,1)
      .toLocaleDateString('pt-BR',{month:'short',year:'numeric'})
      .replace('.','');
  };

  const allMonthKeys=[...new Set(all.map(r=>monthKey(r.searchDate)))].sort();

  const routes=[...new Map(
    all.map(r=>[`${r.origin}|${r.airport}`, {origin:r.origin,airport:r.airport}])
  ).values()].sort((a,b)=>a.origin.localeCompare(b.origin));

  const datasets=routes.map(route=>{
    const routeRows=all.filter(r=>r.origin===route.origin && r.airport===route.airport);
    return {
      label:`${route.origin} (${route.airport})`,
      data:allMonthKeys.map(key=>{
        const values=routeRows
          .filter(r=>monthKey(r.searchDate)===key)
          .map(r=>r.price)
          .filter(Number.isFinite);
        return values.length ? avg(values) : null;
      }),
      tension:.25,
      spanGaps:false,
      borderWidth:2.5,
      pointRadius:4,
      pointHoverRadius:6,
      borderColor: originColor(route.origin),
      backgroundColor: originColor(route.origin),
      pointBackgroundColor: originColor(route.origin),
      pointBorderColor: '#ffffff'
    };
  });

  const options=chartBaseOptions();
  options.interaction={mode:'index',intersect:false};
  options.plugins.tooltip={
    mode:'index',
    intersect:false,
    callbacks:{
      label(ctx){
        return `${ctx.dataset.label}: ${Number.isFinite(ctx.raw)?fmtBRL(ctx.raw):'sem dados'}`;
      }
    }
  };

  charts.gf=new Chart(document.getElementById('googleFlightsChart'),{
    type:'line',
    data:{
      labels:allMonthKeys.map(monthLabelGF),
      datasets
    },
    options
  });

}



function renderMeasures(){
  const c=document.getElementById('measuresList');
  if(!workbookData.measures.length){
    c.innerHTML='<div class="empty-state">Sem medidas cadastradas.</div>';
    return;
  }

  const measures=workbookData.measures.filter(m=>measureMatches(m));
  const costRelief=measures.filter(m=>(m.effect||'').includes('↓'));
  const passengerCost=measures.filter(m=>(m.effect||'').includes('↑'));

  const normalizeEffect = effect => String(effect||'').replace('Custos das aéreas','Custos das cias aéreas').replace('Custo das aéreas','Custo das cias aéreas');

  const card=(m,type)=>`<div class="measure-item measure-item--${type}">
    <strong>${m.title}</strong>
    <p class="measure-detail">${m.detail||''}</p>
    ${m.impact?`<div class="measure-impact"><span>Como pode impactar</span><p>${m.impact}</p></div>`:''}
    <div class="measure-meta">
      <span class="tag">${m.period||'Período não informado'}</span>
      <span class="tag">${m.scope||'Abrangência não informada'}</span>
      ${m.effect?`<span class="tag measure-effect measure-effect--${type}">${normalizeEffect(m.effect)}</span>`:''}
    </div>
  </div>`;

  c.innerHTML=`
    <div class="measure-group-intro">As medidas abaixo não têm o mesmo efeito. Algumas reduzem custos ou pressão financeira das companhias e podem influenciar indiretamente as tarifas; outras, como tarifas aeroportuárias, aumentam o custo final do passageiro, mas <strong>não entram na Tarifa Aérea Média da ANAC</strong>.</div>
    ${costRelief.length?`<h3 class="measure-group-title measure-group-title--positive">Redução de custos e pressão financeira das cias aéreas</h3>${costRelief.map(m=>card(m,'positive')).join('')}`:''}
    ${passengerCost.length?`<h3 class="measure-group-title measure-group-title--negative">Aumento do custo total para o passageiro</h3>${passengerCost.map(m=>card(m,'negative')).join('')}`:''}
  `;
}

function renderMethodology(){
  const c=document.getElementById('methodologyList');
  c.innerHTML=`
    <div class="method-text-block">
      <h3>Tarifa Aérea Média — ANAC</h3>
      <p>A Tarifa Aérea Média da ANAC representa os valores das passagens efetivamente comercializadas ao público. O indicador é calculado de forma ponderada pelo volume de bilhetes vendidos em cada valor tarifário, portanto não é uma média simples dos preços anunciados. O mês de referência corresponde ao mês em que a passagem foi vendida, e não necessariamente ao mês da viagem.</p>
      <p>A série utilizada no painel é a <strong>Tarifa Aérea Real Média</strong>, ou seja, os valores são corrigidos pela inflação para permitir comparação ao longo do tempo. Em viagens de ida e volta, o valor é tratado por trecho de origem e destino.</p>

      <h3>O que entra e o que não entra</h3>
      <p>O indicador considera o valor do serviço de transporte aéreo comercializado. <strong>Não entram</strong> a tarifa de embarque/aeroportuária nem serviços opcionais cobrados separadamente, como bagagem despachada e marcação de assento. Por isso, um reajuste da tarifa aeroportuária pode aumentar o custo final pago pelo passageiro, mas não altera diretamente a Tarifa Aérea Média divulgada pela ANAC.</p>

      <h3>Incentivos e medidas que podem impactar</h3>
      <p>Algumas medidas não fazem parte do cálculo da tarifa, mas podem influenciar o preço definido pelas companhias. Entre elas estão custo e tributação do QAV, ICMS e outros tributos sobre combustível, preço do petróleo, câmbio, concorrência entre empresas, demanda, ocupação das aeronaves, oferta de voos e assentos, sazonalidade, alterações de malha e incentivos para ampliação da conectividade. Já as tarifas aeroportuárias devem ser tratadas como impacto no <strong>custo total da viagem</strong>, e não como componente da tarifa média ANAC.</p>

      <h3>Google Flights — média mensal</h3>
      <p>Os valores coletados no Google Flights correspondem ao <strong>menor preço exibido em cada dia</strong> para a rota acompanhada. O valor mensal mostrado no gráfico é calculado pela média aritmética desses menores preços diários disponíveis no mês: soma dos menores preços registrados ÷ número de dias com preço válido.</p>
      <p>Dias sem valor identificado não entram no cálculo. Quando o mês ainda não terminou, a média é parcial e deve ser interpretada como preliminar. Essa série não é diretamente equivalente à ANAC: a ANAC representa tarifas efetivamente vendidas, enquanto o Google Flights representa o menor preço ofertado observado diariamente.</p>
    </div>
  `;
}

function renderTable(data){
  document.getElementById('rowCount').textContent=`${data.length} registro${data.length===1?'':'s'}`;
  const tbody=document.getElementById('dataTableBody');
  tbody.innerHTML=data.slice().sort((a,b)=>a.year-b.year||a.month-b.month||a.airline.localeCompare(b.airline)).map(r=>`<tr><td>${monthLabel(r.year,r.month)}</td><td>${r.origin}</td><td>${r.airline}</td><td>${fmtBRL(r.fare)}</td><td>${r.source||'—'}</td></tr>`).join('');
}

function updateTimestamp(){
  const tariffs=scopedTariffs();
  const dates=tariffs.map(r=>r.updated).filter(Boolean);
  const latest=dates.length?new Date(Math.max(...dates)):null;
  document.getElementById('updatedAt').textContent=latest?`Última atualização informada: ${latest.toLocaleDateString('pt-BR')}`:`${tariffs.length} tarifas válidas carregadas`;
}


function updateYearComparison(){
  destroyChart('yearComparison');

  const origin=document.getElementById('originFilter').value;
  const airline=document.getElementById('airlineFilter').value;
  const selectedYear=document.getElementById('yearFilter').value;

  const data=scopedTariffs().filter(r =>
    (origin==='TODOS'||r.origin===origin) &&
    (airline==='TODOS'||r.airline===airline) &&
    (r.year===2025||r.year===2026)
  );

  // Mostra somente meses que existem nos dois anos.
  // Assim 2025 acompanha exatamente o período disponível de 2026.
  const months2025=[...new Set(data.filter(r=>r.year===2025).map(r=>r.month))];
  const months2026=[...new Set(data.filter(r=>r.year===2026).map(r=>r.month))];
  const comparableMonths=months2025.filter(m=>months2026.includes(m)).sort((a,b)=>a-b);

  const monthLabels=['','Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  const labels=comparableMonths.map(m=>monthLabels[m]);

  const seriesFor = year => comparableMonths.map(m =>
    avg(data.filter(r=>r.year===year && r.month===m).map(r=>r.fare))
  );

  const y2025=seriesFor(2025);
  const y2026=seriesFor(2026);

  // A direção das variações acompanha o ano selecionado no filtro.
  const currentYear = selectedYear==='2025' ? 2025 : 2026;
  const currentSeries = currentYear===2026 ? y2026 : y2025;
  const comparisonSeries = currentYear===2026 ? y2025 : y2026;
  const monthlyVariation=currentSeries.map((v,i)=>{
    const base=comparisonSeries[i];
    return Number.isFinite(v)&&Number.isFinite(base)&&base!==0 ? (v/base-1)*100 : null;
  });

  const variationLabelsPlugin={
    id:'variationLabels',
    afterDatasetsDraw(chart){
      const {ctx}=chart;
      const meta=chart.getDatasetMeta(currentYear===2026 ? 1 : 0);
      ctx.save();
      ctx.font="700 11px Stolzl, system-ui, sans-serif";
      ctx.textAlign='center';
      ctx.textBaseline='bottom';

      monthlyVariation.forEach((variation,i)=>{
        if(!Number.isFinite(variation) || !meta.data[i]) return;
        const point=meta.data[i];
        const text=`${variation>0?'+':''}${variation.toFixed(1).replace('.',',')}%`;
        ctx.fillStyle=variation>0 ? UI.up : variation<0 ? UI.down : UI.muted;
        ctx.fillText(text, point.x, point.y-10);
      });
      ctx.restore();
    }
  };

  const options=chartBaseOptions();
  options.layout={padding:{top:24,right:30,left:6}};
  options.plugins.tooltip={
    mode:'index',
    intersect:false,
    callbacks:{
      afterBody(items){
        if(!items.length) return '';
        const i=items[0].dataIndex;
        const v=monthlyVariation[i];
        return Number.isFinite(v)
          ? `Variação ${currentYear} vs ${currentYear===2026?2025:2026}: ${v>0?'+':''}${v.toFixed(1).replace('.',',')}%`
          : '';
      }
    }
  };

  charts.yearComparison=new Chart(document.getElementById('yearComparisonChart'),{
    type:'line',
    data:{
      labels,
      datasets:[
        {label:'2025',data:y2025,tension:.32,spanGaps:true,borderWidth:2.5,pointRadius:4,borderColor:'#62BAB6',backgroundColor:'#62BAB6'},
        {label:'2026',data:y2026,tension:.32,spanGaps:true,borderWidth:3,pointRadius:4,borderColor:'#4B4F9E',backgroundColor:'#4B4F9E'}
      ]
    },
    options,
    plugins:[variationLabelsPlugin]
  });
}


function updateOriginTrendChart(){
  destroyChart('originTrend');

  const yearValue = document.getElementById('yearFilter').value;
  const airline = document.getElementById('airlineFilter').value;

  // Compara todas as origens entre si; por isso ignora o filtro principal de origem.
  const scoped = scopedTariffs();
  let data = scoped.slice();

  if(yearValue !== 'TODOS'){
    data = data.filter(r => String(r.year) === yearValue);
  } else {
    // Se "Todos" estiver selecionado, usa o ano mais recente disponível
    // para evitar misturar meses de anos diferentes no mesmo eixo.
    const latestYear = Math.max(...data.map(r=>r.year).filter(Number.isFinite));
    data = data.filter(r => r.year === latestYear);
  }

  if(airline !== 'TODOS'){
    data = data.filter(r => r.airline === airline);
  }

  // Mercados emissores presentes neste destino (RJ, SP, MG, SC, PR primeiro; novos entram depois)
  const presentOrigins = [...new Set(scoped.map(r => r.origin).filter(Boolean))];
  const origins = [...PREFERRED_ORIGINS.filter(o => presentOrigins.includes(o)), ...presentOrigins.filter(o => !PREFERRED_ORIGINS.includes(o)).sort()];
  const months = [...new Set(data.map(r => r.month).filter(Number.isFinite))]
    .sort((a,b)=>a-b);

  const monthNamesOrigin = ['','Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  const labels = months.map(m => monthNamesOrigin[m] || String(m));

  const datasets = origins.map(origin => ({
    label: origin,
    data: months.map(m => {
      const vals = data
        .filter(r => r.origin === origin && r.month === m)
        .map(r => r.fare)
        .filter(Number.isFinite);
      return vals.length ? avg(vals) : null;
    }),
    tension: .28,
    spanGaps: false,
    borderWidth: 2.5,
    pointRadius: 3,
    pointHoverRadius: 6,
    borderColor: originColor(origin),
    backgroundColor: originColor(origin),
    pointBackgroundColor: originColor(origin),
    pointBorderColor: '#ffffff'
  }));

  const canvas = document.getElementById('originTrendChart');
  if(!canvas) return;

  const options = chartBaseOptions();
  options.interaction = {mode:'index', intersect:false};
  options.plugins.tooltip = {
    mode:'index',
    intersect:false,
    callbacks:{
      label(ctx){
        return `${ctx.dataset.label}: ${Number.isFinite(ctx.raw) ? fmtBRL(ctx.raw) : 'sem dados'}`;
      }
    }
  };

  charts.originTrend = new Chart(canvas,{
    type:'line',
    data:{labels,datasets},
    options
  });
}

function renderFiltered(){
  const data=filteredTariffs();
  updateKPIs(data); updateMonthlyChart(data); updateYearComparison();
  updateBarChart('originChart','origin',[...new Set(data.map(r=>r.origin))].sort(),data,'origin');
  updateBarChart('airlineChart','airline',[...new Set(data.map(r=>r.airline))].sort(),data,'airline');
  renderTable(data);
  updateOriginTrendChart();
}

/* ---------- Seletor de destino ---------- */
function buildDestinationTabs(){
  const host=document.getElementById('destTabs');
  host.innerHTML=DESTINATIONS.map(d=>{
    const ready=!dataLoaded||destinationHasData(d), selected=d.id===currentDestinationId;
    return `<button type="button" role="tab" class="dest-tab" data-dest="${d.id}" aria-selected="${selected}" tabindex="${selected?0:-1}">${d.label}${ready?'':'<span class="dest-soon">em breve</span>'}</button>`;
  }).join('');
}
function selectDestination(id, {updateHash=true}={}){
  if(!DESTINATIONS.some(d=>d.id===id)) return;
  currentDestinationId=id;
  if(updateHash){
    try{ history.replaceState(null,'','#destino='+currentDestination().slug); }catch(e){}
  }
  renderAll();
}
function destinationFromHash(){
  const m=(location.hash||'').match(/destino=([a-z]+)/i);
  const found=m&&DESTINATIONS.find(d=>d.slug===m[1].toLowerCase());
  return found?found.id:null;
}
document.getElementById('destTabs').addEventListener('click',e=>{
  const b=e.target.closest('[data-dest]'); if(b) selectDestination(b.dataset.dest);
});
document.getElementById('destTabs').addEventListener('keydown',e=>{
  if(e.key!=='ArrowRight'&&e.key!=='ArrowLeft') return;
  const i=DESTINATIONS.findIndex(d=>d.id===currentDestinationId);
  const next=DESTINATIONS[(i+(e.key==='ArrowRight'?1:-1)+DESTINATIONS.length)%DESTINATIONS.length];
  selectDestination(next.id);
  document.querySelector(`[data-dest="${next.id}"]`)?.focus();
});
document.getElementById('csBack').addEventListener('click',()=>{
  const first=DESTINATIONS.find(d=>destinationHasData(d))||DESTINATIONS[0];
  selectDestination(first.id);
});

function renderAll(){
  const dest=currentDestination();
  const ready=destinationHasData(dest);
  buildDestinationTabs();
  document.getElementById('destSubtitle').textContent=ready?dest.subtitle:(dest.subtitleSoon||dest.subtitle);

  // Destino sem dados: mostra "Estamos trabalhando nisso" e esconde os gráficos.
  document.getElementById('dashboardContent').hidden=!ready;
  const soon=document.getElementById('comingSoon');
  soon.hidden=ready;
  if(!ready){
    Object.keys(charts).forEach(destroyChart);
    const back=DESTINATIONS.find(d=>destinationHasData(d));
    document.getElementById('csText').innerHTML=`Os gráficos de tarifas aéreas para <strong>${dest.label}</strong> ainda estão sendo preparados. Em breve você acompanha aqui as mesmas análises disponíveis para ${back?back.label:'os demais destinos'}.`;
    const btn=document.getElementById('csBack');
    btn.hidden=!back; if(back) btn.textContent=`Ver ${back.label}`;
    document.getElementById('updatedAt').textContent='Dados em preparação';
    return;
  }
  setupFilters(); updateTimestamp(); renderFiltered(); updateGoogleFlights(); renderMeasures(); renderMethodology();
}

document.getElementById('originFilter').addEventListener('change',()=>{
  renderFiltered();
  updateGoogleFlights();
});
['airlineFilter','yearFilter'].forEach(id=>document.getElementById(id).addEventListener('change',renderFiltered));
document.getElementById('clearFilters').addEventListener('click',()=>{
  document.getElementById('originFilter').value='TODOS';
  document.getElementById('airlineFilter').value='TODOS';
  document.getElementById('yearFilter').value=[...document.getElementById('yearFilter').options].some(o=>o.value==='2026')?'2026':'TODOS';
  renderFiltered();
  updateGoogleFlights();
});
document.getElementById('resetBtn').addEventListener('click',loadDefault);
document.getElementById('fileInput').addEventListener('change',async e=>{
  const file=e.target.files[0]; if(!file) return;
  const buffer=await file.arrayBuffer(); const wb=XLSX.read(buffer,{type:'array',cellDates:true}); normalizeWorkbook(wb);
  document.getElementById('dataSourceLabel').textContent=`Fonte temporária: ${file.name}`; renderAll();
});

const hashDest=destinationFromHash(); if(hashDest) currentDestinationId=hashDest;
buildDestinationTabs();
loadDefault();

// A Stolzl é usada dentro dos gráficos (canvas): redesenha quando a fonte terminar de carregar.
if(document.fonts&&document.fonts.load){
  Promise.all([document.fonts.load('400 12px Stolzl'),document.fonts.load('700 12px Stolzl')]).then(()=>{
    if(workbookData.tariffs.length && destinationHasData(currentDestination())){ renderFiltered(); updateGoogleFlights(); }
  }).catch(()=>{});
}
