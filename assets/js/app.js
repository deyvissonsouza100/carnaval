
const COLS = "ABCDEFGHIJKLMN".split("");

function showToast(msg, type='ok'){
  const el = document.querySelector('#toast');
  el.textContent = msg;
  el.classList.toggle('warn', type==='warn');
  el.style.display = 'block';
}
function clearToast(){ document.querySelector('#toast').style.display='none'; }
function cell(r, c){ return document.querySelector(`[data-r="${r}"][data-c="${c}"]`); }
function esc(s){
  return String(s||'')
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;');
}

function focusCell(r, c){
  const td = cell(r,c);
  if(td) td.focus();
}

function clearGrid(){
  for(let r=1;r<=10;r++){
    for(const c of COLS) cell(r, c).textContent='';
  }
  document.querySelector('#out').innerHTML='';
  clearToast();
}

function gridToRows(){
  const rows=[];
  for(let r=1;r<=10;r++){
    const obj={};
    let filled=false;
    for(const c of COLS){
      const v = cell(r,c).textContent.trim();
      obj[c]=v;
      if(v) filled=true;
    }
    if(filled) rows.push({r, ...obj});
  }
  return rows;
}

function mapEvent(row){
  // MAPEAMENTO FIXO (padrão da sua planilha):
  // A = Evento (ID do bloco)
  // B = Horário de Concentração
  // D = Data do evento
  // F = Previsão para Dispersão
  // H = Público Estimado (declarado)
  // J = Percurso
  // K = Local de Concentração (endereço)
  // Agentes = manual (no quadro)
  return {
    rowNum: row.r,
    evento: row.A || '',
    data: row.D || '',
    conc: row.B || '',
    disp: row.F || '',
    publico: row.H || '',
    local: row.K || '',
    percurso: row.J || '',
    agentes: '',
    agentesColor: '#111827',
  };
}

// Word table with borders + agents color
function wordHTML(i, e){
  const tableStyle = "border-collapse:collapse;width:100%;font-family:Calibri,Arial;font-size:11pt;";
  const td1 = "border:1px solid #9CA3AF;padding:6px 8px;font-weight:bold;background:#F3F4F6;color:#111827;width:42%;";
  const td2 = "border:1px solid #9CA3AF;padding:6px 8px;background:#FFFFFF;color:#111827;";
  const titleStyle = "font-family:Calibri,Arial;font-size:14pt;font-weight:bold;margin:0 0 6px 0;color:#111827;";
  const badgeStyle = "display:inline-block;border:1px solid #0EA5E9;background:#E0F2FE;color:#075985;border-radius:999px;padding:2px 10px;font-size:10pt;";
  const agentesSpan = `<span style="color:${esc(e.agentesColor||'#111827')};font-weight:bold">${esc(e.agentes||'')}</span>`;

  return `
    <div>
      <p style="${titleStyle}">Evento ${i+1}: ${esc(e.evento || '(sem nome)')} <span style="${badgeStyle}">${esc(e.data||'')}</span></p>
      <table style="${tableStyle}">
        <tr><td style="${td1}">Data</td><td style="${td2}">${esc(e.data)}</td></tr>
        <tr><td style="${td1}">Horário de Concentração</td><td style="${td2}">${esc(e.conc)}</td></tr>
        <tr><td style="${td1}">Previsão para Dispersão</td><td style="${td2}">${esc(e.disp)}</td></tr>
        <tr><td style="${td1}">Público Estimado</td><td style="${td2}">${esc(e.publico)}</td></tr>
        <tr><td style="${td1}">Local de Concentração</td><td style="${td2}">${esc(e.local)}</td></tr>
        <tr><td style="${td1}">Percurso</td><td style="${td2}">${esc(e.percurso)}</td></tr>
        <tr><td style="${td1}">Quantidade de Agentes</td><td style="${td2}">${agentesSpan}</td></tr>
      </table>
    </div>
  `;
}

async function copyHtmlToClipboard(html, plain){
  try{
    if(navigator.clipboard && window.ClipboardItem){
      const item = new ClipboardItem({
        "text/html": new Blob([html], {type:"text/html"}),
        "text/plain": new Blob([plain], {type:"text/plain"})
      });
      await navigator.clipboard.write([item]);
      return true;
    }
  }catch(e){}
  const tmp = document.createElement('div');
  tmp.style.position='fixed';
  tmp.style.left='-9999px';
  tmp.innerHTML = html;
  document.body.appendChild(tmp);
  const range = document.createRange();
  range.selectNodeContents(tmp);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  try{
    document.execCommand('copy');
    sel.removeAllRanges();
    document.body.removeChild(tmp);
    return true;
  }catch(e){
    sel.removeAllRanges();
    document.body.removeChild(tmp);
    return false;
  }
}

function plainTextFromEvent(e, i){
  return `Evento ${i+1}: ${e.evento}\nData: ${e.data}\nHorário de Concentração: ${e.conc}\nPrevisão para Dispersão: ${e.disp}\nPúblico Estimado: ${e.publico}\nLocal de Concentração: ${e.local}\nPercurso: ${e.percurso}\nQuantidade de Agentes: ${e.agentes}\n`;
}

function readEventFromCard(idx){
  const card = document.querySelector(`[data-card-idx="${idx}"]`);
  if(!card) return null;
  const get = (sel)=> (card.querySelector(sel)?.textContent || '').trim();
  const getInput = (sel)=> (card.querySelector(sel)?.value || '').trim();
  const color = card.dataset.agentesColor || '#111827';
  return {
    evento: get('[data-f="evento"]'),
    data: get('[data-f="data"]'),
    conc: get('[data-f="conc"]'),
    disp: get('[data-f="disp"]'),
    publico: get('[data-f="publico"]'),
    local: get('[data-f="local"]'),
    percurso: get('[data-f="percurso"]'),
    agentes: getInput('[data-f="agentes"]'),
    agentesColor: color,
  };
}

async function copyOne(idx){
  const e = readEventFromCard(idx);
  if(!e) return;
  const html = wordHTML(idx, e);
  const plain = plainTextFromEvent(e, idx);
  const ok = await copyHtmlToClipboard(html, plain);
  showToast(ok ? 'Quadro copiado! Cole no Word.' : 'Falha ao copiar. Tente novamente.', ok ? 'ok':'warn');
}

async function copyAll(){
  const cards = [...document.querySelectorAll('[data-card-idx]')];
  const events = cards.map((_, i)=>readEventFromCard(i)).filter(Boolean);
  const html = events.map((e,i)=>wordHTML(i,e)).join('<br/>');
  const plain = events.map((e,i)=>plainTextFromEvent(e,i)).join('\n');
  const ok = await copyHtmlToClipboard(html, plain);
  showToast(ok ? 'Todos os quadros foram copiados (cole no Word).' : 'Falha ao copiar. Tente copiar um por um.', ok ? 'ok':'warn');
}

// colors for agentes
const COLORS = [
  {name:'Preto', value:'#111827'},
  {name:'Vermelho', value:'#dc2626'},
  {name:'Azul', value:'#2563eb'},
  {name:'Verde', value:'#16a34a'},
  {name:'Roxo', value:'#7c3aed'},
];

function colorButtonsHTML(active){
  return `
    <div class="colorbar" role="group" aria-label="Cores de agentes">
      ${COLORS.map(c=>`
        <button type="button" class="colorbtn ${active===c.value?'active':''}" data-color="${c.value}" title="${c.name}">
          <span style="background:${c.value}"></span>
        </button>
      `).join('')}
      <span class="small-muted">Cor do número/letra (vai pro Word).</span>
    </div>
  `;
}

function makeEventCard(i, e){
  return `
  <div class="event" data-card-idx="${i}" data-agentes-color="${esc(e.agentesColor||'#111827')}">
    <div class="title">
      <b>Evento ${i+1}: <span data-f="evento">${esc(e.evento || '(sem nome)')}</span></b>
      <span class="badge" data-f="data">${esc(e.data || 'sem data')}</span>
    </div>
    <table>
      <tr><td>Data</td><td data-f="data">${esc(e.data)}</td></tr>
      <tr><td>Horário de Concentração</td><td data-f="conc">${esc(e.conc)}</td></tr>
      <tr><td>Previsão para Dispersão</td><td data-f="disp">${esc(e.disp)}</td></tr>
      <tr><td>Público Estimado</td><td data-f="publico">${esc(e.publico)}</td></tr>
      <tr><td>Local de Concentração</td><td data-f="local">${esc(e.local)}</td></tr>
      <tr><td>Percurso</td><td data-f="percurso">${esc(e.percurso)}</td></tr>
      <tr>
        <td>Quantidade de Agentes</td>
        <td>
          <input class="input" data-f="agentes" placeholder="Digite após gerar..." value="${esc(e.agentes)}"
                 style="font-weight:900;color:${esc(e.agentesColor||'#111827')}" />
          ${colorButtonsHTML(e.agentesColor||'#111827')}
        </td>
      </tr>
    </table>
    <div class="bd copybar" style="padding:12px 14px;border-top:1px solid rgba(255,255,255,.08)">
      <button data-copy="one" data-idx="${i}">Copiar este quadro (Word)</button>
      <span class="small-muted">Copia com bordas e cor de “Agentes”.</span>
    </div>
  </div>`;
}

function bindColorPickers(){
  document.querySelectorAll('.event').forEach(card=>{
    const input = card.querySelector('[data-f="agentes"]');
    card.querySelectorAll('.colorbtn').forEach(btn=>{
      btn.addEventListener('click', ()=>{
        const color = btn.dataset.color;
        card.dataset.agentesColor = color;
        input.style.color = color;
        card.querySelectorAll('.colorbtn').forEach(b=>b.classList.toggle('active', b===btn));
      });
    });
  });
}

function generate(){
  clearToast();
  const gridRows = gridToRows();
  if(gridRows.length===0){
    showToast('Preencha pelo menos 1 linha (A..N) ou use "Importar do Excel".', 'warn');
    document.querySelector('#out').innerHTML='';
    return;
  }
  const events = gridRows.map(mapEvent);
  document.querySelector('#out').innerHTML = events.map((e,i)=>makeEventCard(i,e)).join('');

  document.querySelector('#copy-all').onclick = ()=>copyAll();
  document.querySelectorAll('[data-copy="one"]').forEach(btn=>{
    btn.addEventListener('click', ()=>copyOne(parseInt(btn.dataset.idx,10)));
  });
  bindColorPickers();
  showToast(`Gerado(s) ${events.length} quadro(s) de evento.`, 'ok');
}

// ---- paste / import
function detectAndBuildMatrix(text){
  const rawLines = text.split(/\r?\n/);
  const tabCounts = rawLines.map(l=> (l.match(/\t/g)||[]).length );
  const linesWithManyTabs = tabCounts.filter(t=>t>=2).length;

  // multiple lines but only one looks like a "real row" => merge as one row
  if(rawLines.length > 1 && linesWithManyTabs <= 1){
    const merged = rawLines.join(' ↵ ');
    return [merged.split('\t')];
  }
  const lines = rawLines.filter(l=>l.length>0);
  return lines.map(line=>line.split('\t'));
}

function fillFromMatrix(matrix, r0=1, c0=0){
  for(let i=0;i<matrix.length;i++){
    for(let j=0;j<matrix[i].length;j++){
      const rr = r0 + i;
      const cc = c0 + j;
      if(rr>10 || cc>=COLS.length) continue;
      cell(rr, COLS[cc]).textContent = matrix[i][j];
    }
  }
}

function handlePaste(e){
  const text = e.clipboardData.getData('text/plain');
  if(!text) return;

  const start = e.target;
  const r0 = parseInt(start.dataset.r, 10);
  const c0 = COLS.indexOf(start.dataset.c);
  if(!Number.isFinite(r0) || c0<0) return;

  const matrix = detectAndBuildMatrix(text);
  fillFromMatrix(matrix, r0, c0);
  e.preventDefault();
}

function importFromTextarea(){
  clearToast();
  const text = document.querySelector('#import-text').value;
  if(!text.trim()){
    showToast('Cole no campo "Importar do Excel" e clique em Importar.', 'warn');
    return;
  }
  const matrix = detectAndBuildMatrix(text);
  fillFromMatrix(matrix, 1, 0);
  showToast('Importado para a grade (A1). Agora clique em "Gerar quadros".', 'ok');
}

function buildGrid(){
  const tbody = document.querySelector('#sheet-body');
  let html='';
  for(let r=1;r<=10;r++){
    html += `<tr><th>${r}</th>`;
    for(const c of COLS){
      html += `<td contenteditable="true" data-r="${r}" data-c="${c}" tabindex="0"></td>`;
    }
    // button after N
    html += `<td style="text-align:center;min-width:96px">
              <button type="button" class="primary" data-next-row="${r}" style="padding:6px 10px;border-radius:10px">↵ Próxima</button>
            </td>`;
    html += `</tr>`;
  }
  tbody.innerHTML = html;

  document.querySelectorAll('.sheet td[contenteditable="true"]').forEach(td=>{
    td.addEventListener('paste', handlePaste);
  });

  document.querySelectorAll('[data-next-row]').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const r = parseInt(btn.dataset.nextRow,10);
      focusCell(Math.min(10, r+1), 'A');
    });
  });
}

function loadExample(){
  clearGrid();
  const row = {A:"BLOCO DAS CRIANÇAS1", D:"01/02/2026", E:"07:00", F:"14:00", H:"1000", K:"PRACA RUI BARBOSA, 50", J:"PRACA RUI BARBOSA, 50, Centro -> PRACA RUI BARBOSA, 20, Centro -> RUA AARAO REIS, 585, Centro", I:"2"};
  for(const [k,v] of Object.entries(row)) cell(1,k).textContent = v;
  showToast('Exemplo preenchido na linha 1. Clique em "Gerar quadros".', 'ok');
}

document.addEventListener('DOMContentLoaded', ()=>{
  buildGrid();
  document.querySelector('#btn-generate').addEventListener('click', generate);
  document.querySelector('#btn-clear').addEventListener('click', clearGrid);
  document.querySelector('#btn-example').addEventListener('click', loadExample);
  document.querySelector('#btn-import').addEventListener('click', importFromTextarea);
});
