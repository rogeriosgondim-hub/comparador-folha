import * as pdfjsLib from 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs';

const $ = id => document.getElementById(id);
const state = { excelFile:null, pdfFile:null, original:[], folha:[], results:[], diagnostics:{} };

const STATUS = {
  OK:'OK', DIV:'DIVERGÊNCIA', CENT:'DIFERENÇA DE CENTAVOS', TRCT:'TRCT',
  NOVO:'NOVO COLABORADOR', SO_ORIG:'SOMENTE NO ORIGINAL', SO_FOLHA:'SOMENTE NA FOLHA',
  DUP:'DUPLICADO', VER:'VERIFICAR'
};
const moneyFmt = new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'});

function norm(v=''){
  return String(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[^A-Za-z0-9 ]+/g,' ').replace(/\s+/g,' ').trim().toUpperCase();
}
function normName(v=''){
  return norm(v).replace(/\bTRCT\b/g,'').replace(/\bDEMISSAO\b/g,'')
    .replace(/\bDEMITID[OA]\b/g,'').replace(/\s+/g,' ').trim();
}
function cpf(v=''){ const d=String(v).replace(/\D/g,''); return d.length===11?d:''; }
function fmtCpf(v){ const d=cpf(v); return d?d.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/,'$1.$2.$3-$4'):''; }
function money(v){
  if(v===null||v===undefined||v==='') return null;
  if(typeof v==='number'&&Number.isFinite(v)) return v;
  const raw=String(v).trim();
  if(!raw||/^(NDA|N\/A|NA|-)$/i.test(raw)) return null;

  let s=raw.replace(/R\$/gi,'').replace(/\s/g,'').replace(/[^0-9,().+\-]/g,'');
  let negative=false;
  if(/^\(.*\)$/.test(s)){ negative=true; s=s.slice(1,-1); }
  if(s.startsWith('-')){ negative=true; s=s.slice(1); }
  s=s.replace(/\+/g,'');
  if(!s) return null;

  const lastComma=s.lastIndexOf(',');
  const lastDot=s.lastIndexOf('.');
  let normalized=s;

  if(lastComma>=0 && lastDot>=0){
    if(lastComma>lastDot){
      normalized=s.replace(/\./g,'').replace(',','.');
    }else{
      normalized=s.replace(/,/g,'');
    }
  }else if(lastComma>=0){
    const decimals=s.length-lastComma-1;
    normalized=decimals===2 ? s.replace(/\./g,'').replace(',','.') : s.replace(/,/g,'');
  }else if(lastDot>=0){
    const decimals=s.length-lastDot-1;
    normalized=decimals===2 ? s.replace(/,/g,'') : s.replace(/\./g,'');
  }

  const n=Number(normalized);
  return Number.isFinite(n) ? (negative?-n:n) : null;
}
function hasTrct(...v){ return /\bTRCT\b|\bRESCISAO\b|\bDEMITID[OA]\b|\bDEMISSAO\b/.test(norm(v.join(' '))); }
function hasNovo(...v){ return /NOVO COLABORADOR/.test(norm(v.join(' '))); }
function esc(s){return String(s??'').replace(/[&<>'"]/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch]));}

function msg(text,type='success'){ const e=$('message'); e.textContent=text; e.className='message '+type; }
function hideMsg(){ $('message').className='message hidden'; }
function progress(p,t){ $('progressWrap').classList.remove('hidden'); $('progressBar').style.width=p+'%'; $('progressText').textContent=t; }
function ready(){ $('compareBtn').disabled=!(state.excelFile&&state.pdfFile); }

function setupDrop(inputId,dropId,statusId,key){
  const input=$(inputId), drop=$(dropId), status=$(statusId);
  const set=file=>{ if(!file)return; state[key]=file; status.textContent=file.name+' • '+(file.size/1024/1024).toFixed(2)+' MB'; ready(); hideMsg(); };
  input.addEventListener('change',()=>set(input.files[0]));
  ['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('dragover');}));
  ['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('dragover');}));
  drop.addEventListener('drop',e=>set(e.dataTransfer.files[0]));
}
setupDrop('excelFile','excelDrop','excelStatus','excelFile');
setupDrop('pdfFile','pdfDrop','pdfStatus','pdfFile');

function detectHeader(rows){
  const aliases={
    name:['NOME','COLABORADOR','FUNCIONARIO','EMPREGADO'],
    type:['FUNCAO','VINCULO','TIPO','CATEGORIA'],
    obs:['OBSERVACOES','OBSERVACAO','OBS'],
    net:['REMUNERACAO LIQUIDA A RECEBER','REMUNERACAO LIQUIDA','LIQUIDO','VALOR LIQUIDO','LIQUIDO A RECEBER'],
    cpf:['CPF']
  };
  let best=null;
  for(let r=0;r<Math.min(rows.length,100);r++){
    const rr=rows[r].map(norm), cols={};
    for(const [k,list] of Object.entries(aliases)){ const i=rr.findIndex(v=>list.includes(v)); if(i>=0) cols[k]=i; }
    const score=(cols.name!==undefined?3:0)+(cols.cpf!==undefined?3:0)+(cols.net!==undefined?3:0)+(cols.obs!==undefined?1:0)+(cols.type!==undefined?1:0);
    if(!best||score>best.score) best={row:r,cols,score};
  }
  return best&&best.score>=6?best:null;
}

async function parseExcel(file){
  if(!window.XLSX) throw new Error('Biblioteca de Excel não carregou. Recarregue a página com internet ativa.');
  const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true});
  let chosen=null;
  for(const sheetName of wb.SheetNames){
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,raw:false,defval:''});
    const header=detectHeader(rows);
    if(header&&(!chosen||header.score>chosen.header.score)) chosen={sheetName,rows,header};
  }
  if(!chosen) throw new Error('Não encontrei no Excel um cabeçalho com Nome, CPF e Líquido/Remuneração líquida.');
  const out=[];
  for(let r=chosen.header.row+1;r<chosen.rows.length;r++){
    const row=chosen.rows[r], c=chosen.header.cols;
    const rawName=String(row[c.name]??'').trim();
    if(!rawName||norm(rawName)==='VALOR TOTAL') continue;
    const ccpf=c.cpf!==undefined?cpf(row[c.cpf]):'';
    const nameNorm=normName(rawName);
    if(!nameNorm) continue;
    const obs=c.obs!==undefined?String(row[c.obs]??'').trim():'';
    const type=c.type!==undefined?String(row[c.type]??'').trim():'';
    const netRaw=c.net!==undefined?row[c.net]:'';
    out.push({source:'ORIGINAL',row:r+1,rawName,nameNorm,cpf:ccpf,type,obs,net:money(netRaw),netRaw:String(netRaw??''),isTRCT:hasTrct(rawName,obs),isNew:hasNovo(rawName,obs)});
  }
  return {records:out,sheetName:chosen.sheetName,headerRow:chosen.header.row+1};
}

async function pdfLines(file){
  const pdf=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise;
  const pages=[];
  for(let p=1;p<=pdf.numPages;p++){
    progress(30+Math.round((p/pdf.numPages)*35),'Lendo PDF: página '+p+' de '+pdf.numPages+'...');
    const page=await pdf.getPage(p), tc=await page.getTextContent();
    const items=tc.items.map(x=>({str:String(x.str||'').trim(),x:x.transform?.[4]||0,y:x.transform?.[5]||0})).filter(x=>x.str);
    items.sort((a,b)=>Math.abs(b.y-a.y)>2?b.y-a.y:a.x-b.x);
    const lines=[];
    for(const it of items){
      let line=lines.find(l=>Math.abs(l.y-it.y)<=2);
      if(!line){ line={y:it.y,items:[]}; lines.push(line); }
      line.items.push(it);
    }
    const textLines=lines.sort((a,b)=>b.y-a.y).map(l=>l.items.sort((a,b)=>a.x-b.x).map(i=>i.str).join(' ').replace(/\s+/g,' ').trim());
    pages.push({page:p,lines:textLines});
  }
  return {numPages:pdf.numPages,pages};
}

function parsePdfPage(lines,page){
  const records=[];
  const startRe=/Empr\.:\s*\d+\s+(.+?)\s+Situa[cç][aã]o:\s*([^\s]+(?:\s+[^\s]+)?)\s+CPF:\s*(\d{3}\.\d{3}\.\d{3}-\d{2})/i;
  const simpleRe=/Empr\.:\s*\d+\s+(.+?)\s+Situa/i;
  const cpfRe=/CPF:\s*(\d{3}\.\d{3}\.\d{3}-\d{2})/i;

  for(let i=0;i<lines.length;i++){
    if(!/Empr\.:/i.test(lines[i])) continue;
    let header=lines[i], j=i+1;
    while(j<lines.length && j<i+3 && !/CPF:/i.test(header)){ header+=' '+lines[j]; j++; }
    let m=header.match(startRe), rawName='', situation='', ccpf='';
    if(m){ rawName=m[1].trim(); situation=m[2].trim(); ccpf=cpf(m[3]); }
    else {
      const nm=header.match(simpleRe), cm=header.match(cpfRe);
      rawName=nm?nm[1].trim():''; ccpf=cm?cpf(cm[1]):'';
    }
    if(!rawName||!ccpf) continue;

    let end=i+1;
    const isSummaryBoundary=line=>/^(RESUMO POR RUBRICA|RESUMO POR CENTRO|RESUMO GERAL|TOTAIS? DA FOLHA|TOTALIZACAO|BASES? DA FOLHA)/.test(norm(line));
    while(end<lines.length && !/Empr\.:/i.test(lines[end]) && !isSummaryBoundary(lines[end])) end++;
    const block=lines.slice(i,end);
    const text=block.join(' ');
    const vinc=(text.match(/V[ií]nculo:\s*(.+?)(?:\s+CC:|\s+Depto:|\s+Horas M[eê]s:)/i)||[])[1]?.trim()||'';

    let net=null;
    for(let k=0;k<block.length;k++){
      if(/L[ií]quido:/i.test(block[k])){
        const vals=[...block[k].matchAll(/-?\d{1,3}(?:\.\d{3})*,\d{2}|-?\d+,\d{2}/g)].map(x=>money(x[0])).filter(v=>v!==null);
        if(vals.length) net=vals[vals.length-1];
        if(net===null && k+1<block.length){
          const next=[...block[k+1].matchAll(/-?\d{1,3}(?:\.\d{3})*,\d{2}|-?\d+,\d{2}/g)].map(x=>money(x[0])).filter(v=>v!==null);
          if(next.length) net=next[0];
        }
      }
    }

    let trctValue=null;
    for(const line of block){
      if(/LIQUIDO\s+RESCISAO/i.test(norm(line))){
        const vals=[...line.matchAll(/-?\d{1,3}(?:\.\d{3})*,\d{2}|-?\d+,\d{2}/g)].map(x=>money(x[0])).filter(v=>v!==null&&v>0);
        if(vals.length) trctValue=Math.max(...vals);
      }
    }
    if(trctValue===null && hasTrct(text)){
      const rub=block.findIndex(x=>/LIQUIDO\s+RESCISAO/i.test(norm(x)));
      if(rub>=0){
        for(let k=Math.max(0,rub-2);k<=Math.min(block.length-1,rub+2);k++){
          const vals=[...block[k].matchAll(/-?\d{1,3}(?:\.\d{3})*,\d{2}|-?\d+,\d{2}/g)].map(x=>money(x[0])).filter(v=>v!==null&&v>1);
          if(vals.length){ trctValue=Math.max(...vals); break; }
        }
      }
    }

    records.push({source:'FOLHA',page,rawName,nameNorm:normName(rawName),cpf:ccpf,type:vinc,obs:situation,net,trctValue,isTRCT:hasTrct(text,situation)||trctValue!==null});
    i=end-1;
  }
  return records;
}

async function parsePdf(file){
  const data=await pdfLines(file);
  let records=[];
  for(const pg of data.pages) records.push(...parsePdfPage(pg.lines,pg.page));
  return {records,pages:data.numPages};
}

function mapPush(map,key,obj){ if(!key)return; if(!map.has(key))map.set(key,[]); map.get(key).push(obj); }
function result(o,f,status,matchKey,diff){
  return {status,cpf:o?.cpf||f?.cpf||'',name:o?.rawName||f?.rawName||'',type:o?.type||f?.type||'',obs:o?.obs||'',
    originalNet:o?.net??null,folhaNet:f?.net??null,trctNet:f?.trctValue??null,diff,matchKey,originalRow:o?.row??'',pdfPage:f?.page??''};
}
function rank(s){ const a=[STATUS.DIV,STATUS.CENT,STATUS.VER,STATUS.DUP,STATUS.SO_ORIG,STATUS.SO_FOLHA,STATUS.TRCT,STATUS.NOVO,STATUS.OK]; return a.indexOf(s); }

function compare(original,folha,tol,cents){
  const cp=new Map(), nm=new Map();
  folha.forEach(f=>{mapPush(cp,f.cpf,f);mapPush(nm,f.nameNorm,f);});
  const used=new Set(), out=[];
  for(const o of original){
    let matches=[], key='';
    if(o.cpf&&cp.has(o.cpf)){matches=cp.get(o.cpf);key='CPF';}
    else if(o.nameNorm&&nm.has(o.nameNorm)){matches=nm.get(o.nameNorm);key='NOME';}

    if(matches.length>1){out.push(result(o,null,STATUS.DUP,key,null));continue;}
    if(!matches.length){out.push(result(o,null,o.isNew?STATUS.NOVO:STATUS.SO_ORIG,'',null));continue;}
    const f=matches[0]; used.add(f);
    if(o.isTRCT||f.isTRCT){
      const target=f.trctValue??f.net;
      const diff=o.net!==null&&target!==null?target-o.net:null;
      out.push(result(o,f,STATUS.TRCT,key,diff)); continue;
    }
    const diff=o.net!==null&&f.net!==null?f.net-o.net:null;
    if(o.isNew){out.push(result(o,f,STATUS.NOVO,key,diff));continue;}
    if(diff===null){out.push(result(o,f,STATUS.VER,key,null));continue;}
    const ad=Math.abs(diff);
    out.push(result(o,f,ad<=tol?STATUS.OK:(ad<=cents?STATUS.CENT:STATUS.DIV),key,diff));
  }
  for(const f of folha) if(!used.has(f)) out.push(result(null,f,STATUS.SO_FOLHA,'',null));
  return out.sort((a,b)=>rank(a.status)-rank(b.status)||a.name.localeCompare(b.name,'pt-BR'));
}

function statusClass(s){
  if(s===STATUS.OK)return'ok'; if(s===STATUS.DIV)return'divergencia'; if(s===STATUS.CENT)return'centavos';
  if(s===STATUS.TRCT)return'trct'; if(s===STATUS.NOVO)return'novo'; if(s===STATUS.DUP)return'duplicado'; if(s===STATUS.VER)return'verificar'; return'ausente';
}
function renderSummary(){
  const r=state.results, count=s=>r.filter(x=>x.status===s).length;
  const sum=(arr,k)=>arr.reduce((a,x)=>a+(Number.isFinite(x[k])?x[k]:0),0);
  const validOriginal=r.filter(x=>x.status!==STATUS.SO_FOLHA && Number.isFinite(x.originalNet));
  const comparable=r.filter(x=>Number.isFinite(x.originalNet)&&Number.isFinite(x.folhaNet)&&x.status!==STATUS.TRCT);
  const originalTotal=sum(validOriginal,'originalNet');
  const folhaTotal=sum(r.filter(x=>x.status!==STATUS.SO_ORIG),'folhaNet');
  const trctTotal=sum(r,'trctNet');
  const eff=comparable.filter(x=>Number.isFinite(x.diff)).reduce((a,x)=>a+x.diff,0);
  const canCompareMoney=comparable.length>0;
  const metrics=[
    ['Total original',validOriginal.length?moneyFmt.format(originalTotal):'N/D','info'],['Total folha (mensal)',moneyFmt.format(folhaTotal),'info'],
    ['Total TRCT extraído',moneyFmt.format(trctTotal),'warn'],['Diferença efetiva',canCompareMoney?moneyFmt.format(eff):'N/D',canCompareMoney?(Math.abs(eff)<=Number($('tolerance').value||.01)?'good':'bad'):'warn'],
    ['Colaboradores OK',count(STATUS.OK),'good'],['Divergências de líquido',canCompareMoney?(count(STATUS.DIV)+count(STATUS.CENT)):'N/D',canCompareMoney?((count(STATUS.DIV)+count(STATUS.CENT))>0?'critical':'good'):'warn'],
    ['Ausentes',count(STATUS.SO_ORIG)+count(STATUS.SO_FOLHA),'warn'],['Pendentes / duplicados',count(STATUS.VER)+count(STATUS.DUP),'warn'],
    ['TRCT / rescisões',count(STATUS.TRCT),'warn'],['Novos colaboradores',count(STATUS.NOVO),'info']
  ];
  $('summaryCards').innerHTML=metrics.map(([l,v,c])=>'<div class="metric '+c+'"><span class="label">'+esc(l)+'</span><span class="value">'+esc(v)+'</span></div>').join('');
  const divCount=count(STATUS.DIV)+count(STATUS.CENT);
  const centsCount=count(STATUS.CENT);
  const alert=$('resultAlert');
  alert.classList.remove('hidden','has-div','no-div','no-data');
  if(!canCompareMoney){
    alert.classList.add('no-data');
    alert.innerHTML='⚠ Comparação de líquido não realizada.'+
      '<span class="sub">O Excel original não contém valores preenchidos em “Remuneração líquida a receber”. CPF, nome, TRCT, novos colaboradores e presença/ausência continuam sendo validados, mas não é correto afirmar que há 0 divergências de líquido.</span>';
  }else if(divCount>0){
    alert.classList.add('has-div');
    alert.innerHTML='⚠ Foram encontradas <strong>'+divCount+'</strong> divergência(s) de líquido.'+
      '<span class="sub">'+centsCount+' classificada(s) como diferença de centavos. Use “Mostrar só divergências” para revisar apenas esses casos.</span>';
  }else{
    alert.classList.add('no-div');
    alert.innerHTML='✓ Nenhuma divergência de líquido encontrada dentro dos critérios atuais.'+
      '<span class="sub">TRCT, novos colaboradores, ausências, duplicidades e itens “Verificar” continuam sendo mostrados separadamente.</span>';
  }
}
function renderResults(){
  const q=norm($('searchInput').value), sf=$('statusFilter').value;
  const rows=state.results.filter(r=>(sf==='ALL'||r.status===sf)&&(!q||norm(r.name+' '+r.cpf).includes(q)));
  $('resultCount').textContent=rows.length+' de '+state.results.length+' registros exibidos';
  $('resultsBody').innerHTML=rows.map(r=>'<tr class="row-'+statusClass(r.status)+'">'+
    '<td><span class="status '+statusClass(r.status)+'">'+esc(r.status)+'</span></td>'+
    '<td>'+esc(fmtCpf(r.cpf))+'</td><td>'+esc(r.name)+'</td><td>'+esc(r.type||'—')+'</td><td>'+esc(r.obs||'—')+'</td>'+
    '<td class="money">'+(r.originalNet===null?'—':moneyFmt.format(r.originalNet))+'</td>'+
    '<td class="money">'+(r.folhaNet===null?'—':moneyFmt.format(r.folhaNet))+'</td>'+
    '<td class="money">'+(r.trctNet===null?'—':moneyFmt.format(r.trctNet))+'</td>'+
    '<td class="money">'+(r.diff===null?'—':moneyFmt.format(r.diff))+'</td><td>'+esc(r.matchKey||'—')+'</td></tr>').join('');
}
function renderDiag(){
  const d=state.diagnostics;
  $('diagnostics').innerHTML='<div><strong>Excel:</strong> aba “'+esc(d.sheetName||'')+'”, cabeçalho na linha '+(d.headerRow||'—')+', '+(d.originalCount||0)+' registros importados.</div>'+
  '<div><strong>PDF:</strong> '+(d.pdfPages||0)+' páginas, '+(d.folhaCount||0)+' colaboradores identificados.</div>'+
  '<div><strong>Chave:</strong> CPF primeiro; nome normalizado como alternativa.</div>'+
  '<div><strong>TRCT:</strong> rescisões são tratadas separadamente do líquido mensal.</div>';
}
function exportXlsx(){
  const comparableCount=state.results.filter(r=>Number.isFinite(r.originalNet)&&Number.isFinite(r.folhaNet)&&r.status!==STATUS.TRCT).length;
  if(!comparableCount){
    msg('O Excel original não possui valores de líquido preenchidos. Não há comparação monetária válida para exportar.','error');
    return;
  }
  const rows=state.results.filter(r=>r.status===STATUS.DIV||r.status===STATUS.CENT).map(r=>({
    Status:r.status,CPF:fmtCpf(r.cpf),Nome:r.name,'Vínculo / Tipo':r.type,'Observação original':r.obs,
    'Líquido original':r.originalNet,'Líquido folha':r.folhaNet,'TRCT folha':r.trctNet,'Diferença':r.diff,
    'Chave usada':r.matchKey,'Linha Excel':r.originalRow,'Página PDF':r.pdfPage
  }));
  if(!rows.length){
    msg('Não há divergências de líquido para exportar.','success');
    return;
  }
  const wb=XLSX.utils.book_new(), ws=XLSX.utils.json_to_sheet(rows);
  ws['!cols']=[18,16,38,22,28,16,16,16,16,20,12,12].map(wch=>({wch}));
  XLSX.utils.book_append_sheet(wb,ws,'Divergências');
  XLSX.writeFile(wb,'Divergencias_Folha_'+new Date().toISOString().slice(0,10)+'.xlsx');
}

async function run(){
  hideMsg(); $('summarySection').classList.add('hidden'); $('resultsSection').classList.add('hidden'); $('diagnosticsSection').classList.add('hidden');
  $('compareBtn').disabled=true;
  try{
    progress(5,'Lendo Excel...');
    const ex=await parseExcel(state.excelFile); state.original=ex.records;
    progress(25,'Excel lido. Iniciando PDF...');
    const pf=await parsePdf(state.pdfFile); state.folha=pf.records;
    progress(75,'Comparando CPF, nomes e valores...');
    const tol=Number($('tolerance').value||.01), cents=Math.max(tol,Number($('centsLimit').value||1));
    state.results=compare(state.original,state.folha,tol,cents);
    state.diagnostics={sheetName:ex.sheetName,headerRow:ex.headerRow,originalCount:ex.records.length,pdfPages:pf.pages,folhaCount:pf.records.length};
    renderSummary(); renderResults(); renderDiag();
    $('summarySection').classList.remove('hidden'); $('resultsSection').classList.remove('hidden'); $('diagnosticsSection').classList.remove('hidden');
    progress(100,'Concluído.'); setTimeout(()=>$('progressWrap').classList.add('hidden'),500);
    msg('Comparação concluída: '+state.original.length+' registros no Excel e '+state.folha.length+' colaboradores identificados no PDF.');
  }catch(e){ console.error(e); $('progressWrap').classList.add('hidden'); msg(e?.message||'Erro ao processar os arquivos.','error'); }
  finally{ ready(); }
}
function reset(){
  Object.assign(state,{excelFile:null,pdfFile:null,original:[],folha:[],results:[],diagnostics:{}});
  $('excelFile').value='';$('pdfFile').value='';$('excelStatus').textContent='Nenhum arquivo selecionado.';$('pdfStatus').textContent='Nenhum arquivo selecionado.';
  $('summarySection').classList.add('hidden');$('resultsSection').classList.add('hidden');$('diagnosticsSection').classList.add('hidden');hideMsg();$('progressWrap').classList.add('hidden');ready();
}
$('compareBtn').addEventListener('click',run);
$('resetBtn').addEventListener('click',reset);
$('exportBtn').addEventListener('click',exportXlsx);
$('searchInput').addEventListener('input',renderResults);
$('statusFilter').addEventListener('change',renderResults);
$('divergenceOnlyBtn').addEventListener('click',()=>{
  $('statusFilter').value='ALL';
  const only=state.results.filter(r=>r.status===STATUS.DIV||r.status===STATUS.CENT);
  const q=norm($('searchInput').value);
  const rows=only.filter(r=>!q||norm(r.name+' '+r.cpf).includes(q));
  $('resultCount').textContent=rows.length+' divergência(s) de líquido exibida(s)';
  $('resultsBody').innerHTML=rows.map(r=>'<tr class="row-'+statusClass(r.status)+'">'+
    '<td><span class="status '+statusClass(r.status)+'">'+esc(r.status)+'</span></td>'+
    '<td>'+esc(fmtCpf(r.cpf))+'</td><td>'+esc(r.name)+'</td><td>'+esc(r.type||'—')+'</td><td>'+esc(r.obs||'—')+'</td>'+
    '<td class="money">'+(r.originalNet===null?'—':moneyFmt.format(r.originalNet))+'</td>'+
    '<td class="money">'+(r.folhaNet===null?'—':moneyFmt.format(r.folhaNet))+'</td>'+
    '<td class="money">'+(r.trctNet===null?'—':moneyFmt.format(r.trctNet))+'</td>'+
    '<td class="money">'+(r.diff===null?'—':moneyFmt.format(r.diff))+'</td><td>'+esc(r.matchKey||'—')+'</td></tr>').join('');
});
