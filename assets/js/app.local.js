import * as pdfjsLib from '../vendor/pdfjs-4/pdf.min.mjs';
import { FIELD_SCHEMA, parseFieldPdf, auditFields, auditSummary } from './field-audit.js';
pdfjsLib.GlobalWorkerOptions.workerSrc = 'assets/vendor/pdfjs-4/pdf.worker.min.mjs';

const $ = id => document.getElementById(id);
const state = { excelFile:null, pdfFile:null, original:[], folha:[], results:[], diagnostics:{}, auditRows:[], auditPdf:[], auditLimit:200, fill:{excelFile:null,pdfFile:null,excel:null,pdf:null,rows:[],competence:''} };

const STATUS = {
  OK:'OK', DIV:'DIVERGÊNCIA', CENT:'DIFERENÇA DE CENTAVOS', TRCT:'TRCT',
  NOVO:'NOVO COLABORADOR', SO_ORIG:'SOMENTE NO ORIGINAL', SO_FOLHA:'SOMENTE NA FOLHA',
  DUP:'DUPLICADO', VER:'VERIFICAR', FILL:'LÍQUIDO VAZIO NO EXCEL'
};
const moneyFmt = new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'});
const HISTORY_KEY = 'comparadorFolhaHistoricoV1';

function switchModule(mode){
  const compare=mode==='compare';
  $('compareModule').classList.toggle('hidden',!compare);
  $('fillModule').classList.toggle('hidden',compare);
  $('tabCompare').classList.toggle('active',compare);
  $('tabFill').classList.toggle('active',!compare);
}
$('tabCompare').addEventListener('click',()=>switchModule('compare'));
$('tabFill').addEventListener('click',()=>switchModule('fill'));


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
    const auditCells={};
    for(const field of FIELD_SCHEMA){const cell=wb.Sheets[chosen.sheetName][field.col+(r+1)];const header=String(chosen.rows[chosen.header.row][colIndexFromRef(field.col)]??'');
      if(header)auditCells[field.col]={header,value:cell?.v??null,display:cell?.w||String(row[colIndexFromRef(field.col)]??'')};
    }
    out.push({source:'ORIGINAL',auditCells,row:r+1,rawName,nameNorm,cpf:ccpf,type,obs,net:money(netRaw),netRaw:String(netRaw??''),netCell:c.net!==undefined?colLetter(c.net)+(r+1):'',isTRCT:hasTrct(rawName,obs),isNew:hasNovo(rawName,obs)});
  }
  return {records:out,sheetName:chosen.sheetName,headerRow:chosen.header.row+1};
}

async function pdfLines(file,onProgress=progress){
  const pdf=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise;
  const pages=[];
  for(let p=1;p<=pdf.numPages;p++){
    onProgress(30+Math.round((p/pdf.numPages)*35),'Lendo PDF: página '+p+' de '+pdf.numPages+'...');
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

    records.push({source:'FOLHA',fields:parseFieldPdf(block),page,rawName,nameNorm:normName(rawName),cpf:ccpf,type:vinc,obs:situation,net,trctValue,isTRCT:hasTrct(text,situation)||trctValue!==null});
    i=end-1;
  }
  return records;
}

async function parsePdf(file,onProgress=progress){
  const data=await pdfLines(file,onProgress);
  let records=[];
  for(const pg of data.pages) records.push(...parsePdfPage(pg.lines,pg.page));
  const flat=data.pages.flatMap(pg=>pg.lines);
  // Emissão e admissão também contêm datas: somente o campo Competência identifica a folha.
  const comp=(flat.join(' ').match(/Compet[eê]ncia:\s*((?:0[1-9]|1[0-2])\/20\d{2})\b/i)||[])[1]||'';
  if(!records.length) throw new Error('Nenhum colaborador identificado no PDF. Confira se é um extrato mensal com texto selecionável.');
  return {records,pages:data.numPages,competence:comp};
}

function mapPush(map,key,obj){ if(!key)return; if(!map.has(key))map.set(key,[]); map.get(key).push(obj); }

function fillMsg(text,type='success'){
  const e=$('fillMessage'); e.textContent=text; e.className='message '+type;
}
function fillHideMsg(){ $('fillMessage').className='message hidden'; }
function fillProgress(p,t){
  $('fillProgressWrap').classList.remove('hidden');
  $('fillProgressBar').style.width=p+'%';
  $('fillProgressText').textContent=t;
}
function fillReady(){ $('fillAnalyzeBtn').disabled=!(state.fill.excelFile&&state.fill.pdfFile); }

function setupFillDrop(inputId,dropId,statusId,key,validator){
  const input=$(inputId), drop=$(dropId), status=$(statusId);
  const set=file=>{
    if(!file)return;
    if(validator && !validator(file)){
      fillMsg('Selecione um arquivo .xlsx válido para preservar a estrutura da planilha.','error');
      input.value=''; return;
    }
    state.fill[key]=file;
    status.textContent=file.name+' • '+(file.size/1024/1024).toFixed(2)+' MB';
    fillReady(); fillHideMsg();
  };
  input.addEventListener('change',()=>set(input.files[0]));
  ['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('dragover');}));
  ['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('dragover');}));
  drop.addEventListener('drop',e=>set(e.dataTransfer.files[0]));
}
setupFillDrop('fillExcelFile','fillExcelDrop','fillExcelStatus','excelFile',f=>/\.xlsx$/i.test(f.name));
setupFillDrop('fillPdfFile','fillPdfDrop','fillPdfStatus','pdfFile',f=>/\.pdf$/i.test(f.name)||f.type==='application/pdf');

function colLetter(idx){
  let n=idx+1,s='';
  while(n){ const r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26); }
  return s;
}
function colIndexFromRef(ref=''){
  const m=String(ref).match(/^([A-Z]+)/i); if(!m)return -1;
  let n=0; for(const ch of m[1].toUpperCase()) n=n*26+(ch.charCodeAt(0)-64);
  return n-1;
}

async function parseFillExcel(file){
  if(!window.XLSX) throw new Error('Biblioteca de Excel não carregou. Recarregue a página com internet ativa.');
  const buffer=await file.arrayBuffer();
  const wb=XLSX.read(buffer,{type:'array',cellDates:true,cellStyles:true,cellFormula:true});
  let chosen=null;
  for(const sheetName of wb.SheetNames){
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,raw:false,defval:''});
    const header=detectHeader(rows);
    if(header&&header.cols.net!==undefined&&(!chosen||header.score>chosen.header.score)) chosen={sheetName,rows,header};
  }
  if(!chosen) throw new Error('Não encontrei no Excel um cabeçalho com Nome, CPF e “Remuneração líquida a receber”.');
  const c=chosen.header.cols, out=[];
  for(let r=chosen.header.row+1;r<chosen.rows.length;r++){
    const row=chosen.rows[r], rawName=String(row[c.name]??'').trim();
    if(!rawName||norm(rawName)==='VALOR TOTAL') continue;
    const nameNorm=normName(rawName); if(!nameNorm)continue;
    const ccpf=c.cpf!==undefined?cpf(row[c.cpf]):'';
    const obs=c.obs!==undefined?String(row[c.obs]??'').trim():'';
    out.push({
      row:r+1, rawName, nameNorm, cpf:ccpf, obs,
      currentNet:money(row[c.net]),
      isTRCT:hasTrct(rawName,obs),
      cell:colLetter(c.net)+(r+1)
    });
  }
  return {
    records:out, sheetName:chosen.sheetName, headerRow:chosen.header.row+1,
    netColIndex:c.net, netColLetter:colLetter(c.net), buffer
  };
}

function fillStatusClass(s){
  if(s==='PREENCHER')return'preencher';
  if(s==='JÁ CORRETO')return'correto';
  if(s==='TRCT')return'trct';
  if(s==='NÃO ENCONTRADO')return'naoencontrado';
  if(s==='SOMENTE NO PDF')return'somentepdf';
  if(s==='DUPLICADO')return'duplicado';
  return'verificar';
}

function prepareFillRows(excelRecords,pdfRecords){
  const cp=new Map(),nm=new Map();
  pdfRecords.forEach(f=>{mapPush(cp,f.cpf,f);mapPush(nm,f.nameNorm,f);});
  const excelCpfCount=new Map();
  excelRecords.forEach(e=>{if(e.cpf)excelCpfCount.set(e.cpf,(excelCpfCount.get(e.cpf)||0)+1);});
  const used=new Set(), out=[];
  for(const e of excelRecords){
    if(e.cpf && (excelCpfCount.get(e.cpf)||0)>1){
      out.push({...e,status:'DUPLICADO',pdfNet:null,trctNet:null,target:null,rule:'CPF duplicado no Excel',pdfPage:''});
      continue;
    }
    let matches=[],matchKey='';
    if(e.cpf&&cp.has(e.cpf)){matches=cp.get(e.cpf);matchKey='CPF';}
    else if(e.nameNorm&&nm.has(e.nameNorm)){matches=nm.get(e.nameNorm);matchKey='NOME';}
    if(matches.length>1){
      out.push({...e,status:'DUPLICADO',pdfNet:null,trctNet:null,target:null,rule:'Mais de um registro no PDF para a mesma chave',pdfPage:''});
      continue;
    }
    if(!matches.length){
      out.push({...e,status:'NÃO ENCONTRADO',pdfNet:null,trctNet:null,target:null,rule:'Colaborador não localizado no PDF',pdfPage:''});
      continue;
    }
    const f=matches[0];
    if(used.has(f)){
      out.push({...e,status:'DUPLICADO',pdfNet:f.net,trctNet:f.trctValue,target:null,rule:'Registro do PDF já associado a outro colaborador',pdfPage:f.page});
      continue;
    }
    used.add(f);
    const isTrct=e.isTRCT||f.isTRCT||f.trctValue!==null;
    const target=isTrct && f.trctValue!==null ? f.trctValue : f.net;
    if(target===null){
      out.push({...e,status:'VERIFICAR',pdfNet:f.net,trctNet:f.trctValue,target:null,rule:'Não foi possível identificar o valor líquido',pdfPage:f.page,matchKey});
      continue;
    }
    const equal=e.currentNet!==null&&Math.abs(e.currentNet-target)<=0.005;
    const status=isTrct?'TRCT':(equal?'JÁ CORRETO':'PREENCHER');
    out.push({...e,status,pdfNet:f.net,trctNet:f.trctValue,target,
      rule:isTrct&&f.trctValue!==null?'Líquido da rescisão / TRCT':'Líquido mensal',
      pdfPage:f.page,matchKey});
  }
  for(const f of pdfRecords){
    if(!used.has(f)){
      out.push({row:'',cell:'',rawName:f.rawName,nameNorm:f.nameNorm,cpf:f.cpf,currentNet:null,isTRCT:f.isTRCT,
        status:'SOMENTE NO PDF',pdfNet:f.net,trctNet:f.trctValue,target:null,rule:'Existe no PDF, mas não foi localizado no Excel',pdfPage:f.page,matchKey:''});
    }
  }
  return out;
}

function renderFill(){
  const rows=state.fill.rows;
  const count=s=>rows.filter(r=>r.status===s).length;
  const writeable=rows.filter(r=>Number.isFinite(r.target)&&r.row);
  const unresolved=rows.filter(r=>['NÃO ENCONTRADO','DUPLICADO','VERIFICAR','SOMENTE NO PDF'].includes(r.status));
  const changes=writeable.filter(r=>r.currentNet===null||Math.abs(r.currentNet-r.target)>0.005);
  $('fillSummaryCards').innerHTML=[
    ['Colaboradores no Excel',state.fill.excel?.records.length||0,'info'],
    ['Valores identificados',writeable.length,'good'],
    ['Alterações necessárias',changes.length,changes.length?'info':'good'],
    ['TRCT / rescisões',count('TRCT'),'warn'],
    ['Já corretos',count('JÁ CORRETO'),'good'],
    ['Pendências',unresolved.length,unresolved.length?'bad':'good'],
    ['Colaboradores no PDF',state.fill.pdf?.records.length||0,'info'],
    ['Competência',state.fill.competence||'—','info']
  ].map(([l,v,c])=>'<div class="metric '+c+'"><span class="label">'+esc(l)+'</span><span class="value">'+esc(v)+'</span></div>').join('');

  const alert=$('fillAlert');
  alert.classList.remove('hidden','has-div','no-div','no-data');
  if(unresolved.length){
    alert.classList.add('has-div');
    alert.innerHTML='⚠ Há <strong>'+unresolved.length+'</strong> registro(s) que exigem conferência.'+
      '<span class="sub">O Excel ainda pode ser gerado; somente registros com correspondência segura e valor identificado serão preenchidos. Pendências permanecem intactas.</span>';
  }else{
    alert.classList.add('no-div');
    alert.innerHTML='✓ Todos os registros do Excel possuem correspondência segura no PDF.'+
      '<span class="sub">Confira a tabela abaixo e gere a nova cópia quando estiver de acordo.</span>';
  }
  $('fillGenerateBtn').disabled=writeable.length===0;
  renderFillPreview();
  $('fillDiagnostics').innerHTML=
    '<div><strong>Excel:</strong> aba “'+esc(state.fill.excel?.sheetName||'')+'”, cabeçalho na linha '+(state.fill.excel?.headerRow||'—')+', coluna de destino <strong>'+esc(state.fill.excel?.netColLetter||'—')+'</strong>.</div>'+
    '<div><strong>PDF:</strong> '+(state.fill.pdf?.pages||0)+' páginas, '+(state.fill.pdf?.records.length||0)+' colaboradores identificados.</div>'+
    '<div><strong>Chave:</strong> CPF primeiro; nome normalizado apenas como alternativa.</div>'+
    '<div><strong>TRCT:</strong> quando existe rubrica “Líquido Rescisão”, ela substitui o líquido mensal para o valor a gravar.</div>'+
    '<div><strong>Saída:</strong> o pacote .xlsx original é preservado; apenas as células de destino são atualizadas.</div>';
}

function renderFillPreview(){
  const q=norm($('fillSearchInput').value), sf=$('fillStatusFilter').value;
  const rows=state.fill.rows.filter(r=>(sf==='ALL'||r.status===sf)&&(!q||norm((r.rawName||'')+' '+(r.cpf||'')).includes(q)));
  $('fillPreviewCount').textContent=rows.length+' de '+state.fill.rows.length+' registros exibidos';
  $('fillPreviewBody').innerHTML=rows.map(r=>'<tr class="row-'+fillStatusClass(r.status)+'">'+
    '<td><span class="status '+fillStatusClass(r.status)+'">'+esc(r.status)+'</span></td>'+
    '<td>'+esc(fmtCpf(r.cpf))+'</td><td>'+esc(r.rawName||'—')+'</td>'+
    '<td class="money">'+(r.currentNet===null?'—':moneyFmt.format(r.currentNet))+'</td>'+
    '<td class="money">'+(r.pdfNet===null?'—':moneyFmt.format(r.pdfNet))+'</td>'+
    '<td class="money">'+(r.trctNet===null?'—':moneyFmt.format(r.trctNet))+'</td>'+
    '<td class="money"><strong>'+(r.target===null?'—':moneyFmt.format(r.target))+'</strong></td>'+
    '<td>'+esc(r.rule||'—')+'</td><td>'+esc(r.cell||'—')+'</td></tr>').join('');
}

async function runFill(){
  fillHideMsg();
  ['fillSummarySection','fillPreviewSection','fillDiagnosticsSection'].forEach(id=>$(id).classList.add('hidden'));
  $('fillAnalyzeBtn').disabled=true;
  try{
    fillProgress(5,'Lendo estrutura do Excel...');
    const ex=await parseFillExcel(state.fill.excelFile);
    fillProgress(22,'Excel identificado. Lendo PDF...');
    const pf=await parsePdf(state.fill.pdfFile,(p,t)=>fillProgress(Math.max(25,p),t));
    fillProgress(78,'Relacionando CPF, nomes e líquidos...');
    state.fill.excel=ex; state.fill.pdf=pf; state.fill.competence=pf.competence||'';
    state.fill.rows=prepareFillRows(ex.records,pf.records);
    renderFill();
    ['fillSummarySection','fillPreviewSection','fillDiagnosticsSection'].forEach(id=>$(id).classList.remove('hidden'));
    fillProgress(100,'Conferência pronta.');
    setTimeout(()=>$('fillProgressWrap').classList.add('hidden'),500);
    fillMsg('Análise concluída. Confira os valores antes de gerar o Excel preenchido.');
  }catch(e){
    console.error(e); $('fillProgressWrap').classList.add('hidden');
    fillMsg(e?.message||'Erro ao preparar o preenchimento.','error');
  }finally{ fillReady(); }
}

function findXmlByLocalName(root,name){ return Array.from(root.getElementsByTagName('*')).find(x=>x.localName===name)||null; }
function findAllXmlByLocalName(root,name){ return Array.from(root.getElementsByTagName('*')).filter(x=>x.localName===name); }

async function patchXlsxNetValues(file,sheetName,updates){
  if(!window.JSZip) throw new Error('Biblioteca de preservação do Excel não carregou. Recarregue a página com internet ativa.');
  const zip=await JSZip.loadAsync(await file.arrayBuffer());
  const wbEntry=zip.file('xl/workbook.xml'), relEntry=zip.file('xl/_rels/workbook.xml.rels');
  if(!wbEntry||!relEntry) throw new Error('Estrutura interna do .xlsx não reconhecida.');
  const parser=new DOMParser();
  const wbDoc=parser.parseFromString(await wbEntry.async('text'),'application/xml');
  const relDoc=parser.parseFromString(await relEntry.async('text'),'application/xml');
  if(wbDoc.querySelector('parsererror')||relDoc.querySelector('parsererror')) throw new Error('Não foi possível ler a estrutura interna do Excel.');

  const sheetEl=findAllXmlByLocalName(wbDoc,'sheet').find(x=>x.getAttribute('name')===sheetName);
  if(!sheetEl) throw new Error('A aba de destino não foi encontrada no arquivo.');
  const rid=sheetEl.getAttribute('r:id')||sheetEl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
  const rel=findAllXmlByLocalName(relDoc,'Relationship').find(x=>x.getAttribute('Id')===rid);
  if(!rel) throw new Error('Não foi possível localizar o XML da aba de destino.');
  let target=rel.getAttribute('Target')||'';
  target=target.replace(/^\//,'');
  if(!target.startsWith('xl/')) target='xl/'+target.replace(/^\.\//,'').replace(/^\.\.\//,'');
  const sheetEntry=zip.file(target);
  if(!sheetEntry) throw new Error('Arquivo interno da aba não encontrado: '+target);

  const sheetDoc=parser.parseFromString(await sheetEntry.async('text'),'application/xml');
  if(sheetDoc.querySelector('parsererror')) throw new Error('Não foi possível interpretar a aba do Excel.');
  const sheetData=findXmlByLocalName(sheetDoc,'sheetData');
  if(!sheetData) throw new Error('Área de dados da aba não encontrada.');
  const ns=sheetDoc.documentElement.namespaceURI||'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

  for(const u of updates){
    const ref=u.cell;
    let rowEl=findAllXmlByLocalName(sheetData,'row').find(x=>Number(x.getAttribute('r'))===Number(u.row));
    if(!rowEl){
      rowEl=sheetDoc.createElementNS(ns,'row'); rowEl.setAttribute('r',String(u.row));
      const rows=findAllXmlByLocalName(sheetData,'row');
      const next=rows.find(x=>Number(x.getAttribute('r'))>Number(u.row));
      if(next)sheetData.insertBefore(rowEl,next); else sheetData.appendChild(rowEl);
    }
    let cell=findAllXmlByLocalName(rowEl,'c').find(x=>x.getAttribute('r')===ref);
    if(!cell){
      cell=sheetDoc.createElementNS(ns,'c'); cell.setAttribute('r',ref);
      const targetCol=colIndexFromRef(ref);
      const cells=findAllXmlByLocalName(rowEl,'c');
      const next=cells.find(x=>colIndexFromRef(x.getAttribute('r'))>targetCol);
      if(next)rowEl.insertBefore(cell,next); else rowEl.appendChild(cell);
    }
    cell.removeAttribute('t');
    for(const child of [...cell.children]){
      if(['f','v','is'].includes(child.localName)) cell.removeChild(child);
    }
    const v=sheetDoc.createElementNS(ns,'v');
    v.textContent=String(Number(u.target));
    cell.appendChild(v);
  }
  zip.file(target,new XMLSerializer().serializeToString(sheetDoc));
  return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',compression:'DEFLATE'});
}

function downloadBlob(blob,name){
  const a=document.createElement('a'),url=URL.createObjectURL(blob);
  a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1200);
}

async function generateFilledExcel(){
  try{
    const updates=state.fill.rows.filter(r=>Number.isFinite(r.target)&&r.row&&r.cell);
    if(!updates.length) throw new Error('Não há valores seguros para preencher.');
    $('fillGenerateBtn').disabled=true;
    fillProgress(10,'Preservando estrutura do Excel e atualizando líquidos...');
    const blob=await patchXlsxNetValues(state.fill.excelFile,state.fill.excel.sheetName,updates);
    fillProgress(95,'Preparando arquivo final...');
    const base=state.fill.excelFile.name.replace(/\.xlsx$/i,'');
    const comp=state.fill.competence?'_'+state.fill.competence.replace('/','-'):'';
    downloadBlob(blob,base+comp+'_PREENCHIDO.xlsx');
    fillProgress(100,'Arquivo gerado.');
    setTimeout(()=>$('fillProgressWrap').classList.add('hidden'),500);
    fillMsg('Excel preenchido gerado com sucesso. O arquivo original permaneceu intacto.');
  }catch(e){
    console.error(e); $('fillProgressWrap').classList.add('hidden');
    fillMsg(e?.message||'Erro ao gerar o Excel preenchido.','error');
  }finally{ $('fillGenerateBtn').disabled=false; }
}

function exportFillAudit(){
  if(!state.fill.rows.length){fillMsg('Faça a análise antes de exportar a conferência.','error');return;}
  const rows=state.fill.rows.map(r=>({
    Status:r.status,CPF:fmtCpf(r.cpf),Nome:r.rawName,'Valor atual Excel':r.currentNet,
    'Líquido mensal PDF':r.pdfNet,'Líquido TRCT':r.trctNet,'Valor a gravar':r.target,
    'Regra aplicada':r.rule,'Célula destino':r.cell,'Linha Excel':r.row,'Página PDF':r.pdfPage,'Chave usada':r.matchKey||''
  }));
  const out=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(rows);
  ws['!cols']=[18,16,38,18,18,18,18,30,14,12,12,14].map(wch=>({wch}));
  XLSX.utils.book_append_sheet(out,ws,'Conferência');
  const comp=state.fill.competence?'_'+state.fill.competence.replace('/','-'):'';
  XLSX.writeFile(out,'Conferencia_Preenchimento_Liquido'+comp+'.xlsx');
}

function resetFill(){
  state.fill={excelFile:null,pdfFile:null,excel:null,pdf:null,rows:[],competence:''};
  $('fillExcelFile').value=''; $('fillPdfFile').value='';
  $('fillExcelStatus').textContent='Nenhum arquivo selecionado.';
  $('fillPdfStatus').textContent='Nenhum arquivo selecionado.';
  ['fillSummarySection','fillPreviewSection','fillDiagnosticsSection'].forEach(id=>$(id).classList.add('hidden'));
  fillHideMsg(); $('fillProgressWrap').classList.add('hidden'); fillReady();
}


function result(o,f,status,matchKey,diff){
  return {status,cpf:o?.cpf||f?.cpf||'',name:o?.rawName||f?.rawName||'',type:o?.type||f?.type||'',obs:o?.obs||'',
    originalNet:o?.net??null,folhaNet:f?.net??null,trctNet:f?.trctValue??null,diff,matchKey,originalRow:o?.row??'',originalCell:o?.netCell||'',pdfPage:f?.page??''};
}
function rank(s){ const a=[STATUS.DIV,STATUS.CENT,STATUS.VER,STATUS.FILL,STATUS.DUP,STATUS.SO_ORIG,STATUS.SO_FOLHA,STATUS.TRCT,STATUS.NOVO,STATUS.OK]; return a.indexOf(s); }

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
    if(diff===null){
      const status=f.net!==null&&o.net===null?STATUS.FILL:STATUS.VER;
      out.push(result(o,f,status,key,null));continue;
    }
    const ad=Math.abs(diff);
    out.push(result(o,f,ad<=tol?STATUS.OK:(ad<=cents?STATUS.CENT:STATUS.DIV),key,diff));
  }
  for(const f of folha) if(!used.has(f)) out.push(result(null,f,STATUS.SO_FOLHA,'',null));
  return out.sort((a,b)=>rank(a.status)-rank(b.status)||a.name.localeCompare(b.name,'pt-BR'));
}

function statusClass(s){
  if(s===STATUS.OK)return'ok'; if(s===STATUS.DIV)return'divergencia'; if(s===STATUS.CENT)return'centavos';
  if(s===STATUS.TRCT)return'trct'; if(s===STATUS.NOVO)return'novo'; if(s===STATUS.DUP)return'duplicado'; if(s===STATUS.FILL)return'liquidovazio'; if(s===STATUS.VER)return'verificar'; return'ausente';
}
function actionableStatuses(){ return new Set([STATUS.DIV,STATUS.CENT,STATUS.SO_ORIG,STATUS.SO_FOLHA,STATUS.DUP,STATUS.VER,STATUS.FILL]); }
function pendingReason(r){
  if(r.status===STATUS.DIV) return 'Diferença de líquido acima do limite de centavos.';
  if(r.status===STATUS.CENT) return 'Diferença de líquido dentro do limite de centavos configurado.';
  if(r.status===STATUS.SO_ORIG) return 'Registro existe no Excel original, mas não foi encontrado na folha.';
  if(r.status===STATUS.SO_FOLHA) return 'Registro existe na folha, mas não foi encontrado no Excel original.';
  if(r.status===STATUS.DUP) return 'Mais de um registro possível para a mesma chave.';
  if(r.status===STATUS.FILL) return 'Colaborador identificado no PDF por '+r.matchKey+'. O líquido foi lido, mas a célula '+(r.originalCell||('da linha '+r.originalRow))+' (“Remuneração líquida a receber”) está vazia no Excel. Use “Preencher líquidos com estes arquivos”.';
  if(r.status===STATUS.VER) return r.folhaNet===null ? 'Colaborador identificado, mas o valor líquido não foi reconhecido no PDF. Confira o extrato antes de preencher.' : 'Não foi possível comparar os valores líquidos. Confira o Excel e o PDF.';
  return '';
}
function comparisonSnapshot(competence=''){
  const r=state.results;
  const actionable=actionableStatuses();
  const pend=r.filter(x=>actionable.has(x.status));
  const comparable=r.filter(x=>Number.isFinite(x.originalNet)&&Number.isFinite(x.folhaNet)&&x.status!==STATUS.TRCT);
  const hasAnyOriginalValue=r.some(x=>Number.isFinite(x.originalNet));
  const validOriginal=r.filter(x=>x.status!==STATUS.SO_FOLHA&&Number.isFinite(x.originalNet));
  const total=(arr,key)=>arr.reduce((a,x)=>a+(Number.isFinite(x[key])?x[key]:0),0);
  const originalTotal=total(validOriginal,'originalNet');
  const folhaTotal=total(r.filter(x=>x.status!==STATUS.SO_ORIG),'folhaNet');
  const eff=comparable.reduce((a,x)=>a+(Number.isFinite(x.diff)?x.diff:0),0);
  let status='Folha aprovada',statusClassName='approved';
  if(!hasAnyOriginalValue||!comparable.length){ status='Validação parcial';statusClassName='partial'; }
  else if(pend.length){ status='Requer conferência';statusClassName='review'; }
  return {
    competence:competence||'Não identificada',
    status,statusClassName,
    collaborators:r.length,
    ok:r.filter(x=>x.status===STATUS.OK).length,
    divergences:comparable.length?r.filter(x=>x.status===STATUS.DIV||x.status===STATUS.CENT).length:null,
    pending:pend.length,
    trct:r.filter(x=>x.status===STATUS.TRCT).length,
    newEmployees:r.filter(x=>x.status===STATUS.NOVO).length,
    originalTotal:validOriginal.length?originalTotal:null,
    folhaTotal,
    effectiveDifference:comparable.length?eff:null,
    processedAt:new Date().toISOString()
  };
}
function getHistory(){
  try{
    const raw=localStorage.getItem(HISTORY_KEY);
    const data=raw?JSON.parse(raw):[];
    return Array.isArray(data)?data:[];
  }catch{return[];}
}
function setHistory(items){
  localStorage.setItem(HISTORY_KEY,JSON.stringify(items.slice(0,24)));
}
function saveHistory(competence){
  const item=comparisonSnapshot(competence);
  let items=getHistory();
  if(item.competence!=='Não identificada') items=items.filter(x=>x.competence!==item.competence);
  items.unshift(item);
  setHistory(items);
  renderHistory();
}
function renderHistory(){
  const items=getHistory();
  const body=$('historyBody'), empty=$('historyEmpty'), wrap=$('historyWrap');
  if(!body||!empty||!wrap)return;
  empty.classList.toggle('hidden',items.length>0);
  wrap.classList.toggle('hidden',items.length===0);
  body.innerHTML=items.map(x=>'<tr>'+
    '<td><strong>'+esc(x.competence)+'</strong></td>'+
    '<td><span class="history-status '+esc(x.statusClassName||'partial')+'">'+esc(x.status)+'</span></td>'+
    '<td>'+esc(x.collaborators)+'</td><td>'+esc(x.ok)+'</td><td>'+(x.originalTotal===null||x.divergences===null?'N/D':esc(x.divergences))+'</td><td>'+esc(x.pending)+'</td>'+
    '<td>'+esc(x.trct)+'</td><td>'+esc(x.newEmployees)+'</td>'+
    '<td class="money">'+(x.originalTotal===null?'N/D':moneyFmt.format(x.originalTotal))+'</td>'+
    '<td class="money">'+moneyFmt.format(Number(x.folhaTotal||0))+'</td>'+
    '<td class="money">'+(x.effectiveDifference===null?'N/D':moneyFmt.format(x.effectiveDifference))+'</td>'+
    '<td>'+esc(new Date(x.processedAt).toLocaleString('pt-BR'))+'</td></tr>').join('');
}
function renderApprovalStatus(){
  const missing=state.original.filter(r=>r.net===null).length;
  $('prepareFillBtn')?.classList.toggle('hidden',!missing||! /\.xlsx$/i.test(state.excelFile?.name||''));
  const actionable=actionableStatuses();
  const pend=state.results.filter(r=>actionable.has(r.status));
  const comparable=state.results.filter(r=>Number.isFinite(r.originalNet)&&Number.isFinite(r.folhaNet)&&r.status!==STATUS.TRCT);
  const hasAnyOriginalValue=state.results.some(r=>Number.isFinite(r.originalNet));
  const el=$('approvalStatus');
  el.classList.remove('hidden','approved','review','partial');
  if(!hasAnyOriginalValue || !comparable.length){
    el.classList.add('partial');
    el.innerHTML='<div class="icon">◐</div><div><strong>Validação parcial</strong><span>Não há líquido original suficiente para comparar valores. Use “Preencher líquidos com estes arquivos”, confira a prévia e gere uma cópia do Excel. Se os valores vierem deste mesmo PDF, a comparação posterior confirma a importação; não é uma conferência independente da folha.</span></div>';
  }else if(pend.length){
    el.classList.add('review');
    el.innerHTML='<div class="icon">⚠</div><div><strong>Requer conferência</strong><span>'+pend.length+' pendência(s) precisam ser revisada(s) antes do fechamento.</span></div>';
  }else{
    el.classList.add('approved');
    el.innerHTML='<div class="icon">✓</div><div><strong>Folha aprovada</strong><span>Nenhuma divergência ou pendência de conferência foi encontrada nos critérios atuais.</span></div>';
  }
}
function renderPending(){
  const actionable=actionableStatuses();
  const rows=state.results.filter(r=>actionable.has(r.status));
  $('pendingCount').textContent=rows.length ? rows.length+' pendência(s) encontrada(s)' : 'Sem pendências';
  $('pendingBody').innerHTML=rows.map(r=>'<tr class="row-'+statusClass(r.status)+'">'+
    '<td><span class="status '+statusClass(r.status)+'">'+esc(r.status)+'</span></td>'+
    '<td>'+esc(fmtCpf(r.cpf))+'</td><td>'+esc(r.name)+'</td>'+
    '<td class="money">'+(r.originalNet===null?'—':moneyFmt.format(r.originalNet))+'</td>'+
    '<td class="money">'+(r.folhaNet===null?'—':moneyFmt.format(r.folhaNet))+'</td>'+
    '<td class="money">'+(r.diff===null?'—':moneyFmt.format(r.diff))+'</td>'+
    '<td>'+esc(pendingReason(r))+'</td></tr>').join('');
  $('pendingEmpty').classList.toggle('hidden',rows.length!==0);
  $('pendingWrap').classList.toggle('hidden',rows.length===0);
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
    ['Ausentes',count(STATUS.SO_ORIG)+count(STATUS.SO_FOLHA),'warn'],['Dados a revisar / duplicados',count(STATUS.VER)+count(STATUS.DUP),'warn'],
    ['Líquidos vazios no Excel',state.original.filter(x=>x.net===null).length,'warn'],
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
  '<div><strong>Líquidos vazios no Excel:</strong> '+state.original.filter(x=>x.net===null).length+'. O status “Líquido vazio no Excel” indica correspondência encontrada e preenchimento pendente, sem comparação monetária.</div>'+
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
  $('fieldAuditSection')?.classList.add('hidden');
  hideMsg(); $('summarySection').classList.add('hidden'); $('pendingSection').classList.add('hidden'); $('resultsSection').classList.add('hidden'); $('diagnosticsSection').classList.add('hidden');
  $('compareBtn').disabled=true;
  try{
    progress(5,'Lendo Excel...');
    const ex=await parseExcel(state.excelFile); state.original=ex.records;
    progress(25,'Excel lido. Iniciando PDF...');
    const pf=await parsePdf(state.pdfFile); state.folha=pf.records;
    progress(75,'Comparando CPF, nomes e valores...');
    const tol=Number($('tolerance').value||.01), cents=Math.max(tol,Number($('centsLimit').value||1));
    state.results=compare(state.original,state.folha,tol,cents);
    state.auditRows=auditFields(ex.records,pf.records,tol); state.auditPdf=pf.records;state.auditLimit=200;renderFieldAudit();
    state.diagnostics={sheetName:ex.sheetName,headerRow:ex.headerRow,originalCount:ex.records.length,pdfPages:pf.pages,folhaCount:pf.records.length};
    renderSummary(); renderApprovalStatus(); renderPending(); renderResults(); renderDiag();
    saveHistory(pf.competence);
    $('summarySection').classList.remove('hidden'); $('pendingSection').classList.remove('hidden'); $('resultsSection').classList.remove('hidden'); $('diagnosticsSection').classList.remove('hidden');
    progress(100,'Concluído.'); setTimeout(()=>$('progressWrap').classList.add('hidden'),500);
    const missing=state.original.filter(r=>r.net===null).length;
    msg(missing ? 'Arquivos reconhecidos: '+state.original.length+' colaboradores no Excel e '+state.folha.length+' no PDF. Conferência C–AN concluída. Há '+missing+' líquido(s) vazio(s); salário, cadastro e rubricas disponíveis foram conferidos separadamente.' : 'Comparação concluída: '+state.original.length+' registros no Excel e '+state.folha.length+' colaboradores identificados no PDF.',missing?'warning':'success');
  }catch(e){ console.error(e); $('progressWrap').classList.add('hidden'); msg(e?.message||'Erro ao processar os arquivos.','error'); }
  finally{ ready(); }
}
function reset(){
  Object.assign(state,{excelFile:null,pdfFile:null,original:[],folha:[],results:[],diagnostics:{},auditRows:[],auditPdf:[]});
  $('fieldAuditSection')?.classList.add('hidden');
  $('excelFile').value='';$('pdfFile').value='';$('excelStatus').textContent='Nenhum arquivo selecionado.';$('pdfStatus').textContent='Nenhum arquivo selecionado.';
  $('summarySection').classList.add('hidden');$('pendingSection').classList.add('hidden');$('resultsSection').classList.add('hidden');$('diagnosticsSection').classList.add('hidden');hideMsg();$('progressWrap').classList.add('hidden');ready();
}
$('compareBtn').addEventListener('click',run);
$('resetBtn').addEventListener('click',reset);
$('exportBtn').addEventListener('click',exportXlsx);
$('searchInput').addEventListener('input',renderResults);
$('statusFilter').addEventListener('change',renderResults);
$('clearHistoryBtn').addEventListener('click',()=>{
  if(confirm('Deseja apagar o histórico resumido salvo neste navegador?')){
    localStorage.removeItem(HISTORY_KEY);
    renderHistory();
    msg('Histórico local apagado. Os arquivos de folha nunca foram armazenados.','success');
  }
});
renderHistory();

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

$('prepareFillBtn')?.addEventListener('click',async()=>{
  if(!state.excelFile||!state.pdfFile)return;
  state.fill.excelFile=state.excelFile; state.fill.pdfFile=state.pdfFile;
  $('fillExcelStatus').textContent=state.excelFile.name;
  $('fillPdfStatus').textContent=state.pdfFile.name;
  fillReady(); switchModule('fill');
  await runFill();
});

$('fillAnalyzeBtn').addEventListener('click',runFill);
$('fillResetBtn').addEventListener('click',resetFill);
$('fillGenerateBtn').addEventListener('click',generateFilledExcel);
$('fillAuditBtn').addEventListener('click',exportFillAudit);
$('fillSearchInput').addEventListener('input',renderFillPreview);
$('fillStatusFilter').addEventListener('change',renderFillPreview);

const AUDIT_ACTIONS=new Set(['DIVERGÊNCIA','RUBRICA AUSENTE NO PDF','SEM VALOR NO EXCEL','VALOR NÃO RECONHECIDO NO PDF','CORRESPONDÊNCIA AMBÍGUA','COLABORADOR AUSENTE NO PDF','COLABORADOR AUSENTE NO EXCEL','COLUNA NÃO IDENTIFICADA','CABEÇALHO DIFERENTE','EXTRAÇÃO INCOMPLETA']);
function filteredAudit(){const q=norm($('auditSearch').value),sf=$('auditStatus').value,col=$('auditField').value;return state.auditRows.filter(r=>(sf==='ALL'||(sf==='ACTION'?AUDIT_ACTIONS.has(r.status):r.status===sf))&&(col==='ALL'||r.col===col)&&(!q||norm(r.name+' '+r.cpf+' '+r.field+' '+r.reason).includes(q)));}
function renderFieldAudit(){
 $('fieldAuditSection').classList.remove('hidden');const sum=auditSummary(state.auditRows);
 $('auditMetrics').innerHTML=[['Campos analisados',sum.fields],['Comparações compatíveis',sum.compatible],['Divergências',sum.divergences],['Sem valor no Excel',sum.missingExcel],['Rubricas ausentes no PDF',sum.missingRubric],['Conferência manual',sum.manual],['Sem equivalente no PDF',sum.unavailable],['Não reconhecidos no PDF',sum.unreadable]].map(([l,v])=>'<div class="metric '+(l==='Divergências'&&v?'bad':'info')+'"><span class="label">'+l+'</span><span class="value">'+v+'</span></div>').join('');
 $('auditField').innerHTML='<option value="ALL">Todas as colunas C a AN</option>'+FIELD_SCHEMA.map(f=>'<option value="'+f.col+'">'+f.col+' — '+esc(f.label)+'</option>').join('');
 $('auditMapping').innerHTML=FIELD_SCHEMA.map(f=>'<tr><td>'+f.col+'</td><td>'+esc(f.label)+'</td><td>'+esc(f.kind==='unavailable'?'Sem equivalente neste extrato':f.codes.length?'Rubricas '+f.codes.join(', '):['name','category','admission','stipend','salary','cpf','net'].includes(f.kind)?'Campo do cadastro / salário / líquido':f.pattern?'Descrição da rubrica (tipo e unidade conferidos)':'Conferência manual')+'</td></tr>').join('');renderAuditTable();
}
function renderAuditTable(){const rows=filteredAudit(),shown=rows.slice(0,state.auditLimit);$('auditCount').textContent=shown.length+' de '+rows.length+' campo(s) neste filtro; '+state.auditRows.length+' campos no total.';
 $('auditBody').innerHTML=shown.map(r=>'<tr><td><span class="status '+(r.status==='DIVERGÊNCIA'?'divergencia':r.status==='COMPATÍVEL'?'ok':'verificar')+'">'+esc(r.status)+'</span></td><td>'+esc(r.name)+'</td><td>'+esc(fmtCpf(r.cpf))+'</td><td>'+esc(r.cell)+'</td><td>'+esc(r.field)+'</td><td>'+esc(r.excel)+'</td><td>'+esc(r.pdf)+'</td><td>'+esc(r.difference===null?'—':r.unit==='R$'?moneyFmt.format(r.difference):r.difference.toFixed(4)+' '+r.unit)+'</td><td>'+esc(r.reason)+'</td><td>'+esc(r.page)+'</td></tr>').join('');
 $('auditMore').classList.toggle('hidden',rows.length<=state.auditLimit);
}
function exportFieldAudit(){if(!state.auditRows.length)return;const wb=XLSX.utils.book_new();const data=state.auditRows.map(r=>({Status:r.status,Nome:r.name,CPF:fmtCpf(r.cpf),Coluna:r.col,Célula:r.cell,Campo:r.field,Excel:r.excel,PDF:r.pdf,Diferença:r.difference,Unidade:r.unit,Motivo:r.reason,'Página PDF':r.page,Chave:r.matchKey}));
 const append=(name,rows)=>{const ws=XLSX.utils.json_to_sheet(rows);XLSX.utils.book_append_sheet(wb,ws,name);};
 append('Conferência C-AN',data);append('Divergências',data.filter(r=>r.Status==='DIVERGÊNCIA'));append('Pendências',data.filter(r=>AUDIT_ACTIONS.has(r.Status)&&r.Status!=='DIVERGÊNCIA'));append('Mapeamento',FIELD_SCHEMA.map(f=>({Coluna:f.col,Campo:f.label,Tipo:f.kind,Rubricas:f.codes.join(', '),Descrição:f.pattern,Limitação:f.kind==='unavailable'?'Sem equivalente no PDF':'Conferir unidades e vínculo; ausência não equivale a aprovação'})));
 append('Rubricas PDF',state.auditPdf.flatMap(p=>(p.fields?.rubrics||[]).map(r=>({Nome:p.rawName,CPF:fmtCpf(p.cpf),Página:p.page,Código:r.code,Descrição:r.label,Referência:r.reference,Valor:r.value,Tipo:r.type}))));
 XLSX.writeFile(wb,'Conferencia_C_AN_Folha_'+new Date().toISOString().slice(0,10)+'.xlsx');
}
$('auditSearch')?.addEventListener('input',()=>{state.auditLimit=200;renderAuditTable();});
['auditStatus','auditField'].forEach(id=>$(id)?.addEventListener('change',()=>{state.auditLimit=200;renderAuditTable();}));
$('auditMore')?.addEventListener('click',()=>{state.auditLimit+=200;renderAuditTable();});
$('auditExport')?.addEventListener('click',exportFieldAudit);
