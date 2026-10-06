// Fixtures fictícias; nenhum arquivo ou dado de colaborador é incluído no repositório.
const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),path=require('node:path');
const stub=new Proxy(()=>{}, {get:(_,p)=>p==='getItem'?()=>null:p==='value'?'':stub,apply:()=>stub});
(async()=>{
const audit=await import('../assets/js/field-audit.js');
for(const name of ['app.js','app.local.js']){
 const context={console,window:{},document:{getElementById:()=>stub},localStorage:stub,pdfjsLib:{GlobalWorkerOptions:{}},...audit,setTimeout:()=>{}};
 vm.createContext(context);
 vm.runInContext(fs.readFileSync(path.join(__dirname,'../assets/js',name),'utf8').split('\n').filter(l=>!l.startsWith('import ')).join('\n'),context);
 const lines=['Emissão: 06/10/2026','Competência: 09/2026','Empr.: 1 PESSOA TESTE Situação: Trabalhando CPF: 123.456.789-00 Adm: 01/08/2026','Vínculo: Celetista CC: 1','Proventos: 1.000,00 Descontos: 100,00 Líquido: 900,00'];
 context.fixture=lines;
 vm.runInContext('pdfLines=async()=>({numPages:1,pages:[{page:1,lines:fixture}]})',context);
 (async()=>{
  const pdf=await context.parsePdf({},()=>{});
  assert.equal(pdf.competence,'09/2026');assert.equal(pdf.records.length,1);assert.equal(pdf.records[0].net,900);
  const blank={rawName:'PESSOA TESTE',nameNorm:'PESSOA TESTE',cpf:'12345678900',net:null,currentNet:null,isTRCT:false,isNew:false,row:10,cell:'AG10',netCell:'AG10'};
  const comparison=context.compare([blank],pdf.records,.01,1);
  assert.equal(comparison[0].status,'LÍQUIDO VAZIO NO EXCEL');assert.equal(comparison[0].diff,null);
  assert.match(context.pendingReason(comparison[0]),/AG10.*está vazia/);
  const unreadable=context.compare([blank],[{...pdf.records[0],net:null}],.01,1);
  assert.equal(unreadable[0].status,'VERIFICAR');assert.match(context.pendingReason(unreadable[0]),/não foi reconhecido no PDF/);
  const filled=context.compare([{...blank,net:900}],pdf.records,.01,1);assert.equal(filled[0].status,'OK');
  const divergent=context.compare([{...blank,net:890}],pdf.records,.01,1);assert.equal(divergent[0].status,'DIVERGÊNCIA');
  const fill=context.prepareFillRows([blank],pdf.records);
  assert.equal(fill[0].status,'PREENCHER');assert.equal(fill[0].target,900);assert.equal(fill[0].cell,'AG10');
  context.fixture=lines.filter(l=>!l.startsWith('Competência:'));
  assert.equal((await context.parsePdf({},()=>{})).competence,'');
  context.fixture=['Competência: 09/2026'];
  await assert.rejects(()=>context.parsePdf({},()=>{}),/Nenhum colaborador/);
  console.log(name+': competência, líquidos vazios, preenchimento e PDF não reconhecido — OK');
 })().catch(e=>{console.error(e);process.exitCode=1;});
}

})().catch(e=>{console.error(e);process.exitCode=1;});
