(()=>{
  const oldBtn=document.getElementById('exportRecon');
  if(!oldBtn)return;

  const wrap=document.createElement('div');
  wrap.className='vac-export-controls';
  wrap.innerHTML=`
    <select id="exportScope" aria-label="Escopo da exportação">
      <option value="all">Todos os registros</option>
      <option value="divergences">Somente divergências</option>
      <option value="urgent">Somente urgentes</option>
      <option value="filtered">Resultado filtrado atual</option>
    </select>
    <select id="exportFormat" aria-label="Formato da exportação">
      <option value="xlsx">Excel (.xlsx)</option>
      <option value="csv">CSV (.csv)</option>
      <option value="pdf">PDF (.pdf)</option>
      <option value="print">Imprimir</option>
    </select>
    <button id="exportRun" class="secondary" type="button">Exportar</button>`;
  oldBtn.replaceWith(wrap);

  const style=document.createElement('style');
  style.textContent='.vac-export-controls{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.vac-export-controls select{border:1px solid #dce6f2;border-radius:8px;padding:9px 10px;background:#fff;min-width:170px}@media(max-width:700px){.vac-export-controls{width:100%}.vac-export-controls select,.vac-export-controls button{width:100%}}';
  document.head.appendChild(style);

  const normText=v=>String(v??'').trim();
  function currentRecon(){
    const q=typeof norm==='function'?norm(document.getElementById('reconSearch')?.value||''):'';
    const s=document.getElementById('reconStatus')?.value||'ALL';
    const src=Array.isArray(state?.recon)?state.recon:[];
    return src.filter(r=>(!q||norm(r.name).includes(q))&&(s==='ALL'||r.status===s));
  }
  function urgentBalances(){
    const src=state?.excel?.saldos||[];
    return src.filter(r=>/Prazo ultrapassado|Urgente/i.test(String(r['Alerta']||''))).map(r=>({
      status:r['Alerta']||'Urgente',name:r['Colaborador']||'',period:r['Competência']||'',excel:r['Saldo após reservas']??r['Saldo efetivo']??'',pdf:'',right:'',limit:typeof excelDate==='function'?excelDate(r['Limite oficial']||r['Fim concessivo']):r['Limite oficial']||'',detail:r['Conferência']||''
    }));
  }
  function dataForScope(scope){
    const all=Array.isArray(state?.recon)?state.recon:[];
    if(scope==='divergences') return all.filter(r=>['Divergente','Verificar','Conferir histórico'].includes(r.status));
    if(scope==='urgent') return urgentBalances();
    if(scope==='filtered') return currentRecon();
    return all;
  }
  const headers=['Status','Colaborador','Competência / período','Saldo Excel','Saldo PDF','Direito PDF','Limite','Detalhe'];
  const rowsFrom=data=>data.map(r=>[r.status||'',r.name||'',r.period||'',r.excel??'',r.pdf??'',r.right??'',r.limit??'',r.detail??'']);
  function saveBlob(blob,name){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1200)}
  function exportCsv(data){
    const rows=[headers,...rowsFrom(data)];
    const csv='\ufeff'+rows.map(r=>r.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(';')).join('\r\n');
    saveBlob(new Blob([csv],{type:'text/csv;charset=utf-8'}),'ferias-exportacao.csv');
  }
  function exportXlsx(data,scope){
    if(!window.XLSX)return alert('Biblioteca de Excel indisponível.');
    const wb=XLSX.utils.book_new();
    const ws=XLSX.utils.aoa_to_sheet([headers,...rowsFrom(data)]);
    ws['!cols']=[{wch:18},{wch:34},{wch:28},{wch:14},{wch:14},{wch:14},{wch:14},{wch:50}];
    XLSX.utils.book_append_sheet(wb,ws,'Dados');
    const summary=[['Relatório','Férias - Portal Nomus'],['Escopo',scope],['Registros',data.length],['Gerado em',new Date().toLocaleString('pt-BR')],['Fonte PDF',state?.pdf?.fileName||'Não carregado'],['Data-base PDF',state?.pdf?.baseDate||'']];
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(summary),'Resumo');
    XLSX.writeFile(wb,'ferias-exportacao.xlsx');
  }
  function htmlReport(data,title){
    const esc2=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    return `<!doctype html><html><head><meta charset="utf-8"><title>${esc2(title)}</title><style>body{font-family:Arial,sans-serif;color:#1e293b;padding:24px}h1{font-size:22px;margin:0 0 6px}p{color:#64748b;margin:0 0 18px}table{width:100%;border-collapse:collapse;font-size:10px}th,td{border:1px solid #cbd5e1;padding:6px;text-align:left;vertical-align:top}th{background:#eef5ff}@page{size:A4 landscape;margin:12mm}</style></head><body><h1>${esc2(title)}</h1><p>${data.length} registro(s) • ${new Date().toLocaleString('pt-BR')}</p><table><thead><tr>${headers.map(h=>`<th>${esc2(h)}</th>`).join('')}</tr></thead><tbody>${rowsFrom(data).map(r=>`<tr>${r.map(v=>`<td>${esc2(v)}</td>`).join('')}</tr>`).join('')}</tbody></table></body></html>`;
  }
  function printReport(data,title){
    const w=window.open('','_blank');if(!w)return alert('Permita pop-ups para imprimir o relatório.');w.document.write(htmlReport(data,title));w.document.close();w.focus();setTimeout(()=>w.print(),250);
  }
  async function exportPdf(data,title){
    if(!window.jspdf?.jsPDF){return printReport(data,title+' — use “Salvar como PDF” na impressão');}
    const {jsPDF}=window.jspdf;const doc=new jsPDF({orientation:'landscape',unit:'pt',format:'a4'});doc.setFontSize(14);doc.text(title,36,32);doc.setFontSize(9);doc.text(`${data.length} registro(s) • ${new Date().toLocaleString('pt-BR')}`,36,48);if(typeof doc.autoTable==='function'){doc.autoTable({head:[headers],body:rowsFrom(data),startY:60,styles:{fontSize:7,cellPadding:3},headStyles:{fillColor:[31,95,191]}});doc.save('ferias-exportacao.pdf');}else{printReport(data,title+' — use “Salvar como PDF” na impressão');}
  }
  async function ensurePdfLib(){
    if(window.jspdf?.jsPDF&&window.jspdf.jsPDF.API?.autoTable)return;
    const load=src=>new Promise((resolve,reject)=>{const s=document.createElement('script');s.src=src;s.onload=resolve;s.onerror=reject;document.head.appendChild(s)});
    try{if(!window.jspdf?.jsPDF)await load('assets/vendor/jspdf/jspdf.umd.min.js');await load('assets/vendor/jspdf/jspdf.plugin.autotable.min.js');}catch(e){console.warn('Falha ao carregar gerador PDF',e)}
  }

  document.getElementById('exportRun').onclick=async()=>{
    if(typeof buildRecon==='function')buildRecon();
    const scope=document.getElementById('exportScope').value;
    const format=document.getElementById('exportFormat').value;
    const data=dataForScope(scope);
    if(!data.length)return alert('Não há registros para o escopo selecionado.');
    const label={all:'Todos os registros',divergences:'Somente divergências',urgent:'Somente urgentes',filtered:'Resultado filtrado atual'}[scope]||scope;
    const title=`Férias — ${label}`;
    if(format==='csv')return exportCsv(data);
    if(format==='xlsx')return exportXlsx(data,label);
    if(format==='print')return printReport(data,title);
    if(format==='pdf'){await ensurePdfLib();return exportPdf(data,title);}
  };
})();