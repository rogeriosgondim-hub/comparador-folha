// Conferência local por campo. Ausência de evidência nunca equivale a aprovação.
export const FIELD_SCHEMA = [
 ['C','Nome','name'],['D','Função','category'],['E','Admissão','admission'],['F','Nascimento','unavailable'],['G','Idade','unavailable'],['H','Carga horária','hoursMonthly'],
 ['I','Bolsa de estágio','stipend'],['J','Salário Base','salary'],['K','Custo do Plano do Funcionário (Sem IOF)','unavailable'],['L','Custo do Plano do Funcionário (Com IOF)','unavailable'],['M','Custo do Plano dos dependentes','unavailable'],['N','Custo do Plano dos dependentes (Com IOF)','unavailable'],['O','Valor pago pela Nomus do plano Básico','unavailable'],
 ['P','Descontos Plano de Saúde','money','339'],['Q','Descontos Odonto Dependentes','money','334'],['R','Desconto Wellhub (Gympass)','money','', 'WELLHUB|GYMPASS'],['S','Recebe VR ?','indicator','236,248'],['T','Recebe VA ?','indicator','9383'],['U','Descontos de Vale Transporte','transport','', 'VALE TRANSPORTE|DESC VT'],['V','Desconto/ Ref. Faltas e Atrasos','absence','8792,8794'],
 ['W','Horas extras 50%','hours','150','HORAS EXTRAS 50'],['X','Horas extras 100%','hours','','HORAS EXTRAS 100'],['Y','Adicional Noturno (22:00 às 06:00)','hours','25','ADICIONAL NOTURNO'],['Z','Comissões','money','','COMISSAO|COMISSOES'],['AA','Auxílio Home office','money','416'],['AB','Auxílio Educação','money','447'],['AC','Auxílio Creche','money','492'],['AD','Empréstimo Desconto','loan','474'],['AE','Bônus','money','','BONUS|BONIFICACAO'],['AF','Observações','notes'],['AG','Remuneração líquida a receber','net'],['AH','Banco á Pagar / PIX','unavailable'],['AI','Chave','unavailable'],['AJ','Agencia','unavailable'],['AK','Conta Corrente','unavailable'],['AL','CPF','cpf'],['AM','Plano de Saúde','indicator','339'],['AN','Plano escolhido','unavailable']
].map(([col,label,kind,codes='',pattern=''])=>({col,label,kind,codes:codes?codes.split(','):[],pattern}));
export function normalize(v=''){return String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Za-z0-9 ]/g,' ').replace(/\s+/g,' ').trim().toUpperCase();}
const digits=v=>String(v??'').replace(/\D/g,'');
const blank=v=>v===null||v===undefined||v==='';
function amount(v){if(blank(v))return null;if(typeof v==='number')return Number.isFinite(v)?v:null;let t=String(v).trim().replace(/R\$|\s/g,'');if(!/^-?[\d.,]+$/.test(t))return null;if(t.includes(','))t=t.replace(/\./g,'').replace(',','.');const n=Number(t);return Number.isFinite(n)?n:null;}
function dateValue(v){if(v instanceof Date||Object.prototype.toString.call(v)==='[object Date]'){if(!Number.isFinite(v.getTime()))return '';return [v.getFullYear(),String(v.getMonth()+1).padStart(2,'0'),String(v.getDate()).padStart(2,'0')].join('-');}const m=String(v??'').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);return m?`${m[3]}-${m[2]}-${m[1]}`:'';}
function category(v){const t=normalize(v);if(/ESTAG/.test(t))return 'ESTAGIARIO';if(/APRENDIZ/.test(t))return 'APRENDIZ';if(/FUNCIONARIO|CELETISTA|CLT/.test(t))return 'CELETISTA';return t;}
function hours(v,display=''){const t=String(display||v||'').trim();const m=t.match(/^(\d+):(\d{2})(?::(\d{2}))?$/);if(m)return +m[1]+(+m[2])/60+(+m[3]||0)/3600;if(Object.prototype.toString.call(v)==='[object Date]')return v.getUTCHours()+v.getUTCMinutes()/60+v.getUTCSeconds()/3600;return amount(v);}
export function parseFieldPdf(block){
 const text=block.join(' '),rubrics=[];
 // As duas colunas da folha se tornam duas ocorrências na mesma linha.
 const re=/(?:^|\s)(\d{1,5})\s+(.+?)\s+(-?(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2}|\d+:\d{2})\s+(-?(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2})\s+([PDI])(?=\s|$)/g;
 for(const line of block){for(const m of line.matchAll(re)){rubrics.push({code:m[1],label:m[2].trim(),reference:m[3],value:amount(m[4]),type:m[5]});}}
 const proventos=amount((text.match(/Proventos:\s*(-?(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2})/i)||[])[1]),descontos=amount((text.match(/Descontos:\s*(-?(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2})/i)||[])[1]);
 const totalP=rubrics.filter(r=>r.type==='P').reduce((a,r)=>a+r.value,0),totalD=rubrics.filter(r=>r.type==='D').reduce((a,r)=>a+r.value,0);
 const totalsMatched=proventos!==null&&descontos!==null&&Math.abs(totalP-proventos)<=.02&&Math.abs(totalD-descontos)<=.02;
 return {proventos,descontos,totalP,totalD,totalsMatched,salary:amount((text.match(/Sal[aá]rio:\s*(-?(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2})/i)||[])[1]),admission:(text.match(/Adm:\s*(\d{2}\/\d{2}\/\d{4})/i)||[])[1]||'',monthlyHours:amount((text.match(/Horas M[eê]s:\s*(\d+,\d{2})/i)||[])[1]),role:(text.match(/Cargo:\s*\d+\s+(.+?)\s+C\.B\.O:/i)||[])[1]||'',rubrics};
}
export function matchedRubrics(field,pdf){const rs=pdf.fields?.rubrics||[];return rs.filter(r=>{
 if(['money','transport','loan','absence','indicator'].includes(field.kind)&&r.type!=='D'&& !['Z','AA','AB','AC','AE'].includes(field.col))return false;
 if(['Z','AA','AB','AC','AE','W','X','Y'].includes(field.col)&&r.type!=='P')return false;
 return field.codes.includes(r.code)||(field.pattern&&new RegExp(field.pattern).test(normalize(r.label)));
 });}
function readable(v,display){if(blank(v))return 'Vazio';if(Object.prototype.toString.call(v)==='[object Date]')return dateValue(v);return String(display||v);}
export function auditFields(excelRecords,pdfRecords,tolerance=.01){
 const out=[],cp=new Map(),nm=new Map(),ec=new Map(),en=new Map();
 const add=(m,k,v)=>{if(k)m.set(k,[...(m.get(k)||[]),v]);};
 for(const p of pdfRecords){add(cp,p.cpf,p);add(nm,p.nameNorm,p);}for(const e of excelRecords){add(ec,e.cpf,e);add(en,e.nameNorm,e);}
 for(const e of excelRecords){const candidates=e.cpf&&cp.has(e.cpf)?cp.get(e.cpf):nm.get(e.nameNorm)||[];const key=e.cpf&&cp.has(e.cpf)?'CPF':'NOME';const duplicate=(e.cpf&&(ec.get(e.cpf)||[]).length>1)||(!e.cpf&&(en.get(e.nameNorm)||[]).length>1)||candidates.length>1;const p=!duplicate&&candidates.length===1?candidates[0]:null;
 for(const f of FIELD_SCHEMA){const cell=e.auditCells?.[f.col];const row={name:e.rawName,cpf:e.cpf,row:e.row,col:f.col,cell:f.col+e.row,field:f.label,excel:readable(cell?.value,cell?.display),pdf:'—',difference:null,unit:'',status:'',reason:'',page:p?.page||'',matchKey:p?key:''};const finish=(status,reason)=>{row.status=status;row.reason=reason;out.push(row);};
 if(!cell){finish('COLUNA NÃO IDENTIFICADA','O cabeçalho desta coluna não foi encontrado no intervalo C a AN.');continue;}
 if(normalize(cell.header)!==normalize(f.label)){finish('CABEÇALHO DIFERENTE','A coluna contém “'+cell.header+'”; o mapeamento esperado é “'+f.label+'”.');continue;}
 if(duplicate){finish('CORRESPONDÊNCIA AMBÍGUA','Há registros duplicados. Nenhum valor foi associado automaticamente.');continue;}
 if(!p){finish('COLABORADOR AUSENTE NO PDF','Não foi encontrada correspondência segura para este colaborador.');continue;}
 if(f.kind==='unavailable'){finish('SEM EQUIVALENTE NO PDF','O extrato não informa este dado; não há validação entre os arquivos.');continue;}
 const v=cell.value, rs=matchedRubrics(f,p), evidence=rs.map(r=>`${r.code} ${r.label}`).join('; '), total=rs.reduce((a,r)=>a+r.value,0);
 const numeric=(a,b,unit='R$')=>{row.unit=unit;row.pdf=b===null?'Não reconhecido':unit==='R$'?new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(b):b.toFixed(4)+' horas';if(a!==null)row.excel=unit==='R$'?new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(a):a.toFixed(4)+' horas';if(a===null){finish('SEM VALOR NO EXCEL','O PDF informa o valor, mas '+row.cell+' está vazio ou não contém um valor válido.');return;}if(b===null){finish('VALOR NÃO RECONHECIDO NO PDF','Confira o campo correspondente no extrato.');return;}row.difference=b-a;const limit=unit==='horas'?1/3600:tolerance;finish(Math.abs(row.difference)<=limit+1e-8?'COMPATÍVEL':'DIVERGÊNCIA',unit==='R$'?'PDF menos Excel; tolerância monetária de R$ '+tolerance.toFixed(2)+'. '+evidence:'Referência de horas da rubrica; diferença em horas. '+evidence);};
 if(f.kind==='name'){row.pdf=p.rawName;finish(e.nameNorm===p.nameNorm?'COMPATÍVEL':'DIVERGÊNCIA','Nome normalizado; marcações de demissão e acentos são desconsiderados.');continue;}
 if(f.kind==='cpf'){row.pdf=p.cpf;finish(digits(v)===p.cpf?'COMPATÍVEL':'DIVERGÊNCIA','CPF comparado somente pelos dígitos.');continue;}
 if(f.kind==='category'){row.pdf=p.type+' / '+(p.fields?.role||'');const b=category(/APRENDIZ/.test(normalize(p.fields?.role))?'APRENDIZ':p.type);finish(category(v)===b?'COMPATÍVEL':'DIVERGÊNCIA','A coluna Função contém a categoria do vínculo; não o título detalhado do cargo.');continue;}
 if(f.kind==='admission'){row.pdf=p.fields?.admission||'Não reconhecido';finish(!dateValue(v)?'SEM VALOR NO EXCEL':!dateValue(row.pdf)?'VALOR NÃO RECONHECIDO NO PDF':dateValue(v)===dateValue(row.pdf)?'COMPATÍVEL':'DIVERGÊNCIA','Data de admissão comparada em dia, mês e ano.');continue;}
 if(f.kind==='hoursMonthly'){row.pdf=p.fields?.monthlyHours==null?'Não reconhecido':p.fields.monthlyHours+' horas/mês';finish('CONFERÊNCIA MANUAL','A planilha descreve jornada semanal/período e o PDF informa horas mensais. Não se aplica um fator de conversão presumido.');continue;}
 if(['salary','stipend'].includes(f.kind)){const intern=category(p.type)==='ESTAGIARIO';const apprenticeInI=category(p.type)==='APRENDIZ'&&amount(e.auditCells?.I?.value)>0&&!(amount(e.auditCells?.J?.value)>0);const usesI=intern||apprenticeInI;const applies=f.kind==='stipend'?usesI:!usesI;if(!applies){row.pdf='Campo de salário destinado a '+(usesI?'Remuneração na coluna I':'Salário Base (J)');finish(amount(v)>0?'CONFERÊNCIA MANUAL':'NÃO APLICÁVEL','Não somar bolsa e salário-base; verificar o vínculo do mês.');}else {numeric(amount(v),p.fields?.salary??null);if(apprenticeInI)row.reason+=' Remuneração do aprendiz cadastrada na coluna I, com J vazio.';}continue;}
 if(f.kind==='net'){numeric(amount(v),p.isTRCT?(p.trctValue??p.net):p.net);row.reason+=(p.isTRCT?' Rescisão: priorizado líquido TRCT.':' Líquido mensal.');continue;}
 if(f.kind==='notes'){row.pdf=p.obs;finish('CONFERÊNCIA MANUAL','Observação livre na planilha; situação e eventos da folha podem exigir interpretação da competência.');continue;}
 if(f.kind==='indicator'){row.pdf=rs.length?evidence+': '+total.toFixed(2):'Nenhuma rubrica relacionada localizada';finish('CONFERÊNCIA MANUAL','O desconto é apenas um indício; não comprova o direito ao benefício nem o valor integral recebido.');continue;}
 if(f.kind==='absence'){row.pdf=rs.length?rs.map(r=>`${r.code} ${r.label}: ref. ${r.reference}, R$ ${r.value.toFixed(2)}`).join('; '):'Nenhuma rubrica de faltas localizada';finish('CONFERÊNCIA MANUAL','A planilha pode conter texto, datas ou horas, enquanto o PDF pode descontar dias e DSR. Confira as unidades e o período.');continue;}
 if(f.kind==='loan'){const consign=(p.fields?.rubrics||[]).filter(r=>r.type==='D'&&/CRED[ .]*TRAB/.test(normalize(r.label)));if(consign.length){row.pdf=consign.map(r=>`${r.code}: R$ ${r.value.toFixed(2)}`).join('; ')+(rs.length?'; empréstimo 474: '+total.toFixed(2):'');finish('CONFERÊNCIA MANUAL','Há crédito consignado/provisões. Não somar automaticamente contratos e provisões ao empréstimo interno da planilha.');continue;}}
 if(f.kind==='transport'&&normalize(v)==='NAO'){row.pdf=rs.length?String(total):'Nenhuma rubrica localizada';row.difference=rs.length?total:null;finish(rs.length&&total>0?'DIVERGÊNCIA':'SEM MOVIMENTO','A planilha indica Não; conferido o lançamento de vale-transporte no extrato.');continue;}
 if(p.fields?.totalsMatched===false){row.pdf=rs.length?String(total):'Não reconhecido com segurança';finish('EXTRAÇÃO INCOMPLETA','A soma das rubricas extraídas não confere com proventos/descontos do bloco. Confira a página antes de concluir a comparação.');continue;}
 if(f.kind==='transport'&&!blank(v)&&amount(v)===null){row.pdf=rs.length?evidence+': '+total.toFixed(2):'Nenhuma rubrica relacionada localizada';finish('CONFERÊNCIA MANUAL','A planilha contém indicação textual ou percentual de vale-transporte; o PDF contém desconto monetário.');continue;}
 if(!rs.length){row.pdf='Rubrica não localizada';finish(blank(v)||amount(v)===0?'SEM MOVIMENTO':'RUBRICA AUSENTE NO PDF','A ausência da rubrica não foi transformada em valor zero nem em aprovação.');continue;}
 if(f.kind==='hours'){numeric(hours(v,cell.display),rs.reduce((a,r)=>a+(hours(r.reference)||0),0),'horas');continue;}
 numeric(amount(v),total);
 }
 }
 for(const p of pdfRecords){if(!excelRecords.some(e=>(e.cpf&&e.cpf===p.cpf)||e.nameNorm===p.nameNorm))out.push({name:p.rawName,cpf:p.cpf,row:'',col:'—',cell:'—',field:'Colaborador',excel:'Ausente',pdf:'Presente',difference:null,unit:'',status:'COLABORADOR AUSENTE NO EXCEL',reason:'Não foi encontrado no Excel.',page:p.page,matchKey:''});}
 return out;
}
export function auditSummary(rows){const count=s=>rows.filter(r=>r.status===s).length;return {fields:rows.length,compatible:count('COMPATÍVEL'),divergences:count('DIVERGÊNCIA'),missingExcel:count('SEM VALOR NO EXCEL'),unavailable:count('SEM EQUIVALENTE NO PDF'),manual:count('CONFERÊNCIA MANUAL'),missingRubric:count('RUBRICA AUSENTE NO PDF'),unreadable:count('VALOR NÃO RECONHECIDO NO PDF')+count('EXTRAÇÃO INCOMPLETA')};}
