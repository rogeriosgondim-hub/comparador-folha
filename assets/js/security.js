(()=>{
  const nav=document.querySelector('.portal-nav');
  const main=document.querySelector('.portal-container');
  if(!nav||!main)return;

  const btn=document.createElement('button');
  btn.className='portal-tab';
  btn.type='button';
  btn.dataset.target='securityModule';
  btn.textContent='🔐 Segurança / Privacidade';
  nav.appendChild(btn);

  const panel=document.createElement('section');
  panel.id='securityModule';
  panel.className='portal-panel hidden';
  panel.innerHTML=`
    <div class="section-heading portal-heading">
      <div>
        <p class="eyebrow">Proteção de dados</p>
        <h2>Segurança / Privacidade</h2>
        <p>Resumo das proteções aplicadas ao portal e controles para limpar dados armazenados neste navegador.</p>
      </div>
    </div>

    <div class="portal-kpis compact">
      <article class="portal-kpi"><span>Processamento</span><strong>Local</strong><small>Excel e PDF são lidos no navegador.</small></article>
      <article class="portal-kpi"><span>Bibliotecas</span><strong>Locais</strong><small>Sem CDN em tempo de execução.</small></article>
      <article class="portal-kpi"><span>Conexões externas</span><strong>Bloqueadas</strong><small>CSP usa connect-src 'none'.</small></article>
    </div>

    <section class="card portal-card">
      <div class="section-heading table-heading"><div><p class="eyebrow">Armazenamento neste navegador</p><h2>O que pode permanecer salvo</h2></div></div>
      <div class="security-grid">
        <div class="security-item"><strong>Movimentações</strong><span>Usam <code>sessionStorage</code>: ficam disponíveis apenas na sessão atual da aba/janela e deixam de persistir após o encerramento da sessão.</span></div>
        <div class="security-item"><strong>Resumo da base e competência</strong><span>Podem usar armazenamento local para preferências e indicadores resumidos.</span></div>
        <div class="security-item"><strong>Histórico resumido do comparador</strong><span>Armazena apenas totais/indicadores; não deve conter nomes, CPFs, PDFs ou planilhas.</span></div>
      </div>
      <div class="info-note" id="securityStorageStatus">Calculando armazenamento local...</div>
      <div class="actions">
        <button id="clearPortalData" class="secondary" type="button">Apagar dados locais e da sessão</button>
      </div>
    </section>

    <section class="card portal-card">
      <div class="section-heading table-heading"><div><p class="eyebrow">Boas práticas</p><h2>Uso recomendado</h2></div></div>
      <div class="security-grid">
        <div class="security-item"><strong>Não publicar arquivos de RH</strong><span>Nunca envie PDF, Excel, CSV, backups ou exportações com dados pessoais para o repositório público.</span></div>
        <div class="security-item"><strong>Use dispositivo confiável</strong><span>Evite usar o portal em computador compartilhado ou sem bloqueio de tela.</span></div>
        <div class="security-item"><strong>Feche a sessão ao terminar</strong><span>Ao encerrar o navegador/aba, dados de movimentações mantidos em sessão deixam de ficar disponíveis.</span></div>
      </div>
    </section>`;
  main.appendChild(panel);

  const style=document.createElement('style');
  style.textContent=`
    .security-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}
    .security-item{border:1px solid var(--portal-line);border-radius:10px;background:#fbfdff;padding:14px;display:flex;flex-direction:column;gap:6px}
    .security-item strong{color:var(--portal-blue-dark)}
    .security-item span{color:var(--portal-muted);font-size:.9rem;line-height:1.45}
    .security-item code{font-size:.85em}
    @media(max-width:900px){.security-grid{grid-template-columns:1fr}}
  `;
  document.head.appendChild(style);

  function showPanel(){
    document.querySelectorAll('.portal-panel').forEach(x=>x.classList.toggle('hidden',x.id!=='securityModule'));
    document.querySelectorAll('.portal-tab').forEach(x=>x.classList.toggle('active',x===btn));
    updateStatus();
  }
  btn.addEventListener('click',showPanel);

  function updateStatus(){
    const localKeys=Object.keys(localStorage);
    const sessionKeys=Object.keys(sessionStorage);
    const target=document.getElementById('securityStorageStatus');
    if(target)target.textContent=`Armazenamento detectado neste site: ${localKeys.length} item(ns) persistente(s) e ${sessionKeys.length} item(ns) de sessão.`;
  }

  const clearBtn=panel.querySelector('#clearPortalData');
  clearBtn.addEventListener('click',()=>{
    const ok=confirm('Apagar todos os dados locais e de sessão deste portal neste navegador? Esta ação não altera arquivos no seu computador nem no GitHub.');
    if(!ok)return;
    localStorage.clear();
    sessionStorage.clear();
    alert('Dados locais do portal apagados. A página será recarregada.');
    location.reload();
  });

  updateStatus();
})();
