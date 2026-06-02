// admin-valores.js — módulo Valores de Refeição Homy · HOTFIX 4
(function(){
  function $(id){return document.getElementById(id)}
  function toastSafe(msg,type){ if(window.toast) toast(msg,type||'info'); else alert(msg); }
  function money(v){
    let s=String(v ?? '').trim();
    if(!s) return 0;
    s=s.replace(/R\$/gi,'').replace(/\s/g,'');
    if(s.includes(',')) s=s.replace(/\./g,'').replace(',','.');
    const n=Number(s);
    return Number.isFinite(n)?n:0;
  }
  function normalizeDisplayMoney(v){
    let n=Number(v||0);
    // Proteção contra registros salvos anteriormente como 2248 ao invés de 22,48.
    if(n>999 && String(v).indexOf('.')===-1) n=n/100;
    return n;
  }
  function fmtMoney(v){ return 'R$ ' + normalizeDisplayMoney(v).toFixed(2).replace('.',','); }
  function fmtDate(v){ return v ? String(v).slice(0,10).split('-').reverse().join('/') : ''; }
  function pick(obj,...keys){ for(const k of keys){ if(obj && obj[k]!==undefined && obj[k]!==null) return obj[k]; } return ''; }
  let editandoId=null;

  function ensureModule(){
    const nav=document.querySelector('.sidebar-nav');
    if(nav && !document.querySelector('[data-module="valores"]')){
      const before=document.querySelector('[data-module="relatorios"]');
      const item=document.createElement('div');item.className='nav-item';item.dataset.module='valores';item.innerHTML='<span class="nav-icon">💰</span><span class="nav-label">Valores</span>';
      nav.insertBefore(item,before||null);
      item.addEventListener('click',showValores);
    }
    if(!$('mod-valores')){
      const content=document.querySelector('.content'); if(!content)return;
      const div=document.createElement('div'); div.className='module'; div.id='mod-valores';
      div.innerHTML=`
        <div class="section-header"><div class="section-title">💰 Valores da Refeição</div><button class="btn-primary" id="btnNovoValorRefeicao" type="button">+ Novo valor</button></div>
        <div class="alert alert-info">Cadastre o valor cobrado pela Vascon e o valor descontado do colaborador por período.</div>
        <div class="table-wrap"><table class="table"><thead><tr><th>Título</th><th>Início</th><th>Fim</th><th>Valor Vascon</th><th>Desconto funcionário</th><th>Status</th><th>Ações</th></tr></thead><tbody id="valoresTable"><tr><td colspan="7" style="text-align:center;color:rgba(143,170,210,.4);padding:2rem">Carregando valores...</td></tr></tbody></table></div>`;
      content.appendChild(div);
    }
    if(!$('modalValorRefeicao')){
      document.body.insertAdjacentHTML('beforeend',`
      <div class="modal-overlay" id="modalValorRefeicao"><div class="modal-box"><div class="modal-header"><div class="modal-title" id="valorModalTitle">Valor da Refeição</div><div class="modal-close" id="closeValorRefeicao">✕</div></div><div class="modal-body"><div class="form-grid">
        <div class="form-group" style="grid-column:1/-1"><label class="form-label">Título</label><input class="form-input" id="valorTitulo" placeholder="Ex.: Maio"></div>
        <div class="form-group"><label class="form-label">Data início</label><input class="form-input" id="valorDataInicio" type="date"></div>
        <div class="form-group"><label class="form-label">Data fim</label><input class="form-input" id="valorDataFim" type="date"></div>
        <div class="form-group"><label class="form-label">Valor Vascon</label><input class="form-input" id="valorVascon" inputmode="decimal" placeholder="22,48"></div>
        <div class="form-group"><label class="form-label">Valor descontado funcionário</label><input class="form-input" id="valorDesconto" inputmode="decimal" placeholder="4,00"></div>
        <div class="form-group"><label class="form-label">Ativo</label><select class="form-select" id="valorAtivo"><option value="sim">Sim</option><option value="nao">Não</option></select></div>
        <div class="form-group" style="grid-column:1/-1"><label class="form-label">Observação</label><textarea class="form-textarea" id="valorObs"></textarea></div>
      </div></div><div class="modal-footer"><button class="btn-secondary" id="cancelValorRefeicao">Cancelar</button><button class="btn-success" id="salvarValorRefeicao">Salvar valor</button></div></div></div>`);
    }
    bind();
  }
  function resetForm(){
    editandoId=null;
    ['valorTitulo','valorDataInicio','valorDataFim','valorVascon','valorDesconto','valorObs'].forEach(id=>{const el=$(id);if(el)el.value='';});
    const a=$('valorAtivo'); if(a)a.value='sim';
    const t=$('valorModalTitle'); if(t)t.textContent='Valor da Refeição';
  }
  function abrirModal(){ resetForm(); $('modalValorRefeicao')?.classList.add('open'); }
  function fechar(){ $('modalValorRefeicao')?.classList.remove('open'); resetForm(); }
  function preencher(v){
    editandoId=v.id;
    const t=$('valorModalTitle'); if(t)t.textContent='Editar Valor da Refeição';
    $('valorTitulo').value=pick(v,'Title','Titulo')||'';
    $('valorDataInicio').value=String(pick(v,'Data_Inicio','DataInicio','Inicio')||'').slice(0,10);
    $('valorDataFim').value=String(pick(v,'Data_Fim','DataFim','Fim')||'').slice(0,10);
    $('valorVascon').value=normalizeDisplayMoney(pick(v,'Valor_Vascon','ValorVascon','Vascon')).toFixed(2).replace('.',',');
    $('valorDesconto').value=normalizeDisplayMoney(pick(v,'Valor_Desconto','Valor_Desconto_Funcionario','ValorDesconto','Desconto')).toFixed(2).replace('.',',');
    $('valorAtivo').value=window.SP && SP.isTrue(pick(v,'Ativo')) ? 'sim' : 'nao';
    $('valorObs').value=pick(v,'Observacao','Observação','Obs')||'';
    $('modalValorRefeicao')?.classList.add('open');
  }
  async function salvar(e){
    if(e){e.preventDefault();e.stopPropagation();}
    try{
      await SP.init();
      const dados={
        title:$('valorTitulo')?.value||'Valor refeição',
        dataInicio:$('valorDataInicio')?.value,
        dataFim:$('valorDataFim')?.value,
        valorVascon:money($('valorVascon')?.value),
        valorDesconto:money($('valorDesconto')?.value),
        observacao:$('valorObs')?.value||'',
        ativo:$('valorAtivo')?.value!=='nao'
      };
      if(!dados.dataInicio||!dados.dataFim){toastSafe('Informe data início e fim.','error');return;}
      if(editandoId) await SP.updateValorRefeicao(editandoId,dados); else await SP.createValorRefeicao(dados);
      fechar(); toastSafe('Valor salvo no SharePoint.','success'); await loadValores();
    }catch(err){console.error(err);toastSafe('Erro ao salvar valor: '+(err.message||err),'error');}
  }
  async function excluirValor(id){
    try{
      if(!confirm('Excluir este valor de refeição?')) return;
      await SP.init();
      if(SP.deleteValorRefeicao) await SP.deleteValorRefeicao(id); else await SP.deleteItem('Valores de Refeição',id);
      toastSafe('Valor excluído.','success'); await loadValores();
    }catch(err){console.error(err);toastSafe('Erro ao excluir valor: '+(err.message||err),'error');}
  }
  async function loadValores(){
    const tbody=$('valoresTable'); if(!tbody)return;
    try{
      await SP.init(); const vals=await SP.getValoresRefeicao(false);
      if(!vals.length){tbody.innerHTML='<tr><td colspan="7" style="text-align:center;color:rgba(143,170,210,.4);padding:2rem">Nenhum valor cadastrado.</td></tr>';return;}
      tbody.innerHTML=vals.map(v=>{
        const desconto = pick(v,'Valor_Desconto','Valor_Desconto_Funcionario','ValorDesconto','Desconto') || 0;
        const dataJson=encodeURIComponent(JSON.stringify(v));
        return `<tr><td>${pick(v,'Title','Titulo')}</td><td>${fmtDate(pick(v,'Data_Inicio','DataInicio','Inicio'))}</td><td>${fmtDate(pick(v,'Data_Fim','DataFim','Fim'))}</td><td>${fmtMoney(pick(v,'Valor_Vascon','ValorVascon','Vascon'))}</td><td>${fmtMoney(desconto)}</td><td><span class="badge ${SP.isTrue(pick(v,'Ativo'))?'badge-green':'badge-red'}">${SP.isTrue(pick(v,'Ativo'))?'Ativo':'Inativo'}</span></td><td><div class="table-actions"><button class="btn-icon" onclick="window.editarValorRefeicao('${dataJson}')">✏️</button><button class="btn-icon danger" onclick="window.excluirValorRefeicao('${v.id}')">🗑️</button></div></td></tr>`;
      }).join('');
    }catch(err){tbody.innerHTML='<tr><td colspan="7" style="text-align:center;color:rgba(255,120,120,.8);padding:2rem">Erro: '+(err.message||err)+'</td></tr>';}
  }
  function showValores(){
    document.querySelectorAll('.module').forEach(m=>m.classList.remove('active'));
    $('mod-valores')?.classList.add('active');
    document.querySelectorAll('.nav-item').forEach(i=>i.classList.toggle('active',i.dataset.module==='valores'));
    const t=$('topbarTitle'),s=$('topbarSub'); if(t)t.textContent='Valores'; if(s)s.textContent='Valor da refeição por período'; loadValores();
  }
  function bind(){
    const pairs=[['btnNovoValorRefeicao',abrirModal],['closeValorRefeicao',fechar],['cancelValorRefeicao',fechar]];
    for(const [id,fn] of pairs){ const el=$(id); if(el){ el.onclick=fn; } }
    const salvarBtn=$('salvarValorRefeicao'); if(salvarBtn){ salvarBtn.onclick=salvar; }
    const nav=document.querySelector('[data-module="valores"]'); if(nav){ nav.onclick=showValores; }
  }
  window.loadValores=loadValores;
  window.editarValorRefeicao=function(data){ try{preencher(JSON.parse(decodeURIComponent(data)));}catch(e){console.error(e);} };
  window.excluirValorRefeicao=excluirValor;
  window.salvarValorRefeicaoHomy=salvar;
  document.addEventListener('DOMContentLoaded',ensureModule); setTimeout(ensureModule,700); setTimeout(ensureModule,1600); setInterval(bind,1500);
})();
