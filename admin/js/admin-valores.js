// admin-valores.js — módulo Valores de Refeição Homy
(function(){
  function $(id){return document.getElementById(id)}
  function semana(){return window.getSemanaIdAtualSelecionada ? window.getSemanaIdAtualSelecionada() : (window.SP?.getCurrentWeekId ? SP.getCurrentWeekId() : '');}
  function toastSafe(msg,type){ if(window.toast) toast(msg,type||'info'); else alert(msg); }
  function money(v){const n=Number(String(v||0).replace(',','.'));return isNaN(n)?0:n;}
  function ensureModule(){
    const nav=document.querySelector('.sidebar-nav');
    if(nav && !document.querySelector('[data-module="valores"]')){
      const before=document.querySelector('[data-module="relatorios"]');
      const item=document.createElement('div');item.className='nav-item';item.dataset.module='valores';item.innerHTML='<span class="nav-icon">💰</span><span class="nav-label">Valores</span>';
      nav.insertBefore(item,before||null);
      item.addEventListener('click',()=>window.showModule?window.showModule('valores'):showValores());
    }
    if(!$("mod-valores")){
      const content=document.querySelector('.content'); if(!content)return;
      const div=document.createElement('div'); div.className='module'; div.id='mod-valores';
      div.innerHTML=`
        <div class="section-header"><div class="section-title">💰 Valores da Refeição</div><button class="btn-primary" id="btnNovoValorRefeicao" type="button">+ Novo valor</button></div>
        <div class="alert alert-info">Cadastre o valor cobrado pela Vascon e o valor descontado do colaborador por período.</div>
        <div class="table-wrap"><table class="table"><thead><tr><th>Título</th><th>Início</th><th>Fim</th><th>Valor Vascon</th><th>Desconto funcionário</th><th>Status</th><th>Ações</th></tr></thead><tbody id="valoresTable"><tr><td colspan="7" style="text-align:center;color:rgba(143,170,210,.4);padding:2rem">Carregando valores...</td></tr></tbody></table></div>`;
      content.appendChild(div);
    }
    if(!$("modalValorRefeicao")){
      document.body.insertAdjacentHTML('beforeend',`
      <div class="modal-overlay" id="modalValorRefeicao"><div class="modal-box"><div class="modal-header"><div class="modal-title">Valor da Refeição</div><div class="modal-close" id="closeValorRefeicao">✕</div></div><div class="modal-body"><div class="form-grid">
        <div class="form-group" style="grid-column:1/-1"><label class="form-label">Título</label><input class="form-input" id="valorTitulo" placeholder="Ex.: Maio"></div>
        <div class="form-group"><label class="form-label">Data início</label><input class="form-input" id="valorDataInicio" type="date"></div>
        <div class="form-group"><label class="form-label">Data fim</label><input class="form-input" id="valorDataFim" type="date"></div>
        <div class="form-group"><label class="form-label">Valor Vascon</label><input class="form-input" id="valorVascon" placeholder="22,48"></div>
        <div class="form-group"><label class="form-label">Valor descontado funcionário</label><input class="form-input" id="valorDesconto" placeholder="4,00"></div>
        <div class="form-group"><label class="form-label">Ativo</label><select class="form-select" id="valorAtivo"><option value="sim">Sim</option><option value="nao">Não</option></select></div>
        <div class="form-group" style="grid-column:1/-1"><label class="form-label">Observação</label><textarea class="form-textarea" id="valorObs"></textarea></div>
      </div></div><div class="modal-footer"><button class="btn-secondary" id="cancelValorRefeicao">Cancelar</button><button class="btn-success" id="salvarValorRefeicao">Salvar valor</button></div></div></div>`);
    }
    bind();
  }
  function abrirModal(){
    ['valorTitulo','valorDataInicio','valorDataFim','valorVascon','valorDesconto','valorObs'].forEach(id=>{const el=$(id);if(el)el.value='';});
    const a=$('valorAtivo'); if(a)a.value='sim';
    $('modalValorRefeicao')?.classList.add('open');
  }
  function fechar(){ $('modalValorRefeicao')?.classList.remove('open'); }
  async function salvar(){
    try{
      await SP.init();
      const dados={title:$('valorTitulo')?.value||'Valor refeição',dataInicio:$('valorDataInicio')?.value,dataFim:$('valorDataFim')?.value,valorVascon:money($('valorVascon')?.value),valorDesconto:money($('valorDesconto')?.value),observacao:$('valorObs')?.value||'',ativo:$('valorAtivo')?.value!=='nao'};
      if(!dados.dataInicio||!dados.dataFim){toastSafe('Informe data início e fim.','error');return;}
      await SP.createValorRefeicao(dados);
      fechar(); toastSafe('Valor salvo no SharePoint.','success'); await loadValores();
    }catch(e){console.error(e);toastSafe('Erro ao salvar valor: '+(e.message||e),'error');}
  }
  async function loadValores(){
    const tbody=$('valoresTable'); if(!tbody)return;
    try{await SP.init(); const vals=await SP.getValoresRefeicao(false); if(!vals.length){tbody.innerHTML='<tr><td colspan="7" style="text-align:center;color:rgba(143,170,210,.4);padding:2rem">Nenhum valor cadastrado.</td></tr>';return;}
      tbody.innerHTML=vals.map(v=>`<tr><td>${v.Title||''}</td><td>${String(v.Data_Inicio||'').slice(0,10).split('-').reverse().join('/')}</td><td>${String(v.Data_Fim||'').slice(0,10).split('-').reverse().join('/')}</td><td>R$ ${String(v.Valor_Vascon??'0').replace('.',',')}</td><td>R$ ${String(v.Valor_Desconto_Funcionario??'0').replace('.',',')}</td><td><span class="badge ${SP.isTrue(v.Ativo)?'badge-green':'badge-red'}">${SP.isTrue(v.Ativo)?'Ativo':'Inativo'}</span></td><td>—</td></tr>`).join('');
    }catch(e){tbody.innerHTML='<tr><td colspan="7" style="text-align:center;color:rgba(255,120,120,.8);padding:2rem">Erro: '+(e.message||e)+'</td></tr>';}
  }
  function showValores(){ document.querySelectorAll('.module').forEach(m=>m.classList.remove('active')); $('mod-valores')?.classList.add('active'); document.querySelectorAll('.nav-item').forEach(i=>i.classList.toggle('active',i.dataset.module==='valores')); const t=$('topbarTitle'),s=$('topbarSub'); if(t)t.textContent='Valores'; if(s)s.textContent='Valor da refeição por período'; loadValores(); }
  function bind(){
    $('btnNovoValorRefeicao')?.addEventListener('click',abrirModal); $('closeValorRefeicao')?.addEventListener('click',fechar); $('cancelValorRefeicao')?.addEventListener('click',fechar); $('salvarValorRefeicao')?.addEventListener('click',salvar);
    document.querySelector('[data-module="valores"]')?.addEventListener('click',showValores);
  }
  document.addEventListener('DOMContentLoaded',ensureModule); setTimeout(ensureModule,700); setTimeout(ensureModule,1600);
})();
