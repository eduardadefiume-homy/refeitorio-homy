/* ============================================================
   admin-extensions.js — Refeitório Homy · melhorias gerenciais
   Recursos incluídos:
   - Centro de custo no colaborador
   - Editar/excluir/desativar em colaboradores, pedidos, extras e cardápio
   - Cadastro de valores Vascon/desconto funcionário
   - Cadastro de ausências por período
   - Relatórios por período, centro de custo e colaborador
   - Operação do Dia com Principal, Light, Carne, Massa e Lanche
   ============================================================ */
(function(){
  'use strict';

  const DIAS = ["segunda","terca","quarta","quinta","sexta"];
  const DIA_LABEL = {segunda:"Segunda",terca:"Terça",quarta:"Quarta",quinta:"Quinta",sexta:"Sexta"};
  const OPCOES = ["principal","light","carne","massa","lanche"];
  const OP_LABEL = {principal:"Principal",light:"Light",carne:"Carne",massa:"Massa",lanche:"Lanche"};
  const STATUS_CANCELADOS_EXT = ["Cancelado","Afastado","Férias","Bloqueado","Não vai almoçar"];
  const STATUS_VALIDOS = ["Confirmado","Cancelado","Afastado","Férias","Bloqueado","Extra","Não vai almoçar"];

  function $(id){ return document.getElementById(id); }
  function safe(v){ return String(v ?? "").replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }
  function norm(v){ return String(v ?? "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,""); }
  function money(v){ return Number(v || 0).toLocaleString("pt-BR",{style:"currency",currency:"BRL"}); }
  function toast2(msg,type="info"){
    if(typeof window.toast === "function") return window.toast(msg,type);
    alert(msg);
  }
  function semanaAtual(){
    if(typeof window.getSemanaIdAtualSelecionada === "function") return window.getSemanaIdAtualSelecionada();
    if(typeof window.getSemanaId === "function") return window.getSemanaId();
    return new Date().getFullYear()+"-W"+String(1).padStart(2,"0");
  }
  function getPedidoStatus(p){
    if(p.Status) return p.Status;
    if(p.status) return p.status;
    return (p.Confirmado || p.confirmado) ? "Confirmado" : "Pendente";
  }
  function pedidoConta(p){
    return !STATUS_CANCELADOS_EXT.includes(getPedidoStatus(p));
  }
  function currentUser(){
    try{ return SP.getUserName ? SP.getUserName() : "Admin"; }catch(e){ return "Admin"; }
  }

  function parseSemanaId(semanaId){
    const m = String(semanaId || "").match(/(\d{4})-?W?(\d{1,2})/i);
    if(!m) return null;
    const year = Number(m[1]);
    const week = Number(m[2]);
    const jan4 = new Date(Date.UTC(year,0,4));
    const jan4Day = jan4.getUTCDay() || 7;
    const monday = new Date(jan4);
    monday.setUTCDate(jan4.getUTCDate() - jan4Day + 1 + (week - 1) * 7);
    return monday;
  }
  function pedidoDate(p){
    const monday = parseSemanaId(p.Semana_id || p.semana_id || p.Semana || p.semana || semanaAtual());
    if(!monday) return null;
    const idx = DIAS.indexOf(norm(p.Dia || p.dia));
    if(idx < 0) return monday;
    const d = new Date(monday);
    d.setUTCDate(monday.getUTCDate()+idx);
    return d;
  }
  function dateOnly(d){ return d ? d.toISOString().slice(0,10) : ""; }
  function inRange(date, ini, fim){
    if(!date) return true;
    const iso = dateOnly(date);
    return (!ini || iso >= ini) && (!fim || iso <= fim);
  }

  function patchModulesAndNav(){
    try{
      if(typeof MODULES !== "undefined"){
        MODULES.valores = {title:"Valores",sub:"Valores da Vascon e desconto dos colaboradores"};
        MODULES.ausencias = {title:"Ausências",sub:"Férias, afastamentos e períodos sem refeição"};
      }
    }catch(e){}

    const nav = document.querySelector(".sidebar-nav");
    if(nav && !document.querySelector('[data-module="valores"]')){
      const rel = document.querySelector('[data-module="relatorios"]');
      const valores = document.createElement("div");
      valores.className = "nav-item";
      valores.dataset.module = "valores";
      valores.innerHTML = '<span class="nav-icon">💰</span><span class="nav-label">Valores</span>';
      const aus = document.createElement("div");
      aus.className = "nav-item";
      aus.dataset.module = "ausencias";
      aus.innerHTML = '<span class="nav-icon">🚫</span><span class="nav-label">Ausências</span>';
      nav.insertBefore(valores, rel || null);
      nav.insertBefore(aus, rel || null);
      [valores,aus].forEach(item=>item.addEventListener("click",()=>abrirModulo(item.dataset.module)));
    }

    const content = document.querySelector(".content");
    if(content && !$('mod-valores')){
      const mod = document.createElement("div");
      mod.className = "module";
      mod.id = "mod-valores";
      mod.innerHTML = `
        <div class="section-header">
          <div>
            <div class="section-title">💰 Valores das Refeições</div>
            <div style="font-size:0.78rem;color:rgba(143,170,210,0.58);margin-top:0.25rem">Cadastro usado para comparar a NF da Vascon e calcular o desconto em folha.</div>
          </div>
          <button class="btn-primary" type="button" onclick="abrirModalValorRefeicao()">+ Novo valor</button>
        </div>
        <div class="alert alert-info">O valor descontado do funcionário é único para todos. Se mudar em algum mês, cadastre um novo período.</div>
        <div class="table-wrap"><table class="table"><thead><tr><th>Período</th><th>Data início</th><th>Data fim</th><th>Vascon</th><th>Desconto funcionário</th><th>Status</th><th>Ações</th></tr></thead><tbody id="valoresTable"><tr><td colspan="7" style="text-align:center;padding:2rem;color:rgba(143,170,210,0.4)">Carregando valores...</td></tr></tbody></table></div>`;
      content.insertBefore(mod, $('mod-relatorios') || null);
    }
    if(content && !$('mod-ausencias')){
      const mod = document.createElement("div");
      mod.className = "module";
      mod.id = "mod-ausencias";
      mod.innerHTML = `
        <div class="section-header">
          <div>
            <div class="section-title">🚫 Ausências / Bloqueios</div>
            <div style="font-size:0.78rem;color:rgba(143,170,210,0.58);margin-top:0.25rem">Períodos em que o colaborador fica indisponível para marcação de refeição.</div>
          </div>
          <button class="btn-primary" type="button" onclick="abrirModalAusencia()">+ Nova ausência</button>
        </div>
        <div class="alert alert-info">No refeitório, o colaborador deve aparecer como: <strong>Período de DD/MM/AAAA a DD/MM/AAAA</strong>, sem fixar o motivo.</div>
        <div class="table-wrap"><table class="table"><thead><tr><th>Colaborador</th><th>Período</th><th>Motivo interno</th><th>Status</th><th>Observação</th><th>Ações</th></tr></thead><tbody id="ausenciasTable"><tr><td colspan="6" style="text-align:center;padding:2rem;color:rgba(143,170,210,0.4)">Carregando ausências...</td></tr></tbody></table></div>`;
      content.insertBefore(mod, $('mod-relatorios') || null);
    }
  }

  function abrirModulo(mod){
    document.querySelectorAll(".nav-item").forEach(i=>i.classList.remove("active"));
    document.querySelectorAll(".module").forEach(m=>m.classList.remove("active"));
    const nav = document.querySelector(`[data-module="${mod}"]`);
    const module = $(`mod-${mod}`);
    if(nav) nav.classList.add("active");
    if(module) module.classList.add("active");
    const titles = {
      valores:["Valores","Valores da Vascon e desconto dos colaboradores"],
      ausencias:["Ausências","Férias, afastamentos e períodos sem refeição"]
    };
    if(titles[mod]){
      if($('topbarTitle')) $('topbarTitle').textContent = titles[mod][0];
      if($('topbarSub')) $('topbarSub').textContent = titles[mod][1];
    }
    loadModuleExtended(mod);
  }

  const oldLoadModule = window.loadModule;
  window.loadModule = function(mod){
    if(mod === "valores" || mod === "ausencias") return loadModuleExtended(mod);
    if(oldLoadModule) oldLoadModule(mod);
    if(mod === "relatorios") setTimeout(loadRelatoriosAvancados, 50);
  };
  function loadModuleExtended(mod){
    if(mod === "valores") return loadValoresRefeicao();
    if(mod === "ausencias") return loadAusencias();
    if(mod === "relatorios") return loadRelatoriosAvancados();
  }

  function patchColaboradorUI(){
    const headRow = document.querySelector('#mod-colaboradores thead tr');
    if(headRow && !headRow.querySelector('[data-col="centro"]')){
      const th = document.createElement('th'); th.dataset.col='centro'; th.textContent='Centro de custo';
      headRow.insertBefore(th, headRow.children[2] || null);
    }
    const formGrid = document.querySelector('#modalColaborador .form-grid') || document.querySelector('#modalColab .form-grid');
    if(formGrid && !$('colabCentroCusto')){
      const group = document.createElement('div');
      group.className = 'form-group';
      group.innerHTML = '<label class="form-label">Centro de custo</label><input class="form-input" id="colabCentroCusto" placeholder="Ex.: TI, RH, Comercial, PCP">';
      const tipo = $('colabTipo')?.closest('.form-group');
      formGrid.insertBefore(group, tipo || null);
    }
  }

  window.abrirModalColaborador = function(id=null){
    patchColaboradorUI();
    const modal = $('modalColaborador') || $('modalColab');
    if(!modal) return;
    modal.dataset.editId = id || "";
    const title = modal.querySelector('.modal-title');
    if(title) title.textContent = id ? "Editar colaborador" : "Novo colaborador";
    ['colabNome','colabDepartamento','colabDept','colabEmail','colabCentroCusto'].forEach(x=>{ if($(x)) $(x).value=''; });
    if($('colabTipo')) $('colabTipo').value = 'colaborador';
    modal.classList.add('open');
  };

  window.editarColaborador = async function(id){
    window.abrirModalColaborador(id);
    const colabs = await SP.getItems('Colaboradores');
    const c = colabs.find(x=>String(x.id)===String(id));
    if(!c) return;
    if($('colabNome')) $('colabNome').value = c.Nome || c.Title || '';
    if($('colabDepartamento')) $('colabDepartamento').value = c.Departamento || '';
    if($('colabDept')) $('colabDept').value = c.Departamento || '';
    if($('colabEmail')) $('colabEmail').value = c.Email || '';
    if($('colabTipo')) $('colabTipo').value = c.tipo || c.Tipo || 'colaborador';
    if($('colabCentroCusto')) $('colabCentroCusto').value = c.Centro_Custo || '';
  };

  window.salvarNovoColaborador = async function(){
    const modal = $('modalColaborador') || $('modalColab');
    const id = modal?.dataset?.editId || '';
    const nome = ($('colabNome')?.value || '').trim();
    const departamento = ($('colabDepartamento')?.value || $('colabDept')?.value || '').trim();
    const email = ($('colabEmail')?.value || '').trim();
    const tipo = $('colabTipo')?.value || 'colaborador';
    const centroCusto = ($('colabCentroCusto')?.value || '').trim();
    if(!nome){ toast2('Informe o nome do colaborador.','error'); return; }
    if(id){
      await SP.updateColaborador(id,{nome,departamento,email,tipo,centroCusto,Ativo:true});
      toast2('Colaborador atualizado.','success');
    }else{
      await SP.createColaborador({nome,departamento,email,tipo,centroCusto});
      toast2('Colaborador cadastrado.','success');
    }
    modal?.classList.remove('open');
    await window.loadColaboradores();
  };

  window.loadColaboradores = async function(){
    patchColaboradorUI();
    const tbody = $('colabTable');
    if(!tbody) return;
    try{
      const colabs = await SP.getColaboradores();
      if(!colabs.length){ tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:rgba(143,170,210,0.4);padding:2rem">Nenhum colaborador cadastrado.</td></tr>'; return; }
      tbody.innerHTML = colabs.map(c=>{
        const id = c.id || '';
        const nome = c.Nome || c.nome || c.Title || '';
        const dept = c.Departamento || c.departamento || '';
        const cc = c.Centro_Custo || c.centroCusto || '';
        const tipo = c.tipo || c.Tipo || 'colaborador';
        return `<tr><td>${safe(nome)}</td><td>${safe(dept)}</td><td>${safe(cc) || '—'}</td><td><span class="badge badge-blue">${safe(tipo)}</span></td><td><span class="badge badge-green">Ativo</span></td><td><div class="table-actions"><button class="btn-icon" title="Editar" onclick="editarColaborador('${id}')">✏️</button><button class="btn-icon danger" title="Desativar" onclick="desativarColab('${id}')">🗑️</button></div></td></tr>`;
      }).join('');
    }catch(e){ tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:rgba(255,100,100,0.6);padding:2rem">Erro: ${safe(e.message)}</td></tr>`; }
  };

  window.desativarColab = async function(id){
    if(!confirm('Desativar este colaborador?')) return;
    await SP.desativarColaborador(id);
    toast2('Colaborador desativado.','success');
    await window.loadColaboradores();
  };

  async function getValorVigente(dataISO){
    try{
      const valores = await SP.getValoresRefeicao(false);
      const d = dataISO || dateOnly(new Date());
      return valores.find(v=>SP.isTrue(v.Ativo) && (!v.Data_Inicio || String(v.Data_Inicio).slice(0,10)<=d) && (!v.Data_Fim || String(v.Data_Fim).slice(0,10)>=d)) || valores.find(v=>SP.isTrue(v.Ativo)) || null;
    }catch(e){ return null; }
  }

  function patchOperacaoCards(){
    const grid = document.querySelector('#mod-operacao .stats-grid');
    if(!grid || $('opTotalCarne')) return;
    const cards = [
      ['opTotalCarne','🥩','Carne'],['opTotalMassa','🍝','Massa'],['opTotalLanche','🍔','Lanche']
    ].map(([id,ico,label])=>`<div class="stat-card"><div class="stat-icon">${ico}</div><div class="stat-value" id="${id}">—</div><div class="stat-label">${label}</div></div>`).join('');
    grid.insertAdjacentHTML('beforeend',cards);
  }
  window.setOperacaoTotals = function(pedidosDia){
    patchOperacaoCards();
    const validos = pedidosDia.filter(p=>pedidoConta(p));
    const cancelados = pedidosDia.filter(p=>!pedidoConta(p));
    const count = op => validos.filter(p=>norm(p.Opcao || p.opcao)===op).length;
    if($('opTotalConfirmado')) $('opTotalConfirmado').textContent = validos.length;
    if($('opTotalPrincipal')) $('opTotalPrincipal').textContent = count('principal');
    if($('opTotalLight')) $('opTotalLight').textContent = count('light');
    if($('opTotalCarne')) $('opTotalCarne').textContent = count('carne');
    if($('opTotalMassa')) $('opTotalMassa').textContent = count('massa');
    if($('opTotalLanche')) $('opTotalLanche').textContent = count('lanche');
    if($('opTotalCancelado')) $('opTotalCancelado').textContent = cancelados.length;
  };

  const oldLoadCardapioAtual = window.loadCardapioAtual;
  window.loadCardapioAtual = async function(semanaId){
    const wrap = $('cardapioAtual');
    if(!wrap) return oldLoadCardapioAtual ? oldLoadCardapioAtual(semanaId) : null;
    try{
      const items = await SP.getCardapio(semanaId);
      if(!items.length){ wrap.innerHTML = "<div class='alert alert-warning'>Nenhum cardápio cadastrado para esta semana.</div>"; return; }
      let out = "<div style='display:flex;flex-direction:column;gap:0.6rem'>";
      DIAS.forEach(dia=>{
        const diaItems = items.filter(i=>norm(i.Dia || i.dia)===dia);
        if(!diaItems.length) return;
        out += `<div style="border-radius:12px;border:1px solid rgba(255,255,255,0.08);background:rgba(255,255,255,0.03);padding:0.8rem 1rem"><div style="font-family:Barlow Condensed,sans-serif;font-size:1rem;font-weight:700;text-transform:uppercase;color:#fff;margin-bottom:0.5rem">${DIA_LABEL[dia]}</div><div style="display:flex;gap:0.5rem;flex-wrap:wrap">`;
        diaItems.forEach(op=>{
          const id = op.id || '';
          out += `<span class="badge badge-blue" style="display:inline-flex;gap:0.4rem;align-items:center">${safe(OP_LABEL[norm(op.Opcao||op.opcao)] || op.Opcao || op.opcao || '')} - ${safe(op.Nome_Prato || op.nome_prato || '')}<button class="btn-icon" style="width:22px;height:22px" title="Editar" onclick="editarCardapio('${id}')">✏️</button><button class="btn-icon danger" style="width:22px;height:22px" title="Excluir" onclick="excluirCardapio('${id}')">🗑️</button></span>`;
        });
        out += "</div></div>";
      });
      wrap.innerHTML = out + "</div>";
    }catch(e){ wrap.innerHTML = `<div class='alert alert-warning'>Erro: ${safe(e.message)}</div>`; }
  };
  window.editarCardapio = async function(id){
    const items = await SP.getItems('Cardapio');
    const c = items.find(x=>String(x.id)===String(id));
    if(!c) return;
    const nome = prompt('Nome do prato:', c.Nome_Prato || '');
    if(nome === null) return;
    const detalhes = prompt('Detalhes/observação:', c.Detalhes || '') ?? '';
    await SP.updateItem('Cardapio', id, {Nome_Prato:nome, Detalhes:detalhes});
    toast2('Cardápio atualizado.','success');
    await window.loadCardapioAtual(semanaAtual());
  };
  window.excluirCardapio = async function(id){
    if(!confirm('Excluir este item do cardápio?')) return;
    await SP.deleteItem('Cardapio', id);
    toast2('Item excluído.','success');
    await window.loadCardapioAtual(semanaAtual());
  };

  window.loadPedidos = async function(semanaId){
    const tbody = $('pedidosTable');
    if(!tbody) return;
    try{
      let pedidos = await SP.getPedidos(semanaId);
      const diaFiltro = $('filtroDia')?.value || '';
      const opFiltro = $('filtroOpcao')?.value || '';
      if(diaFiltro) pedidos = pedidos.filter(p=>norm(p.Dia||p.dia)===norm(diaFiltro));
      if(opFiltro) pedidos = pedidos.filter(p=>norm(p.Opcao||p.opcao)===norm(opFiltro));
      if(!pedidos.length){ tbody.innerHTML='<tr><td colspan="6" style="text-align:center;color:rgba(143,170,210,0.4);padding:2rem">Nenhum pedido encontrado.</td></tr>'; return; }
      tbody.innerHTML = pedidos.map(p=>{
        const id = p.id || '';
        const status = getPedidoStatus(p);
        return `<tr><td>${safe(p.Colaborador_nome||p.colaborador_nome||'')}</td><td>${safe(DIA_LABEL[norm(p.Dia||p.dia)] || p.Dia || p.dia || '')}</td><td><span class="badge badge-blue">${safe(OP_LABEL[norm(p.Opcao||p.opcao)] || p.Opcao || p.opcao || '')}</span></td><td>${safe(p.Nome_Prato||p.nome_prato||'—')}</td><td><span class="badge ${pedidoConta(p)?'badge-green':'badge-red'}">${safe(status)}</span></td><td><div class="table-actions"><button class="btn-icon" title="Editar" onclick="abrirModalEditarPedido('${id}','${safe(p.Colaborador_nome||'')}','${safe(p.Dia||'')}','${safe(p.Opcao||'')}')">✏️</button><button class="btn-icon danger" title="Excluir" onclick="removerPedido('${id}')">🗑️</button></div></td></tr>`;
      }).join('');
    }catch(e){ tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:rgba(255,100,100,0.6);padding:2rem">Erro: ${safe(e.message)}</td></tr>`; }
  };

  window.loadExtras = async function(semanaId){
    const tbody = $('extrasTable');
    if(!tbody) return;
    try{
      const extras = await SP.getExtras(semanaId);
      if(!extras.length){ tbody.innerHTML='<tr><td colspan="6" style="text-align:center;color:rgba(143,170,210,0.4);padding:2rem">Nenhum extra cadastrado.</td></tr>'; return; }
      tbody.innerHTML = extras.map(e=>{
        const id = e.id || '';
        return `<tr><td>${safe(e.Nome||e.nome||e.Title||'')}</td><td><span class="badge badge-yellow">${safe(e.tipo||e.Tipo||'')}</span></td><td>${safe(DIA_LABEL[norm(e.Dia||e.dia)] || e.Dia || e.dia || '')}</td><td><span class="badge badge-blue">${safe(OP_LABEL[norm(e.Opcao||e.opcao)] || e.Opcao || e.opcao || '')}</span></td><td style="font-size:0.78rem;color:rgba(143,170,210,0.6)">${safe(e.Observacao||e.observacao||'—')}</td><td><div class="table-actions"><button class="btn-icon" title="Editar" onclick="editarExtra('${id}')">✏️</button><button class="btn-icon danger" title="Excluir" onclick="removerExtra('${id}')">🗑️</button></div></td></tr>`;
      }).join('');
    }catch(e){ tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:rgba(255,100,100,0.6);padding:2rem">Erro: ${safe(e.message)}</td></tr>`; }
  };
  window.editarExtra = async function(id){
    const extras = await SP.getExtras(semanaAtual());
    const e = extras.find(x=>String(x.id)===String(id));
    if(!e) return;
    if(typeof window.abrirModalExtra === 'function') window.abrirModalExtra();
    const modal = $('modalExtra'); if(modal) modal.dataset.editId = id;
    if($('extraNome')) $('extraNome').value = e.Nome || e.Title || '';
    if($('extraTipo')) $('extraTipo').value = e.tipo || e.Tipo || 'visitante';
    if($('extraDia')) $('extraDia').value = norm(e.Dia || e.dia) || 'segunda';
    if($('extraOpcao')) $('extraOpcao').value = norm(e.Opcao || e.opcao) || 'principal';
    if($('extraObs')) $('extraObs').value = e.Observacao || '';
  };
  window.removerExtra = async function(id){
    if(!confirm('Excluir este extra?')) return;
    await SP.removeExtra(id);
    toast2('Extra excluído.','success');
    await window.loadExtras(semanaAtual());
  };

  const oldSalvarExtra = window.salvarExtra;
  window.salvarExtra = async function(){
    const modal = $('modalExtra');
    const id = modal?.dataset?.editId || '';
    if(!id && oldSalvarExtra) return oldSalvarExtra();
    await SP.updateItem('Extras', id, {
      Nome:($('extraNome')?.value||'').trim(),
      Title:($('extraNome')?.value||'').trim(),
      tipo:$('extraTipo')?.value||'visitante',
      Dia:$('extraDia')?.value||'segunda',
      Opcao:$('extraOpcao')?.value||'principal',
      Observacao:$('extraObs')?.value||''
    });
    delete modal.dataset.editId;
    modal?.classList.remove('open');
    toast2('Extra atualizado.','success');
    await window.loadExtras(semanaAtual());
  };

  function installReportUI(){
    const mod = $('mod-relatorios');
    if(!mod || mod.dataset.advanced === '1') return;
    mod.dataset.advanced = '1';
    mod.innerHTML = `
      <div class="section-header"><div><div class="section-title">📈 Relatórios Gerenciais</div><div style="font-size:0.78rem;color:rgba(143,170,210,0.58);margin-top:0.25rem">Quantidade, centro de custo, colaborador, valores e comparação com NF Vascon.</div></div><button class="btn-secondary" type="button" onclick="exportarRelatorioCSV()">📥 Exportar CSV</button></div>
      <div class="form-grid" style="margin-bottom:1rem"><div class="form-group"><label class="form-label">Data inicial</label><input class="form-input" type="date" id="relDataInicio"></div><div class="form-group"><label class="form-label">Data final</label><input class="form-input" type="date" id="relDataFim"></div><div class="form-group"><label class="form-label">NF Vascon recebida</label><input class="form-input" type="number" step="0.01" id="relValorNF" placeholder="Ex.: 1234,56"></div><div class="form-group" style="justify-content:end"><button class="btn-primary" type="button" onclick="loadRelatoriosAvancados()">🔎 Gerar relatório</button></div></div>
      <div class="stats-grid"><div class="stat-card"><div class="stat-icon">🍽️</div><div class="stat-value" id="relTotalGeral">—</div><div class="stat-label">Refeições no período</div></div><div class="stat-card"><div class="stat-icon">🏢</div><div class="stat-value" id="relCustoVascon">—</div><div class="stat-label">Custo Vascon estimado</div></div><div class="stat-card"><div class="stat-icon">👤</div><div class="stat-value" id="relDescontoFolha">—</div><div class="stat-label">Desconto funcionários</div></div><div class="stat-card"><div class="stat-icon">🧾</div><div class="stat-value" id="relDiferencaNF">—</div><div class="stat-label">Diferença NF x calculado</div></div></div>
      <div class="stats-grid"><div class="stat-card"><div class="stat-icon">🍗</div><div class="stat-value" id="rel-principal">—</div><div class="stat-label">Principal</div></div><div class="stat-card"><div class="stat-icon">🥗</div><div class="stat-value" id="rel-light">—</div><div class="stat-label">Light</div></div><div class="stat-card"><div class="stat-icon">🥩</div><div class="stat-value" id="rel-carne">—</div><div class="stat-label">Carne</div></div><div class="stat-card"><div class="stat-icon">🍝</div><div class="stat-value" id="rel-massa">—</div><div class="stat-label">Massa</div></div><div class="stat-card"><div class="stat-icon">🍔</div><div class="stat-value" id="rel-lanche">—</div><div class="stat-label">Lanche</div></div></div>
      <div style="margin-top:1rem"><div class="section-title" style="font-size:1rem;margin-bottom:0.8rem">Resumo por dia</div><div class="table-wrap"><table class="table"><thead><tr><th>Dia</th><th>Principal</th><th>Light</th><th>Carne</th><th>Massa</th><th>Lanche</th><th>Total</th></tr></thead><tbody id="relTable"></tbody></table></div></div>
      <div style="margin-top:1rem"><div class="section-title" style="font-size:1rem;margin-bottom:0.8rem">Valor total por centro de custo</div><div class="table-wrap"><table class="table"><thead><tr><th>Centro de custo</th><th>Quantidade</th><th>Valor Vascon</th><th>Desconto funcionários</th><th>Rateio NF</th></tr></thead><tbody id="relCentroCustoTable"></tbody></table></div></div>
      <div style="margin-top:1rem"><div class="section-title" style="font-size:1rem;margin-bottom:0.8rem">Quantidade por colaborador</div><div class="table-wrap"><table class="table"><thead><tr><th>Colaborador</th><th>Centro de custo</th><th>Quantidade</th><th>Valor Vascon</th><th>Desconto folha</th></tr></thead><tbody id="relColaboradorTable"></tbody></table></div></div>`;
    const today = new Date();
    const first = new Date(today.getFullYear(), today.getMonth(), 1);
    if(!$('relDataInicio').value) $('relDataInicio').value = dateOnly(first);
    if(!$('relDataFim').value) $('relDataFim').value = dateOnly(today);
  }

  window.loadRelatoriosAvancados = async function(){
    installReportUI();
    try{
      const ini = $('relDataInicio')?.value || '';
      const fim = $('relDataFim')?.value || '';
      const nf = Number($('relValorNF')?.value || 0);
      const semana = semanaAtual();
      let pedidos = await SP.getPedidos(semana);
      pedidos = pedidos.filter(p=>pedidoConta(p) && inRange(pedidoDate(p), ini, fim));
      const valor = await getValorVigente(ini || fim || dateOnly(new Date()));
      const vascon = Number(valor?.Valor_Vascon || valor?.valorVascon || 0);
      const desconto = Number(valor?.Valor_Desconto_Funcionario || valor?.valorDescontoFuncionario || 0);
      const total = pedidos.length;
      const custo = total * vascon;
      const descontoTotal = total * desconto;
      $('relTotalGeral').textContent = total;
      $('relCustoVascon').textContent = money(custo);
      $('relDescontoFolha').textContent = money(descontoTotal);
      $('relDiferencaNF').textContent = nf ? money(nf - custo) : '—';
      OPCOES.forEach(op=>{ const el=$('rel-'+op); if(el) el.textContent = pedidos.filter(p=>norm(p.Opcao||p.opcao)===op).length; });
      const byDia = {};
      DIAS.forEach(d=>byDia[d]={principal:0,light:0,carne:0,massa:0,lanche:0,total:0});
      pedidos.forEach(p=>{ const d=norm(p.Dia||p.dia); const op=norm(p.Opcao||p.opcao); if(byDia[d] && byDia[d][op]!==undefined){ byDia[d][op]++; byDia[d].total++; }});
      $('relTable').innerHTML = DIAS.map(d=>`<tr><td>${DIA_LABEL[d]}</td>${OPCOES.map(op=>`<td>${byDia[d][op]}</td>`).join('')}<td>${byDia[d].total}</td></tr>`).join('');
      const byCC = {}; const byColab = {};
      pedidos.forEach(p=>{
        const cc = p.Centro_Custo || p.centro_custo || 'Sem centro de custo';
        const nome = p.Colaborador_nome || p.colaborador_nome || 'Sem nome';
        byCC[cc] = (byCC[cc] || 0) + 1;
        const key = nome+'|'+cc; byColab[key] = (byColab[key] || 0) + 1;
      });
      $('relCentroCustoTable').innerHTML = Object.entries(byCC).sort((a,b)=>b[1]-a[1]).map(([cc,q])=>`<tr><td>${safe(cc)}</td><td>${q}</td><td>${money(q*vascon)}</td><td>${money(q*desconto)}</td><td>${nf?money(nf*(q/Math.max(total,1))):'—'}</td></tr>`).join('') || '<tr><td colspan="5" style="text-align:center;padding:2rem;color:rgba(143,170,210,0.4)">Sem dados no período.</td></tr>';
      $('relColaboradorTable').innerHTML = Object.entries(byColab).sort((a,b)=>b[1]-a[1]).map(([key,q])=>{ const [nome,cc]=key.split('|'); return `<tr><td>${safe(nome)}</td><td>${safe(cc)}</td><td>${q}</td><td>${money(q*vascon)}</td><td>${money(q*desconto)}</td></tr>`; }).join('') || '<tr><td colspan="5" style="text-align:center;padding:2rem;color:rgba(143,170,210,0.4)">Sem dados no período.</td></tr>';
      window.__ultimoRelatorioCSV = {pedidos,vascon,desconto,total,custo,descontoTotal};
    }catch(e){ toast2('Erro ao gerar relatório: '+e.message,'error'); }
  };
  window.loadRelatorios = window.loadRelatoriosAvancados;

  window.exportarRelatorioCSV = function(){
    const linhas = [['Colaborador','Centro de custo','Dia','Opção','Status','Valor Vascon','Desconto funcionário']];
    const dados = window.__ultimoRelatorioCSV || {pedidos:[],vascon:0,desconto:0};
    dados.pedidos.forEach(p=>linhas.push([p.Colaborador_nome||'',p.Centro_Custo||'',DIA_LABEL[norm(p.Dia||'')]||p.Dia||'',OP_LABEL[norm(p.Opcao||'')]||p.Opcao||'',getPedidoStatus(p),String(dados.vascon).replace('.',','),String(dados.desconto).replace('.',',')]));
    const csv = linhas.map(l=>l.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(';')).join('\n');
    const blob = new Blob([csv],{type:'text/csv;charset=utf-8'});
    const url = URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download='relatorio-refeitorio.csv'; a.click(); URL.revokeObjectURL(url);
  };

  function ensureValueModal(){
    if($('modalValorRefeicao')) return;
    document.body.insertAdjacentHTML('beforeend',`<div class="modal-overlay" id="modalValorRefeicao"><div class="modal-box"><div class="modal-header"><div class="modal-title">Valor da refeição</div><button class="modal-close" type="button" onclick="fecharModalValorRefeicao()">×</button></div><div class="modal-body"><div class="form-grid"><div class="form-group"><label class="form-label">Título</label><input class="form-input" id="valorTitulo" placeholder="Ex.: Junho 2026"></div><div class="form-group"><label class="form-label">Data início</label><input class="form-input" type="date" id="valorDataInicio"></div><div class="form-group"><label class="form-label">Data fim</label><input class="form-input" type="date" id="valorDataFim"></div><div class="form-group"><label class="form-label">Valor Vascon</label><input class="form-input" type="number" step="0.01" id="valorVascon"></div><div class="form-group"><label class="form-label">Valor descontado funcionário</label><input class="form-input" type="number" step="0.01" id="valorDesconto"></div><div class="form-group"><label class="form-label">Ativo</label><select class="form-select" id="valorAtivo"><option value="true">Sim</option><option value="false">Não</option></select></div><div class="form-group" style="grid-column:1/-1"><label class="form-label">Observação</label><textarea class="form-textarea" id="valorObs"></textarea></div></div></div><div class="modal-footer"><button class="btn-secondary" type="button" onclick="fecharModalValorRefeicao()">Cancelar</button><button class="btn-success" type="button" onclick="salvarValorRefeicao()">Salvar valor</button></div></div></div>`);
  }
  window.abrirModalValorRefeicao = async function(id=''){
    ensureValueModal(); const modal=$('modalValorRefeicao'); modal.dataset.editId=id||'';
    ['valorTitulo','valorDataInicio','valorDataFim','valorVascon','valorDesconto','valorObs'].forEach(x=>{ if($(x)) $(x).value=''; }); if($('valorAtivo')) $('valorAtivo').value='true';
    if(id){ const item=(await SP.getValoresRefeicao(false)).find(v=>String(v.id)===String(id)); if(item){ $('valorTitulo').value=item.Title||''; $('valorDataInicio').value=String(item.Data_Inicio||'').slice(0,10); $('valorDataFim').value=String(item.Data_Fim||'').slice(0,10); $('valorVascon').value=item.Valor_Vascon||0; $('valorDesconto').value=item.Valor_Desconto_Funcionario||0; $('valorObs').value=item.Observacao||''; $('valorAtivo').value=SP.isTrue(item.Ativo)?'true':'false'; }}
    modal.classList.add('open');
  };
  window.fecharModalValorRefeicao = function(){ $('modalValorRefeicao')?.classList.remove('open'); };
  window.salvarValorRefeicao = async function(){
    const id=$('modalValorRefeicao')?.dataset?.editId||'';
    const dados={title:($('valorTitulo')?.value||'').trim(),dataInicio:$('valorDataInicio')?.value,dataFim:$('valorDataFim')?.value,valorVascon:Number($('valorVascon')?.value||0),valorDescontoFuncionario:Number($('valorDesconto')?.value||0),observacao:$('valorObs')?.value||'',ativo:$('valorAtivo')?.value==='true'};
    if(!dados.title){ toast2('Informe o título do período.','error'); return; }
    if(id) await SP.updateValorRefeicao(id,dados); else await SP.createValorRefeicao(dados);
    fecharModalValorRefeicao(); toast2('Valor salvo.','success'); await loadValoresRefeicao();
  };
  window.loadValoresRefeicao = async function(){
    ensureValueModal(); const tbody=$('valoresTable'); if(!tbody) return;
    try{ const vals=await SP.getValoresRefeicao(false); tbody.innerHTML = vals.map(v=>`<tr><td>${safe(v.Title||'')}</td><td>${safe(String(v.Data_Inicio||'').slice(0,10))}</td><td>${safe(String(v.Data_Fim||'').slice(0,10))}</td><td>${money(v.Valor_Vascon)}</td><td>${money(v.Valor_Desconto_Funcionario)}</td><td><span class="badge ${SP.isTrue(v.Ativo)?'badge-green':'badge-red'}">${SP.isTrue(v.Ativo)?'Ativo':'Inativo'}</span></td><td><div class="table-actions"><button class="btn-icon" onclick="abrirModalValorRefeicao('${v.id}')">✏️</button><button class="btn-icon danger" onclick="excluirValorRefeicao('${v.id}')">🗑️</button></div></td></tr>`).join('') || '<tr><td colspan="7" style="text-align:center;padding:2rem;color:rgba(143,170,210,0.4)">Nenhum valor cadastrado.</td></tr>'; }catch(e){ tbody.innerHTML=`<tr><td colspan="7" style="text-align:center;color:rgba(255,100,100,0.6);padding:2rem">Erro: ${safe(e.message)}</td></tr>`; }
  };
  window.excluirValorRefeicao = async function(id){ if(!confirm('Excluir este valor?')) return; await SP.deleteItem('Valores_Refeicao', id); toast2('Valor excluído.','success'); await loadValoresRefeicao(); };

  function ensureAusenciaModal(){
    if($('modalAusencia')) return;
    document.body.insertAdjacentHTML('beforeend',`<div class="modal-overlay" id="modalAusencia"><div class="modal-box"><div class="modal-header"><div class="modal-title">Ausência / bloqueio</div><button class="modal-close" type="button" onclick="fecharModalAusencia()">×</button></div><div class="modal-body"><div class="form-grid"><div class="form-group" style="grid-column:1/-1"><label class="form-label">Colaborador</label><select class="form-select" id="ausenciaColaborador"></select></div><div class="form-group"><label class="form-label">Data início</label><input class="form-input" type="date" id="ausenciaDataInicio"></div><div class="form-group"><label class="form-label">Data fim</label><input class="form-input" type="date" id="ausenciaDataFim"></div><div class="form-group"><label class="form-label">Motivo interno</label><select class="form-select" id="ausenciaMotivo"><option>Férias</option><option>Afastamento</option><option>Ausência</option><option>Viagem</option><option>Home Office</option><option>Não almoça na empresa</option><option>Outro</option></select></div><div class="form-group"><label class="form-label">Ativo</label><select class="form-select" id="ausenciaAtivo"><option value="true">Sim</option><option value="false">Não</option></select></div><div class="form-group" style="grid-column:1/-1"><label class="form-label">Observação</label><textarea class="form-textarea" id="ausenciaObs"></textarea></div></div></div><div class="modal-footer"><button class="btn-secondary" type="button" onclick="fecharModalAusencia()">Cancelar</button><button class="btn-success" type="button" onclick="salvarAusencia()">Salvar ausência</button></div></div></div>`);
  }
  async function popularColabsAusencia(){
    const sel=$('ausenciaColaborador'); if(!sel) return; const colabs=await SP.getColaboradores(); sel.innerHTML=colabs.map(c=>`<option value="${safe(c.id)}" data-nome="${safe(c.Nome||c.Title||'')}">${safe(c.Nome||c.Title||'')}</option>`).join('');
  }
  window.abrirModalAusencia = async function(id=''){
    ensureAusenciaModal(); await popularColabsAusencia(); const modal=$('modalAusencia'); modal.dataset.editId=id||''; ['ausenciaDataInicio','ausenciaDataFim','ausenciaObs'].forEach(x=>{if($(x))$(x).value='';}); if($('ausenciaMotivo'))$('ausenciaMotivo').value='Férias'; if($('ausenciaAtivo'))$('ausenciaAtivo').value='true';
    if(id){ const item=(await SP.getAusencias(false)).find(a=>String(a.id)===String(id)); if(item){ $('ausenciaColaborador').value=item.Colaborador_id||''; $('ausenciaDataInicio').value=String(item.Data_Inicio||'').slice(0,10); $('ausenciaDataFim').value=String(item.Data_Fim||'').slice(0,10); $('ausenciaMotivo').value=item.Motivo||'Outro'; $('ausenciaObs').value=item.Observacao||''; $('ausenciaAtivo').value=SP.isTrue(item.Ativo)?'true':'false'; }}
    modal.classList.add('open');
  };
  window.fecharModalAusencia = function(){ $('modalAusencia')?.classList.remove('open'); };
  window.salvarAusencia = async function(){
    const id=$('modalAusencia')?.dataset?.editId||''; const sel=$('ausenciaColaborador'); const nome=sel?.options[sel.selectedIndex]?.dataset?.nome || sel?.options[sel.selectedIndex]?.textContent || '';
    const dados={colaboradorId:sel?.value||'',colaboradorNome:nome,dataInicio:$('ausenciaDataInicio')?.value,dataFim:$('ausenciaDataFim')?.value,motivo:$('ausenciaMotivo')?.value,observacao:$('ausenciaObs')?.value||'',ativo:$('ausenciaAtivo')?.value==='true',criadoPor:currentUser()};
    if(!dados.colaboradorId || !dados.dataInicio || !dados.dataFim){ toast2('Informe colaborador, data inicial e data final.','error'); return; }
    if(id) await SP.updateAusencia(id,dados); else await SP.createAusencia(dados);
    fecharModalAusencia(); toast2('Ausência salva.','success'); await loadAusencias();
  };
  window.loadAusencias = async function(){
    ensureAusenciaModal(); const tbody=$('ausenciasTable'); if(!tbody) return;
    try{ const aus=await SP.getAusencias(false); tbody.innerHTML=aus.map(a=>`<tr><td>${safe(a.Colaborador_nome||'')}</td><td>Período de ${safe(String(a.Data_Inicio||'').slice(0,10).split('-').reverse().join('/'))} a ${safe(String(a.Data_Fim||'').slice(0,10).split('-').reverse().join('/'))}</td><td>${safe(a.Motivo||'')}</td><td><span class="badge ${SP.isTrue(a.Ativo)?'badge-green':'badge-red'}">${SP.isTrue(a.Ativo)?'Ativo':'Inativo'}</span></td><td>${safe(a.Observacao||'—')}</td><td><div class="table-actions"><button class="btn-icon" onclick="abrirModalAusencia('${a.id}')">✏️</button><button class="btn-icon danger" onclick="excluirAusencia('${a.id}')">🗑️</button></div></td></tr>`).join('') || '<tr><td colspan="6" style="text-align:center;padding:2rem;color:rgba(143,170,210,0.4)">Nenhuma ausência cadastrada.</td></tr>'; }catch(e){ tbody.innerHTML=`<tr><td colspan="6" style="text-align:center;color:rgba(255,100,100,0.6);padding:2rem">Erro: ${safe(e.message)}</td></tr>`; }
  };
  window.excluirAusencia = async function(id){ if(!confirm('Excluir esta ausência?')) return; await SP.deleteAusencia(id); toast2('Ausência excluída.','success'); await loadAusencias(); };

  document.addEventListener('DOMContentLoaded',()=>{
    setTimeout(()=>{
      patchModulesAndNav(); patchColaboradorUI(); patchOperacaoCards();
      if($('mod-relatorios')) installReportUI();
      ['filtroDia','filtroOpcao'].forEach(id=>$(id)?.addEventListener('change',()=>window.loadPedidos(semanaAtual())));
    },800);
  });
})();
