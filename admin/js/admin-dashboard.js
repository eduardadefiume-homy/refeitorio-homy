// admin-dashboard.js — Dashboard do Admin Homy

const AdminDashboard = window.AdminDashboard = {

  async load(semanaId) {
    try {
      await SP.init();
      const resumo = await SP.getDashboardResumo(semanaId);
      this._renderCards(resumo);
      this._renderToggleStatus(resumo);
      this._renderExtrasHoje(resumo, semanaId);
      this._renderSemanaTable(resumo);
      this._renderAusencias(resumo, semanaId);
      this._renderOperacaoDia(resumo, semanaId);
      this._renderGerencial(resumo);
      this._renderSetores(resumo);
      this._renderProximosDias(resumo);
      this._renderAlertas(resumo);
      this._carregarPrazo();
      AdminUtils.setTxt("semanaLabel", AdminState.getSemanaLabel());
    } catch (e) {
      console.error("[Dashboard]", e);
      AdminUtils.toast("Erro ao carregar dashboard: " + e.message, "error");
    }
  },

  _renderCards(r) {
    AdminUtils.setTxt("stat-colab",        r.colaboradoresAtivos);
    AdminUtils.setTxt("stat-confirmados",  r.pedidosConfirmadosColaboradores);
    AdminUtils.setTxt("stat-pendentes",    r.pendentesColaboradores);
    AdminUtils.setTxt("stat-checkin",      r.checkinsHoje);
    AdminUtils.setTxt("dashExtrasAtivos",  r.extrasAtivos);
    AdminUtils.setTxt("dashTotalHoje",     r.totalPedidosHoje);
    AdminUtils.setTxt("dashAusenciasHoje", r.ausenciasHoje);
    AdminUtils.setTxt("dashTotalSemana",   r.totalPedidosSemana);
    AdminUtils.setTxt("dashPrincipalHoje", r.principalHoje);
    AdminUtils.setTxt("dashLightHoje",     r.lightHoje);
    AdminUtils.setTxt("dashOutrasHoje",    r.outrasHoje);
    AdminUtils.setTxt("dashSetoresHoje",   r.setoresHoje?.length ?? "—");
  },

  async _renderToggleStatus() {
    const liberado  = await SP.isCardapioLiberado().catch(() => false);
    const cardapio  = await SP.getConfig("cardapio_visivel").catch(() => null);
    const cardapioV = SP.isTrue(cardapio);

    const tMarcacao = document.getElementById("toggleMarcacao");
    const tCardapio = document.getElementById("toggleCardapio");
    if (tMarcacao && !tMarcacao.dataset.bound) {
      tMarcacao.checked = liberado;
      tMarcacao.dataset.bound = "1";
      tMarcacao.addEventListener("change", async function () {
        await SP.setMarcacaoLiberada(this.checked);
        AdminUtils.toast(this.checked ? "Marcação liberada." : "Marcação bloqueada.", "success");
        AdminUtils.setTxt("dashMarcacaoStatus", this.checked ? "Aberta" : "Fechada");
      });
    }
    if (tCardapio && !tCardapio.dataset.bound) {
      tCardapio.checked = cardapioV;
      tCardapio.dataset.bound = "1";
      tCardapio.addEventListener("change", async function () {
        await SP.setCardapioVisivel(this.checked);
        AdminUtils.toast(this.checked ? "Cardápio visível." : "Cardápio ocultado.", "success");
        AdminUtils.setTxt("dashCardapioStatus", this.checked ? "Liberado" : "Bloqueado");
      });
    }
    AdminUtils.setTxt("dashMarcacaoStatus", liberado  ? "Aberta"   : "Fechada");
    AdminUtils.setTxt("dashCardapioStatus", cardapioV ? "Liberado" : "Bloqueado");

    // Botão salvar prazo
    const btnPrazo = document.getElementById("btnSalvarPrazo");
    if (btnPrazo && !btnPrazo.dataset.bound) {
      btnPrazo.dataset.bound = "1";
      btnPrazo.addEventListener("click", async () => {
        const data = AdminUtils.getVal("prazoData");
        const hora = AdminUtils.getVal("prazoHora") || "18:00";
        if (!data) { AdminUtils.toast("Informe a data limite.", "error"); return; }
        await SP.setPrazoMarcacao(`${data}T${hora}:00`);
        AdminUtils.toast("Prazo salvo.", "success");
      });
    }
  },

  async _carregarPrazo() {
    const prazo = await SP.getPrazoMarcacao().catch(() => null);
    if (!prazo) return;
    const dt = new Date(prazo);
    AdminUtils.setVal("prazoData", dt.toISOString().slice(0, 10));
    AdminUtils.setVal("prazoHora", dt.toTimeString().slice(0, 5));
  },

  async _renderExtrasHoje(r, semanaId) {
    AdminUtils.setTxt("dashExtrasSolicitados", r.extrasAtivos);
    AdminUtils.setTxt("dashExtrasConfirmados", r.extrasConfirmados);
    AdminUtils.setTxt("dashExtrasPendentes",   r.extrasPendentes);

    const tbody = document.getElementById("dashExtrasTable");
    if (!tbody) return;
    const extras = await SP.getExtras(semanaId, r.diaHoje).catch(() => []);
    if (!extras.length) {
      tbody.innerHTML = `<tr><td colspan="4" class="empty-cell">Nenhum extra para hoje.</td></tr>`;
      return;
    }
    tbody.innerHTML = extras.slice(0, 12).map(e => `
      <tr>
        <td>${AdminUtils.esc(SP.pick(e, "Nome", "Title"))}</td>
        <td>${AdminUtils.esc(SP.pick(e, "tipo", "Tipo") || "Extra")}</td>
        <td>${AdminUtils.esc(SP.pick(e, "Opcao") || "—")}</td>
        <td>${AdminUtils.badge(SP.pick(e, "Status") || "Confirmado")}</td>
      </tr>`).join("");
  },

  _renderSemanaTable(r) {
    const tbody = document.getElementById("dashSemanaTable");
    if (!tbody) return;
    tbody.innerHTML = AdminUtils.DIAS.map(dia => {
      const d = r.porDia?.[dia] || {};
      return `<tr>
        <td>${AdminUtils.DIA_LABEL[dia]}</td>
        <td>${d.total     ?? 0}</td>
        <td>${d.principal ?? 0}</td>
        <td>${d.light     ?? 0}</td>
        <td>${d.pendentes ?? 0}</td>
      </tr>`;
    }).join("");
  },

  async _renderAusencias(r, semanaId) {
    AdminUtils.setTxt("dashNaoMarcaram", r.pendentesColaboradores);
    AdminUtils.setTxt("dashCancelados",  r.ausenciasHoje);
    AdminUtils.setTxt("dashTravados",    "—");

    const tbody = document.getElementById("dashAusenciasTable");
    if (!tbody) return;

    const pedidos = await SP.getPedidos(semanaId).catch(() => []);
    const norm    = v => AdminUtils.norm(v);
    const ausentes = pedidos.filter(p =>
      norm(SP.pick(p, "Dia")) === norm(r.diaHoje) &&
      ["cancelado", "afastado", "ferias", "nao vai almocar", "bloqueado", "travado"]
        .includes(norm(SP.pick(p, "Status") || ""))
    );

    if (!ausentes.length) {
      tbody.innerHTML = `<tr><td colspan="3" class="empty-cell">Nenhuma ausência hoje.</td></tr>`;
      return;
    }
    tbody.innerHTML = ausentes.slice(0, 14).map(p => `
      <tr>
        <td>${AdminUtils.esc(SP.pick(p, "Colaborador_nome") || "—")}</td>
        <td>${AdminUtils.esc(SP.pick(p, "Centro_Custo")     || "—")}</td>
        <td>${AdminUtils.badge(SP.pick(p, "Status") || "Pendente")}</td>
      </tr>`).join("");
  },

  async _renderOperacaoDia(r, semanaId) {
    const sel = document.getElementById("dashOperacaoDia");
    if (sel && !sel.dataset.bound) {
      sel.value = r.diaHoje;
      sel.dataset.bound = "1";
      sel.addEventListener("change", () => this._carregarOperacaoTabela(semanaId, sel.value));
    }
    const btn = document.getElementById("btnAtualizarDashboard");
    if (btn && !btn.dataset.bound) {
      btn.dataset.bound = "1";
      btn.addEventListener("click", () => AdminDashboard.load(AdminState.getSemanaId()));
    }
    await this._carregarOperacaoTabela(semanaId, r.diaHoje);
  },

  async _carregarOperacaoTabela(semanaId, dia) {
    const tbody = document.getElementById("dashOperacaoTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    const pedidos = await SP.getPedidos(semanaId).catch(() => []);
    const norm = v => AdminUtils.norm(v);
    const lista = pedidos.filter(p => norm(SP.pick(p, "Dia")) === norm(dia));

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhum pedido para este dia.</td></tr>`;
      return;
    }
    tbody.innerHTML = lista.slice(0, 60).map(p => {
      const id    = AdminUtils.esc(p.id || "");
      const nome  = AdminUtils.esc(SP.pick(p, "Colaborador_nome") || "—");
      const opcao = AdminUtils.esc(SP.pick(p, "Opcao") || "—");
      const prato = AdminUtils.esc(SP.pick(p, "Nome_Prato") || "—");
      const status = SP.pick(p, "Status") || "Pendente";
      const diaP  = AdminUtils.esc(SP.pick(p, "Dia") || dia);
      return `<tr>
        <td>${nome}</td>
        <td>${AdminUtils.esc(SP.pick(p, "Centro_Custo") || "—")}</td>
        <td><span class="badge badge-blue">${opcao}</span></td>
        <td>${prato}</td>
        <td>${AdminUtils.badge(status)}</td>
        <td><div class="table-actions">
          <button class="btn-icon" title="Confirmar" onclick="AdminOperacao.alterarStatus('${id}','Confirmado')">✅</button>
          <button class="btn-icon danger" title="Cancelar" onclick="AdminOperacao.alterarStatus('${id}','Cancelado')">❌</button>
          <button class="btn-icon" title="Não vai" onclick="AdminOperacao.alterarStatus('${id}','Não vai almoçar')">🚫</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  _renderGerencial(r) {
    const el = document.getElementById("dashGerencialList");
    if (!el) return;
    el.innerHTML = [
      ["Consumo da semana",    `${r.totalPedidosSemana} refeições`],
      ["Pendências da semana", `${r.pendentesColaboradores} registros`],
      ["Extras confirmados",   `${r.extrasConfirmados} extras`],
      ["Ausências hoje",       `${r.ausenciasHoje} registros`]
    ].map(([a, b]) => `
      <div class="dashboard-list-item">
        <div>
          <div class="dashboard-list-main">${a}</div>
          <div class="dashboard-list-sub">${b}</div>
        </div>
      </div>`).join("");
  },

  _renderSetores(r) {
    const el = document.getElementById("dashSetoresList");
    if (!el) return;
    const arr = r.setoresHoje || [];
    el.innerHTML = arr.length
      ? arr.slice(0, 6).map(([setor, total]) => `
          <div class="dashboard-list-item">
            <div>
              <div class="dashboard-list-main">${AdminUtils.esc(setor)}</div>
              <div class="dashboard-list-sub">${total} refeições hoje</div>
            </div>
            <span class="badge badge-blue">${total}</span>
          </div>`).join("")
      : `<div class="dashboard-list-item"><div class="dashboard-list-main">Sem dados hoje</div></div>`;
  },

  _renderProximosDias(r) {
    const el = document.getElementById("dashProximosDias");
    if (!el) return;
    el.innerHTML = AdminUtils.DIAS.map(dia => {
      const d = r.porDia?.[dia] || {};
      return `<div class="dashboard-list-item">
        <div>
          <div class="dashboard-list-main">${AdminUtils.DIA_LABEL[dia]}</div>
          <div class="dashboard-list-sub">${d.total ?? 0} conf. · ${d.pendentes ?? 0} pend.</div>
        </div>
      </div>`;
    }).join("");
  },

  _renderAlertas(r) {
    const el = document.getElementById("dashAlertas");
    if (!el) return;
    const itens = [];
    if (r.pendentesColaboradores > 0)
      itens.push(`⚠️ ${r.pendentesColaboradores} colaboradores ainda não marcaram.`);
    if (r.extrasPendentes > 0)
      itens.push(`⚠️ ${r.extrasPendentes} extras aguardando confirmação.`);
    if (!itens.length)
      itens.push(`✅ Operação sem alertas críticos no momento.`);

    el.innerHTML = itens.map(x =>
      `<div class="dashboard-alert-item ${x.startsWith("✅") ? "ok" : ""}">${x}</div>`
    ).join("");
    AdminUtils.setTxt("dashAlteracoesManuais", "—");
  }
};
