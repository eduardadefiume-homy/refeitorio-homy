// admin-dashboard.js — Dashboard do Admin Homy
// Inclui: painel de prazo, travamento manual e automático de pendentes

const AdminDashboard = window.AdminDashboard = {

  _prazoTimer: null,   // timer do auto-travamento

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
      await this._carregarPrazo(semanaId, resumo);
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
  },

  // ── Prazo + Painel de Travamento ──────────────────────────────
  async _carregarPrazo(semanaId, resumo) {
    const prazo = await SP.getPrazoMarcacao().catch(() => null);

    // Preenche campos de prazo
    if (prazo) {
      const dt = new Date(prazo);
      AdminUtils.setVal("prazoData", dt.toISOString().slice(0, 10));
      AdminUtils.setVal("prazoHora", dt.toTimeString().slice(0, 5));
    }

    // Atualiza painel de status do prazo
    this._atualizarPainelPrazo(prazo, resumo?.pendentesColaboradores ?? 0);

    // Inicia timer de auto-travamento
    this._iniciarTimerPrazo(prazo, semanaId);

    // Bind botões (uma vez)
    const btnSalvar = document.getElementById("btnSalvarPrazo");
    if (btnSalvar && !btnSalvar.dataset.bound) {
      btnSalvar.dataset.bound = "1";
      btnSalvar.addEventListener("click", async () => {
        const data = AdminUtils.getVal("prazoData");
        const hora = AdminUtils.getVal("prazoHora") || "18:00";
        if (!data) { AdminUtils.toast("Informe a data limite.", "error"); return; }
        const valor = `${data}T${hora}:00`;
        await SP.setPrazoMarcacao(valor);
        AdminUtils.toast("✅ Prazo salvo.", "success");
        const r = await SP.getDashboardResumo(AdminState.getSemanaId()).catch(() => ({}));
        this._atualizarPainelPrazo(valor, r.pendentesColaboradores ?? 0);
        this._iniciarTimerPrazo(valor, AdminState.getSemanaId());
      });
    }

    const btnTravar = document.getElementById("btnTravarPendentes");
    if (btnTravar && !btnTravar.dataset.bound) {
      btnTravar.dataset.bound = "1";
      btnTravar.addEventListener("click", () => this._confirmarTravamento(semanaId));
    }
  },

  _atualizarPainelPrazo(prazoISO, pendentes) {
    const el = document.getElementById("dashPainelPrazo");
    if (!el) return;

    if (!prazoISO) {
      el.innerHTML = `<div class="alert alert-warning">⚠️ Nenhum prazo definido para esta semana.</div>`;
      return;
    }

    const dt    = new Date(prazoISO);
    const agora = new Date();
    const vencido = agora > dt;

    const fmtDt = dt.toLocaleDateString("pt-BR", {
      weekday: "long", day: "2-digit", month: "2-digit"
    }) + " às " + dt.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });

    const alertClass = vencido ? "alert-red" : "alert-info";
    const icone      = vencido ? "🔒" : "⏱️";
    const statusTxt  = vencido
      ? `Prazo encerrado. ${pendentes > 0 ? pendentes + " colaboradores não marcaram." : "Todos marcaram."}`
      : `Prazo: ${fmtDt}. ${pendentes} colaboradores ainda não marcaram.`;

    const btnVisivel = vencido && pendentes > 0;

    el.innerHTML = `
      <div class="alert ${alertClass}" style="display:flex;align-items:center;justify-content:space-between;gap:1rem;flex-wrap:wrap">
        <span>${icone} ${AdminUtils.esc(statusTxt)}</span>
        ${btnVisivel ? `<button id="btnTravarPendentes" class="btn-primary" style="flex-shrink:0;font-size:.82rem;padding:.5rem 1rem">
          🔒 Travar ${pendentes} pendente(s) como Principal
        </button>` : ""}
      </div>
    `;

    // Re-bind após injetar HTML
    const btnTravar = document.getElementById("btnTravarPendentes");
    if (btnTravar && !btnTravar.dataset.bound) {
      btnTravar.dataset.bound = "1";
      btnTravar.addEventListener("click", () =>
        this._confirmarTravamento(AdminState.getSemanaId(), pendentes)
      );
    }
  },

  // ── Confirmação e execução do travamento ──────────────────────
  async _confirmarTravamento(semanaId, qtd) {
    const resumo = await SP.getDashboardResumo(semanaId).catch(() => ({}));
    const total  = qtd ?? resumo.pendentesColaboradores ?? 0;

    if (total === 0) {
      AdminUtils.toast("✅ Todos os colaboradores já marcaram.", "success");
      return;
    }

    const confirmar = window.confirm(
      `Confirmar travamento?\n\n` +
      `${total} colaborador(es) não marcaram no prazo.\n` +
      `Todos serão marcados como PRINCIPAL em todos os dias desta semana.\n\n` +
      `A observação "Marcado automaticamente — prazo encerrado" será registrada para auditoria.\n\n` +
      `Esta ação não pode ser desfeita.`
    );
    if (!confirmar) return;

    const btn = document.getElementById("btnTravarPendentes");
    if (btn) { btn.disabled = true; btn.textContent = "⏳ Travando..."; }

    try {
      const resultado = await SP.travarPendentesComoPrincipal(semanaId);
      AdminUtils.toast(
        `✅ ${resultado.travados} colaboradores travados como Principal.`,
        "success"
      );
      // Bloqueia marcação automaticamente
      await SP.setMarcacaoLiberada(false);
      // Recarrega dashboard
      await AdminDashboard.load(semanaId);
    } catch (e) {
      AdminUtils.toast("Erro ao travar: " + (e.message || e), "error");
      if (btn) { btn.disabled = false; btn.textContent = "🔒 Travar pendentes como Principal"; }
    }
  },

  // ── Timer de auto-travamento ──────────────────────────────────
  _iniciarTimerPrazo(prazoISO, semanaId) {
    // Cancela timer anterior
    if (this._prazoTimer) { clearTimeout(this._prazoTimer); this._prazoTimer = null; }
    if (!prazoISO) return;

    const dt    = new Date(prazoISO);
    const agora = new Date();
    const ms    = dt - agora;

    if (ms <= 0) {
      // Prazo já venceu — verifica se precisa travar (silencioso)
      this._autoTravarSeNecessario(semanaId);
      return;
    }

    // Agenda travamento automático
    this._prazoTimer = setTimeout(async () => {
      console.log("[Dashboard] Prazo atingido — iniciando auto-travamento");
      AdminUtils.toast("⏰ Prazo atingido! Travando pendentes automaticamente...", "info");
      await this._autoTravarSeNecessario(semanaId);
    }, ms);

    console.log(`[Dashboard] Auto-travamento agendado em ${Math.round(ms/60000)} min`);
  },

  async _autoTravarSeNecessario(semanaId) {
    try {
      const resultado = await SP.travarPendentesComoPrincipal(semanaId);
      if (resultado.travados > 0) {
        await SP.setMarcacaoLiberada(false);
        AdminUtils.toast(
          `🔒 ${resultado.travados} colaboradores travados automaticamente como Principal.`,
          "success"
        );
        await AdminDashboard.load(semanaId);
      }
    } catch (e) {
      console.error("[Dashboard] auto-travamento:", e);
    }
  },

  // ── Demais renders (inalterados) ──────────────────────────────
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
        <td>${AdminUtils.esc(SP.pick(e,"Nome","Title"))}</td>
        <td>${AdminUtils.esc(SP.pick(e,"tipo","Tipo")||"Extra")}</td>
        <td>${AdminUtils.esc(SP.pick(e,"Opcao")||"—")}</td>
        <td>${AdminUtils.badge(SP.pick(e,"Status")||"Confirmado")}</td>
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
      norm(SP.pick(p,"Dia")) === norm(r.diaHoje) &&
      ["cancelado","afastado","ferias","nao vai almocar","bloqueado","travado"]
        .includes(norm(SP.pick(p,"Status")||""))
    );
    if (!ausentes.length) {
      tbody.innerHTML = `<tr><td colspan="3" class="empty-cell">Nenhuma ausência hoje.</td></tr>`;
      return;
    }
    tbody.innerHTML = ausentes.slice(0, 14).map(p => `
      <tr>
        <td>${AdminUtils.esc(SP.pick(p,"Colaborador_nome")||"—")}</td>
        <td>${AdminUtils.esc(SP.pick(p,"Centro_Custo")||"—")}</td>
        <td>${AdminUtils.badge(SP.pick(p,"Status")||"Pendente")}</td>
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
    const lista = pedidos.filter(p => norm(SP.pick(p,"Dia")) === norm(dia));
    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhum pedido para este dia.</td></tr>`;
      return;
    }
    tbody.innerHTML = lista.slice(0, 60).map(p => {
      const id    = AdminUtils.esc(p.id||"");
      const nome  = AdminUtils.esc(SP.pick(p,"Colaborador_nome")||"—");
      const opcao = AdminUtils.esc(SP.pick(p,"Opcao")||"—");
      const prato = AdminUtils.esc(SP.pick(p,"Nome_Prato")||"—");
      const status = SP.pick(p,"Status")||"Pendente";
      return `<tr>
        <td>${nome}</td>
        <td>${AdminUtils.esc(SP.pick(p,"Centro_Custo")||"—")}</td>
        <td><span class="badge badge-blue">${opcao}</span></td>
        <td>${prato}</td>
        <td>${AdminUtils.badge(status)}</td>
        <td><div class="table-actions">
          <button class="btn-icon" onclick="AdminOperacao.alterarStatus('${id}','Confirmado')">✅</button>
          <button class="btn-icon danger" onclick="AdminOperacao.alterarStatus('${id}','Cancelado')">❌</button>
          <button class="btn-icon" onclick="AdminOperacao.alterarStatus('${id}','Não vai almoçar')">🚫</button>
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
    ].map(([a,b]) => `
      <div class="dashboard-list-item">
        <div><div class="dashboard-list-main">${a}</div><div class="dashboard-list-sub">${b}</div></div>
      </div>`).join("");
  },

  _renderSetores(r) {
    const el = document.getElementById("dashSetoresList");
    if (!el) return;
    const arr = r.setoresHoje || [];
    el.innerHTML = arr.length
      ? arr.slice(0,6).map(([s,t]) => `
          <div class="dashboard-list-item">
            <div><div class="dashboard-list-main">${AdminUtils.esc(s)}</div>
            <div class="dashboard-list-sub">${t} refeições hoje</div></div>
            <span class="badge badge-blue">${t}</span>
          </div>`).join("")
      : `<div class="dashboard-list-item"><div class="dashboard-list-main">Sem dados hoje</div></div>`;
  },

  _renderProximosDias(r) {
    const el = document.getElementById("dashProximosDias");
    if (!el) return;
    el.innerHTML = AdminUtils.DIAS.map(dia => {
      const d = r.porDia?.[dia] || {};
      return `<div class="dashboard-list-item">
        <div><div class="dashboard-list-main">${AdminUtils.DIA_LABEL[dia]}</div>
        <div class="dashboard-list-sub">${d.total??0} conf. · ${d.pendentes??0} pend.</div></div>
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
      `<div class="dashboard-alert-item ${x.startsWith("✅")?"ok":""}">${x}</div>`
    ).join("");
    AdminUtils.setTxt("dashAlteracoesManuais","—");
  }
};
