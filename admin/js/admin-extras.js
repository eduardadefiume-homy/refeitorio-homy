// admin-extras.js — Extras / Visitantes do Admin Homy
// Correção: Guarda/Extras entram no centro de custo da Portaria e criam pedido correspondente.

const AdminExtras = window.AdminExtras = {
  _lista: [],
  PORTARIA_CC: "120602 - PORTARIA",

  async load(semanaId) {
    this._bindFiltros(semanaId);
    this._bindBotoes(semanaId);
    await this._carregar(semanaId);
  },

  _norm(v) {
    return String(v || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim();
  },

  _pick(obj, ...keys) {
    for (const k of keys) {
      const v = SP.pick ? SP.pick(obj, k) : obj?.[k];
      if (v !== undefined && v !== null && String(v) !== "") return v;
    }
    return "";
  },

  _esc(v) {
    return AdminUtils.esc ? AdminUtils.esc(v) : String(v ?? "");
  },

  _tipoPortaria(tipo, nome = "") {
    const t = this._norm(`${tipo} ${nome}`);
    return t.includes("guarda") || t.includes("extra") || t.includes("visitante") || t.includes("motorista") || t.includes("prestador") || t.includes("marmita");
  },

  _centroCustoPorTipo(tipo, nome = "") {
    // Refeições extras, Guarda e visitantes entram no rateio da Portaria.
    return this._tipoPortaria(tipo, nome) ? this.PORTARIA_CC : this.PORTARIA_CC;
  },

  async _carregar(semanaId) {
    const tbody = document.getElementById("extrasTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const todos = typeof SP.getExtras === "function" ? await SP.getExtras(semanaId) : await SP.getItems("Extras");

      const seenAuto = new Set();
      this._lista = (todos || []).filter(e => {
        const nome = this._norm(this._pick(e, "Nome", "Title") || "");
        const tipo = this._norm(this._pick(e, "tipo", "Tipo") || "");
        const dia = this._norm(this._pick(e, "Dia") || "");
        const isAuto = nome.includes("refeicaoextra") || (nome.includes("extra") && tipo.includes("extra"));
        if (isAuto) {
          const k = `auto-${dia}`;
          if (seenAuto.has(k)) return false;
          seenAuto.add(k);
        }
        return true;
      });

      const diaOrd = { segunda: 1, terca: 2, quarta: 3, quinta: 4, sexta: 5 };
      this._lista.sort((a, b) => {
        const da = diaOrd[this._norm(this._pick(a, "Dia"))] || 9;
        const db = diaOrd[this._norm(this._pick(b, "Dia"))] || 9;
        if (da !== db) return da - db;
        return this._norm(this._pick(a, "tipo", "Tipo")).localeCompare(this._norm(this._pick(b, "tipo", "Tipo")), "pt-BR");
      });

      this._render();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell" style="color:#ff8080">Erro: ${this._esc(e.message || e)}</td></tr>`;
    }
  },

  _render() {
    const tbody = document.getElementById("extrasTable");
    if (!tbody) return;

    const txt = this._norm(AdminUtils.getVal("fExtraTexto"));
    const dia = this._norm(AdminUtils.getVal("fExtraDia"));
    const tipo = this._norm(AdminUtils.getVal("fExtraTipo"));

    const lista = this._lista.filter(e => {
      const all = this._norm([
        this._pick(e, "Nome", "Title"),
        this._pick(e, "Observacao"),
        this._pick(e, "tipo", "Tipo"),
        this._pick(e, "Dia"),
        this._pick(e, "Centro_Custo", "Setor")
      ].join(" "));
      return (!txt || all.includes(txt)) &&
             (!dia || this._norm(this._pick(e, "Dia")) === dia) &&
             (!tipo || this._norm(this._pick(e, "tipo", "Tipo")).includes(tipo));
    });

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhum extra encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(e => {
      const id = this._esc(e.id || "");
      const nome = this._esc(this._pick(e, "Nome", "Title") || "Extra");
      const tipo = this._esc(this._pick(e, "tipo", "Tipo") || "—");
      const dia = this._esc(this._pick(e, "Dia") || "—");
      const opc = this._esc(this._pick(e, "Opcao") || "principal");
      const obs = this._esc(this._pick(e, "Observacao") || "—");
      return `<tr>
        <td>${nome}</td>
        <td><span class="badge badge-yellow">${tipo}</span></td>
        <td>${dia}</td>
        <td><span class="badge badge-blue">${opc}</span></td>
        <td style="font-size:.78rem;color:rgba(143,170,210,.7)">${obs}</td>
        <td><div class="table-actions">
          <button class="btn-icon danger" title="Excluir" onclick="AdminExtras.excluir('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  abrirModal(predefinido = "") {
    const modal = document.getElementById("modalExtra");
    if (!modal) return;
    modal.dataset.predefinido = predefinido;

    ["extraNome", "extraTipo", "extraOpcao", "extraObs"].forEach(id => {
      const el = document.getElementById(id);
      if (el?.closest(".form-group")) el.closest(".form-group").style.display = predefinido ? "none" : "";
      if (!predefinido && el) el.value = "";
    });

    if (predefinido === "investigador") {
      AdminUtils.setVal("extraNome", "Investigador");
      AdminUtils.setVal("extraTipo", "investigador");
      AdminUtils.setVal("extraOpcao", "principal");
      AdminUtils.setVal("extraObs", "Investigador");
    } else if (predefinido === "guarda") {
      AdminUtils.setVal("extraNome", "Guarda");
      AdminUtils.setVal("extraTipo", "guarda");
      AdminUtils.setVal("extraOpcao", "principal");
      AdminUtils.setVal("extraObs", `Guarda — Centro de custo ${this.PORTARIA_CC}`);
    }

    AdminUtils.openModal("modalExtra");
  },

  async _pratoPorOpcao(semanaId, dia, opcao) {
    try {
      if (typeof SP.getCardapio !== "function") return "Cardápio do Dia";
      const cardapio = await SP.getCardapio(semanaId);
      const item = (cardapio || []).find(c =>
        this._norm(this._pick(c, "Dia")) === this._norm(dia) &&
        this._norm(this._pick(c, "Opcao", "Opção")) === this._norm(opcao)
      );
      return this._pick(item, "Nome_Prato", "Descricao", "Descrição", "Title") || "Cardápio do Dia";
    } catch (_) {
      return "Cardápio do Dia";
    }
  },

  async _contarExistentes(semanaId, dia, tipo) {
    const pedidos = await SP.getPedidos(semanaId).catch(() => []);
    return pedidos.filter(p =>
      this._norm(this._pick(p, "Dia")) === this._norm(dia) &&
      this._norm(this._pick(p, "Origem", "tipo", "Tipo", "Colaborador_nome")).includes(this._norm(tipo))
    ).length;
  },

  async _criarExtra(semanaId, dia, nome, tipo, opcao, obs) {
    if (typeof SP.addExtra === "function") {
      return await SP.addExtra(semanaId, dia, nome, tipo, opcao, obs, SP.getUserName ? SP.getUserName() : "Admin");
    }
    return await SP.createItem("Extras", {
      Title: `${semanaId}-${dia}-${nome}`,
      Semana_id: semanaId,
      Dia: dia,
      Nome: nome,
      tipo: tipo,
      Opcao: opcao || "principal",
      Observacao: obs || "",
      Adicionado_Por: SP.getUserName ? SP.getUserName() : "Admin"
    });
  },

  async _criarPedidoDoExtra(semanaId, dia, nome, tipo, opcao, obs) {
    const centroCusto = this._centroCustoPorTipo(tipo, nome);
    const nomePrato = await this._pratoPorOpcao(semanaId, dia, opcao);
    const colabId = `extra-${this._norm(tipo)}-${Date.now()}`;

    if (typeof SP.savePedido === "function") {
      return await SP.savePedido(semanaId, colabId, nome, dia, opcao || "principal", nomePrato, {
        confirmado: true,
        status: "Confirmado",
        centroCusto,
        origem: tipo || "Extra",
        observacao: obs || "",
        alteradoPor: SP.getUserName ? SP.getUserName() : "Admin"
      });
    }

    return await SP.createItem("Pedidos", {
      Title: `${semanaId}-${colabId}-${dia}`,
      Semana_id: semanaId,
      Colaborador_id: colabId,
      Colaborador_nome: nome,
      Dia: dia,
      Opcao: opcao || "principal",
      Nome_Prato: nomePrato,
      Confirmado: true,
      Data_Hora: new Date().toISOString(),
      Centro_Custo: centroCusto,
      Status: "Confirmado",
      Observacao: obs || "",
      Origem: tipo || "Extra",
      Alterado_Por: SP.getUserName ? SP.getUserName() : "Admin"
    });
  },

  async salvar() {
    const modal = document.getElementById("modalExtra");
    const predefinido = modal?.dataset.predefinido || "";
    const semanaId = AdminState.getSemanaId();
    const dia = AdminUtils.getVal("extraDia") || "segunda";

    let nome, tipo, opcao, obs;

    if (predefinido === "investigador") {
      const qtd = await this._contarExistentes(semanaId, dia, "investigador");
      nome = `Investigador ${qtd + 1}`;
      tipo = "investigador";
      opcao = "principal";
      obs = "Investigador";
    } else if (predefinido === "guarda") {
      const qtd = await this._contarExistentes(semanaId, dia, "guarda");
      nome = qtd > 0 ? `Guarda ${qtd + 1}` : "Guarda";
      tipo = "guarda";
      opcao = "principal";
      obs = `Guarda — Centro de custo ${this.PORTARIA_CC}`;
    } else {
      nome = AdminUtils.getVal("extraNome");
      tipo = AdminUtils.getVal("extraTipo") || "visitante";
      opcao = AdminUtils.getVal("extraOpcao") || "principal";
      obs = AdminUtils.getVal("extraObs");
      if (!obs && this._tipoPortaria(tipo, nome)) obs = `Centro de custo ${this.PORTARIA_CC}`;
    }

    if (!nome) { AdminUtils.toast("Informe o nome do extra.", "error"); return; }

    const btn = document.getElementById("salvarExtra");
    const old = btn?.textContent;
    if (btn) { btn.disabled = true; btn.textContent = "⏳ Salvando..."; }

    try {
      await SP.init();
      await this._criarExtra(semanaId, dia, nome, tipo, opcao, obs);
      await this._criarPedidoDoExtra(semanaId, dia, nome, tipo, opcao, obs);
      AdminUtils.closeModal("modalExtra");
      AdminUtils.toast(`Extra adicionado em ${this.PORTARIA_CC}.`, "success");
      await this._carregar(semanaId);
      if (window.AdminDashboard && AdminState.moduloAtivo === "dashboard") await AdminDashboard.load(semanaId);
    } catch (e) {
      console.error("[Extras] salvar", e);
      AdminUtils.toast("Erro ao salvar extra: " + (e.message || e), "error");
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = old || "💾 Adicionar"; }
    }
  },

  async excluir(id) {
    if (!confirm("Excluir este extra?")) return;
    try {
      await SP.init();
      const semanaId = AdminState.getSemanaId();
      const item = this._lista.find(e => String(e.id) === String(id));
      if (item && typeof SP.deleteExtraComPedido === "function") await SP.deleteExtraComPedido(item);
      else if (typeof SP.deleteExtra === "function") await SP.deleteExtra(id);
      else await SP.deleteItem("Extras", id);
      this._lista = this._lista.filter(e => String(e.id) !== String(id));
      this._render();
      AdminUtils.toast("Extra excluído.", "success");
    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + (e.message || e), "error");
    }
  },

  _bindFiltros(semanaId) {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundExt) { el.dataset.boundExt = "1"; el.addEventListener(ev, fn); }
    };
    bind("fExtraTexto", "input", () => this._render());
    bind("fExtraDia", "change", () => this._render());
    bind("fExtraTipo", "change", () => this._render());
    bind("btnFiltrarExtras", "click", () => this._render());
    bind("btnLimparExtras", "click", () => {
      ["fExtraTexto", "fExtraDia", "fExtraTipo"].forEach(id => AdminUtils.setVal(id, ""));
      this._render();
    });
  },

  _bindBotoes(semanaId) {
    const bind = (id, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundExtBtn) { el.dataset.boundExtBtn = "1"; el.addEventListener("click", fn); }
    };
    bind("btnAdicionarExtra", () => this.abrirModal());
    bind("btnExtraInvestigador", () => this.abrirModal("investigador"));
    bind("btnExtraGuarda", () => this.abrirModal("guarda"));
    bind("salvarExtra", () => this.salvar());
    bind("cancelModalExtra", () => AdminUtils.closeModal("modalExtra"));
    bind("closeModalExtra", () => AdminUtils.closeModal("modalExtra"));
  }
};
