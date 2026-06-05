// admin-extras.js — Extras / Visitantes do Admin Homy

const AdminExtras = window.AdminExtras = {

  _lista: [],

  async load(semanaId) {
    this._bindFiltros(semanaId);
    this._bindBotoes(semanaId);
    await this._carregar(semanaId);
  },

  async _carregar(semanaId) {
    const tbody = document.getElementById("extrasTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const todos = await SP.getExtras(semanaId);

      // Remove duplicatas de extra automático — mantém só 1 por dia
      const seenAuto = new Set();
      this._lista = todos.filter(e => {
        const norm = v => AdminUtils.norm(v);
        const nome  = norm(SP.pick(e, "Nome", "Title") || "");
        const tipo  = norm(SP.pick(e, "tipo", "Tipo")  || "");
        const dia   = norm(SP.pick(e, "Dia")           || "");
        const isAuto = nome.includes("refeicaoextra") || (nome.includes("extra") && tipo.includes("extra"));
        if (isAuto) {
          const k = `auto-${dia}`;
          if (seenAuto.has(k)) return false;
          seenAuto.add(k);
        }
        return true;
      });

      // Ordena por dia → tipo → nome
      const diaOrd = { segunda: 1, terca: 2, quarta: 3, quinta: 4, sexta: 5 };
      this._lista.sort((a, b) => {
        const da = diaOrd[AdminUtils.norm(SP.pick(a, "Dia"))] || 9;
        const db = diaOrd[AdminUtils.norm(SP.pick(b, "Dia"))] || 9;
        if (da !== db) return da - db;
        return AdminUtils.norm(SP.pick(a, "tipo")).localeCompare(AdminUtils.norm(SP.pick(b, "tipo")));
      });

      this._render();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
    }
  },

  _render() {
    const tbody = document.getElementById("extrasTable");
    if (!tbody) return;

    const txt  = AdminUtils.norm(AdminUtils.getVal("fExtraTexto"));
    const dia  = AdminUtils.norm(AdminUtils.getVal("fExtraDia"));
    const tipo = AdminUtils.norm(AdminUtils.getVal("fExtraTipo"));

    const lista = this._lista.filter(e => {
      const all = AdminUtils.norm([
        SP.pick(e, "Nome", "Title"), SP.pick(e, "Observacao"), SP.pick(e, "tipo", "Tipo"), SP.pick(e, "Dia")
      ].join(" "));
      return (!txt || all.includes(txt)) &&
             (!dia  || AdminUtils.norm(SP.pick(e, "Dia"))       === dia)  &&
             (!tipo || AdminUtils.norm(SP.pick(e, "tipo", "Tipo")).includes(tipo));
    });

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhum extra encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(e => {
      const id   = AdminUtils.esc(e.id || "");
      const nome = AdminUtils.esc(SP.pick(e, "Nome", "Title") || "Extra");
      const tipo = AdminUtils.esc(SP.pick(e, "tipo", "Tipo")  || "—");
      const dia  = AdminUtils.esc(SP.pick(e, "Dia")           || "—");
      const opc  = AdminUtils.esc(SP.pick(e, "Opcao")         || "principal");
      const obs  = AdminUtils.esc(SP.pick(e, "Observacao")    || "—");
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

  // ── Modal ────────────────────────────────────────────────────
  abrirModal(predefinido = "") {
    const modal = document.getElementById("modalExtra");
    if (!modal) return;
    modal.dataset.predefinido = predefinido;

    // Mostra/oculta campos conforme pré-definido
    ["extraNome", "extraTipo", "extraOpcao", "extraObs"].forEach(id => {
      const el = document.getElementById(id);
      if (el?.closest(".form-group"))
        el.closest(".form-group").style.display = predefinido ? "none" : "";
      if (!predefinido && el) el.value = "";
    });

    if (predefinido === "investigador") {
      AdminUtils.setVal("extraNome",  "Investigador");
      AdminUtils.setVal("extraTipo",  "investigador");
      AdminUtils.setVal("extraOpcao", "principal");
      AdminUtils.setVal("extraObs",   "Investigador");
    } else if (predefinido === "guarda") {
      AdminUtils.setVal("extraNome",  "Guarda");
      AdminUtils.setVal("extraTipo",  "guarda");
      AdminUtils.setVal("extraOpcao", "principal");
      AdminUtils.setVal("extraObs",   "Guarda");
    }

    AdminUtils.openModal("modalExtra");
  },

  async salvar() {
    const modal      = document.getElementById("modalExtra");
    const predefinido = modal?.dataset.predefinido || "";
    const semanaId   = AdminState.getSemanaId();
    const dia        = AdminUtils.getVal("extraDia") || "segunda";

    let nome, tipo, opcao, obs;

    if (predefinido === "investigador") {
      // Conta quantos já existem no dia para numerar
      const pedidos = await SP.getPedidos(semanaId).catch(() => []);
      const qtd = pedidos.filter(p =>
        AdminUtils.norm(SP.pick(p, "Dia")) === AdminUtils.norm(dia) &&
        AdminUtils.norm(SP.pick(p, "Origem", "tipo")).includes("investigador")
      ).length;
      nome  = `Investigador ${qtd + 1}`;
      tipo  = "investigador";
      opcao = "principal";
      obs   = "Investigador";
    } else if (predefinido === "guarda") {
      nome  = "Guarda";
      tipo  = "guarda";
      opcao = "principal";
      obs   = "Guarda";
    } else {
      nome  = AdminUtils.getVal("extraNome");
      tipo  = AdminUtils.getVal("extraTipo")  || "visitante";
      opcao = AdminUtils.getVal("extraOpcao") || "principal";
      obs   = AdminUtils.getVal("extraObs");
    }

    if (!nome) { AdminUtils.toast("Informe o nome do extra.", "error"); return; }

    try {
      await SP.init();
      // Grava em Extras E cria pedido correspondente
      await SP.addExtraPedido(semanaId, dia, nome, tipo, opcao, obs, SP.getUserName());
      AdminUtils.closeModal("modalExtra");
      AdminUtils.toast("Extra adicionado.", "success");
      await this._carregar(semanaId);
    } catch (e) {
      AdminUtils.toast("Erro ao salvar extra: " + e.message, "error");
    }
  },

  async excluir(id) {
    if (!confirm("Excluir este extra?")) return;
    try {
      await SP.init();
      const semanaId = AdminState.getSemanaId();
      const item     = this._lista.find(e => String(e.id) === String(id));
      if (item) await SP.deleteExtraComPedido(item);
      else await SP.deleteExtra(id);
      this._lista = this._lista.filter(e => String(e.id) !== String(id));
      this._render();
      AdminUtils.toast("Extra excluído.", "success");
    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + e.message, "error");
    }
  },

  // ── Bindings ─────────────────────────────────────────────────
  _bindFiltros(semanaId) {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundExt) { el.dataset.boundExt = "1"; el.addEventListener(ev, fn); }
    };
    bind("fExtraTexto", "input",  () => this._render());
    bind("fExtraDia",   "change", () => this._render());
    bind("fExtraTipo",  "change", () => this._render());
    bind("btnFiltrarExtras", "click", () => this._render());
    bind("btnLimparExtras",  "click", () => {
      ["fExtraTexto", "fExtraDia", "fExtraTipo"].forEach(id => AdminUtils.setVal(id, ""));
      this._render();
    });
  },

  _bindBotoes(semanaId) {
    const bind = (id, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundExtBtn) { el.dataset.boundExtBtn = "1"; el.addEventListener("click", fn); }
    };
    bind("btnAdicionarExtra",      () => this.abrirModal());
    bind("btnExtraInvestigador",   () => this.abrirModal("investigador"));
    bind("btnExtraGuarda",         () => this.abrirModal("guarda"));
    bind("salvarExtra",            () => this.salvar());
    bind("cancelModalExtra",       () => AdminUtils.closeModal("modalExtra"));
    bind("closeModalExtra",        () => AdminUtils.closeModal("modalExtra"));
  }
};
