// admin-pedidos.js — Pedidos do Admin Homy

const AdminPedidos = window.AdminPedidos = {

  _lista: [],

  _valorValido(v) {
    const s = String(v ?? "").trim();
    if (!s) return false;
    const n = AdminUtils.norm(s);
    return !["-", "_", "—", "sem", "sem setor", "undefined", "null"].includes(n);
  },

  _pedidoTemConteudo(p) {
    // Remove da tela pedidos quebrados criados por versões anteriores
    // com apenas Principal/Confirmado, mas sem colaborador e sem dia.
    const nome = SP.pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title");
    const dia = SP.pick(p, "Dia");
    const data = SP.pick(p, "Data_Hora", "Data", "Data_Referencia");
    return this._valorValido(nome) && (this._valorValido(dia) || this._valorValido(data));
  },

  async load(semanaId) {
    await this._buscar(semanaId);
    this._bindFiltros();
    this._bindExportar();
  },

  async _buscar(semanaId) {
    const tbody = document.getElementById("pedidosTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      this._lista = (await SP.getPedidos(semanaId)).filter(p => this._pedidoTemConteudo(p));
      this._render();
      this._atualizarInfo(`Exibindo pedidos da semana ${semanaId}.`);
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
    }
  },

  async _buscarPorPeriodo(ini, fim) {
    const tbody = document.getElementById("pedidosTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const todos = await SP.getItems("Pedidos");
      this._lista = todos.filter(p => {
        if (!this._pedidoTemConteudo(p)) return false;
        const d = String(SP.pick(p, "Data_Hora") || "").slice(0, 10);
        return d && d >= ini && d <= fim;
      });
      this._render();
      this._atualizarInfo(`Período ${this._brDate(ini)} a ${this._brDate(fim)} — ${this._lista.length} pedidos.`);
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
    }
  },

  _render() {
    const tbody   = document.getElementById("pedidosTable");
    if (!tbody) return;

    const dia   = AdminUtils.norm(AdminUtils.getVal("filtroDia"));
    const opcao = AdminUtils.norm(AdminUtils.getVal("filtroOpcao"));

    const lista = this._lista.filter(p => {
      const okDia   = !dia   || AdminUtils.norm(SP.pick(p, "Dia"))   === dia;
      const okOpcao = !opcao || AdminUtils.norm(SP.pick(p, "Opcao")) === opcao;
      return okDia && okOpcao;
    });

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhum pedido encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(p => {
      const id     = AdminUtils.esc(p.id || "");
      const nome   = AdminUtils.esc(SP.pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "—");
      const dia    = AdminUtils.esc(SP.pick(p, "Dia") || "—");
      const opcao  = AdminUtils.esc(SP.pick(p, "Opcao") || "—");
      const prato  = AdminUtils.esc(SP.pick(p, "Nome_Prato", "Prato") || "—");
      const status = SP.pick(p, "Status") || (SP.isTrue(SP.pick(p, "Confirmado")) ? "Confirmado" : "Pendente");
      return `<tr>
        <td>${nome}</td>
        <td>${dia}</td>
        <td><span class="badge badge-blue">${opcao}</span></td>
        <td>${prato}</td>
        <td>${AdminUtils.badge(status)}</td>
        <td><div class="table-actions">
          <button class="btn-icon" title="Editar"  onclick="AdminPedidos.abrirEdicao('${id}','${nome}','${dia}','${opcao}')">✏️</button>
          <button class="btn-icon danger" title="Excluir" onclick="AdminPedidos.excluir('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  _atualizarInfo(texto) {
    const el = document.getElementById("pedidosFiltroInfo");
    if (el) el.innerHTML = `ℹ️ ${texto}`;
  },

  _brDate(iso) {
    return iso ? iso.split("-").reverse().join("/") : "—";
  },

  // ── Edição ──────────────────────────────────────────────────
  _editandoId: null,

  abrirEdicao(id, nome, dia, opcao) {
    this._editandoId = id;
    AdminUtils.setVal("editPedidoId",   id);
    AdminUtils.setVal("editPedidoNome", nome);
    AdminUtils.setVal("editPedidoDia",  dia);
    AdminUtils.setVal("editPedidoOpcao", opcao);
    AdminUtils.openModal("modalEditarPedido");
  },

  async salvarEdicao() {
    const id    = AdminUtils.getVal("editPedidoId");
    const dia   = AdminUtils.getVal("editPedidoDia");
    const opcao = AdminUtils.getVal("editPedidoOpcao");
    if (!id) { AdminUtils.toast("Pedido inválido.", "error"); return; }

    try {
      await SP.init();
      await SP.updatePedido(id, { dia, opcao, Alterado_Por: SP.getUserName() });
      AdminUtils.closeModal("modalEditarPedido");
      AdminUtils.toast("Pedido atualizado.", "success");
      await this._buscar(AdminState.getSemanaId());
    } catch (e) {
      AdminUtils.toast("Erro ao salvar: " + e.message, "error");
    }
  },

  async excluir(id) {
    if (!confirm("Excluir este pedido?")) return;
    try {
      await SP.init();
      await SP.deletePedido(id);
      AdminUtils.toast("Pedido excluído.", "success");
      this._lista = this._lista.filter(p => String(p.id) !== String(id));
      this._render();
    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + e.message, "error");
    }
  },

  // ── Filtros ─────────────────────────────────────────────────
  _bindFiltros() {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundPed) { el.dataset.boundPed = "1"; el.addEventListener(ev, fn); }
    };

    bind("btnBuscarPedidosHistorico", "click", async () => {
      const semana = AdminUtils.getVal("filtroSemanaHistorico");
      const ini    = AdminUtils.getVal("filtroDataInicioPedidos");
      const fim    = AdminUtils.getVal("filtroDataFimPedidos");
      if (semana) await this._buscar(semana);
      else if (ini && fim) await this._buscarPorPeriodo(ini, fim);
      else await this._buscar(AdminState.getSemanaId());
    });

    bind("btnLimparPedidosHistorico", "click", async () => {
      ["filtroSemanaHistorico", "filtroDataInicioPedidos", "filtroDataFimPedidos", "filtroDia", "filtroOpcao"]
        .forEach(id => AdminUtils.setVal(id, ""));
      await this._buscar(AdminState.getSemanaId());
    });

    bind("filtroDia",   "change", () => this._render());
    bind("filtroOpcao", "change", () => this._render());

    // Botão salvar edição
    bind("salvarPedidoEdicao", "click", () => this.salvarEdicao());
    bind("cancelarPedidoEdicao", "click", () => AdminUtils.closeModal("modalEditarPedido"));
  },

  // ── Exportar Excel ──────────────────────────────────────────
  _bindExportar() {
    const btn = document.getElementById("btnExportarPedidos");
    if (btn && !btn.dataset.bound) {
      btn.dataset.bound = "1";
      btn.addEventListener("click", () => this._exportar());
    }
  },

  _exportar() {
    if (!this._lista.length) { AdminUtils.toast("Nenhum dado para exportar.", "info"); return; }
    if (typeof XLSX === "undefined") { AdminUtils.toast("Biblioteca XLSX não carregou.", "error"); return; }

    const linhas = this._lista.map(p => ({
      Semana:      SP.pick(p, "Semana_id")       || "",
      Colaborador: SP.pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "",
      Dia:         SP.pick(p, "Dia")              || "",
      Opcao:       SP.pick(p, "Opcao")            || "",
      Prato:       SP.pick(p, "Nome_Prato")       || "",
      Status:      SP.pick(p, "Status")           || "",
      DataHora:    SP.pick(p, "Data_Hora")        || ""
    }));

    const ws = XLSX.utils.json_to_sheet(linhas);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pedidos");
    XLSX.writeFile(wb, `pedidos-${AdminState.getSemanaId()}.xlsx`);
    AdminUtils.toast("Excel exportado.", "success");
  }
};
