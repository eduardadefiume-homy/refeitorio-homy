// admin-valores.js — Valores de Refeição do Admin Homy
// Nome interno real da coluna de desconto: Valor_Desconto_Funcionário (com acento)

const AdminValores = window.AdminValores = {

  _lista: [],
  _editandoId: null,

  // Nome interno exato no SharePoint (confirmado nas Configurações da lista)
  COL_DESCONTO: "Valor_Desconto_Funcionário",
  COL_VASCON:   "Valor_Vascon",
  COL_INICIO:   "Data_Inicio",
  COL_FIM:      "Data_Fim",
  COL_OBS:      "Observacao",
  COL_ATIVO:    "Ativo",

  async load() {
    await this._carregar();
    this._bindBotoes();
  },

  async _carregar() {
    const tbody = document.getElementById("valoresTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Carregando...</td></tr>`;
    try {
      await SP.init();
      this._lista = await SP.getValoresRefeicao(false);
      this._render();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
    }
  },

  _render() {
    const tbody = document.getElementById("valoresTable");
    if (!tbody) return;

    if (!this._lista.length) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Nenhum valor cadastrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = this._lista.map(v => {
      const id      = AdminUtils.esc(v.id || "");
      const titulo  = AdminUtils.esc(SP.pick(v, "Title", "Titulo") || "—");
      const inicio  = (SP.pick(v, this.COL_INICIO) || "").slice(0, 10);
      const fim     = (SP.pick(v, this.COL_FIM)    || "").slice(0, 10);
      // Usa o nome interno correto com acento
      const vascon  = Number(SP.pick(v, this.COL_VASCON)   || 0).toFixed(2);
      const desc    = Number(SP.pick(v, this.COL_DESCONTO) || 0).toFixed(2);
      const ativo   = SP.isTrue(SP.pick(v, this.COL_ATIVO));
      return `<tr>
        <td>${titulo}</td>
        <td>${inicio || "—"}</td>
        <td>${fim    || "—"}</td>
        <td>R$ ${vascon}</td>
        <td>R$ ${desc}</td>
        <td><span class="badge ${ativo ? "badge-green" : "badge-red"}">${ativo ? "Ativo" : "Inativo"}</span></td>
        <td><div class="table-actions">
          <button class="btn-icon" title="Editar"  onclick="AdminValores.abrirEdicao('${id}')">✏️</button>
          <button class="btn-icon danger" title="Excluir" onclick="AdminValores.excluir('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  // ── Modal ────────────────────────────────────────────────────
  abrirNovo() {
    this._editandoId = null;
    this._limparModal();
    const t = document.querySelector("#modalValorRefeicao .modal-title");
    if (t) t.textContent = "Novo valor de refeição";
    AdminUtils.openModal("modalValorRefeicao");
  },

  abrirEdicao(id) {
    this._editandoId = id;
    const v = this._lista.find(x => String(x.id) === String(id));
    if (!v) { AdminUtils.toast("Valor não encontrado.", "error"); return; }

    AdminUtils.setVal("valorTitulo",      SP.pick(v, "Title", "Titulo")     || "");
    AdminUtils.setVal("valorDataInicio",  (SP.pick(v, this.COL_INICIO) || "").slice(0, 10));
    AdminUtils.setVal("valorDataFim",     (SP.pick(v, this.COL_FIM)    || "").slice(0, 10));
    AdminUtils.setVal("valorVascon",      SP.pick(v, this.COL_VASCON)   || "");
    AdminUtils.setVal("valorDesconto",    SP.pick(v, this.COL_DESCONTO) || "");
    AdminUtils.setVal("valorObs",         SP.pick(v, this.COL_OBS)      || "");
    AdminUtils.setVal("valorAtivo",       SP.isTrue(SP.pick(v, this.COL_ATIVO)) ? "sim" : "nao");

    const t = document.querySelector("#modalValorRefeicao .modal-title");
    if (t) t.textContent = "Editar valor de refeição";
    AdminUtils.openModal("modalValorRefeicao");
  },

  _limparModal() {
    ["valorTitulo", "valorDataInicio", "valorDataFim", "valorVascon", "valorDesconto", "valorObs"]
      .forEach(id => AdminUtils.setVal(id, ""));
    AdminUtils.setVal("valorAtivo", "sim");
  },

  async salvar() {
    const titulo   = AdminUtils.getVal("valorTitulo")     || "Valor refeição";
    const inicio   = AdminUtils.getVal("valorDataInicio");
    const fim      = AdminUtils.getVal("valorDataFim");
    const vascon   = AdminUtils.moeda(AdminUtils.getVal("valorVascon"));
    const desconto = AdminUtils.moeda(AdminUtils.getVal("valorDesconto"));
    const obs      = AdminUtils.getVal("valorObs");
    const ativo    = AdminUtils.getVal("valorAtivo") !== "nao";

    if (!inicio || !fim)   { AdminUtils.toast("Informe data início e fim.",               "error"); return; }
    if (vascon === null)    { AdminUtils.toast("Informe o valor Vascon.",                  "error"); return; }
    if (desconto === null)  { AdminUtils.toast("Informe o desconto do funcionário.",       "error"); return; }

    // Monta fields usando os nomes internos EXATOS do SharePoint
    const fields = {
      Title:                        titulo,
      [this.COL_INICIO]:            inicio,
      [this.COL_FIM]:               fim,
      [this.COL_VASCON]:            vascon,
      [this.COL_DESCONTO]:          desconto,   // "Valor_Desconto_Funcionário" — com acento
      [this.COL_OBS]:               obs,
      [this.COL_ATIVO]:             ativo
    };

    try {
      await SP.init();
      if (this._editandoId) {
        await SP.updateItem("Valores de Refeição", this._editandoId, fields);
        AdminUtils.toast("Valor atualizado.", "success");
      } else {
        await SP.createItem("Valores de Refeição", fields);
        AdminUtils.toast("Valor criado.", "success");
      }
      AdminUtils.closeModal("modalValorRefeicao");
      this._editandoId = null;
      await this._carregar();
    } catch (e) {
      AdminUtils.toast("Erro ao salvar: " + e.message, "error");
    }
  },

  async excluir(id) {
    if (!confirm("Excluir este valor?")) return;
    try {
      await SP.init();
      await SP.deleteItem("Valores de Refeição", id);
      this._lista = this._lista.filter(v => String(v.id) !== String(id));
      this._render();
      AdminUtils.toast("Valor excluído.", "success");
    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + e.message, "error");
    }
  },

  _bindBotoes() {
    const bind = (id, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundVal) { el.dataset.boundVal = "1"; el.addEventListener("click", fn); }
    };
    bind("btnNovoValor",          () => this.abrirNovo());
    bind("salvarValorRefeicao",   () => this.salvar());
    bind("cancelarValorRefeicao", () => AdminUtils.closeModal("modalValorRefeicao"));
  }
};
