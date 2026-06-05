// admin-valores.js — Valores de Refeição do Admin Homy

const AdminValores = window.AdminValores = {

  _lista: [],
  _editandoId: null,

  async load() {
    await this._carregar();
    this._bindBotoes();
  },

  async _carregar() {
    const tbody = document.getElementById("valoresTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      this._lista = await SP.getValoresRefeicao(false);
      this._render();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
    }
  },

  _render() {
    const tbody = document.getElementById("valoresTable");
    if (!tbody) return;

    if (!this._lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhum valor cadastrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = this._lista.map(v => {
      const id      = AdminUtils.esc(v.id || "");
      const titulo  = AdminUtils.esc(SP.pick(v, "Title")                   || "—");
      const inicio  = AdminUtils.esc(SP.pick(v, "Data_Inicio")             || "—");
      const fim     = AdminUtils.esc(SP.pick(v, "Data_Fim")                || "—");
      const vascon  = Number(SP.pick(v, "Valor_Vascon")                    || 0).toFixed(2);
      const desc    = Number(SP.pick(v, "Valor_Desconto_Funcionario")      || 0).toFixed(2);
      const ativo   = SP.isTrue(SP.pick(v, "Ativo"));
      return `<tr>
        <td>${titulo}</td>
        <td>${inicio ? inicio.slice(0, 10) : "—"}</td>
        <td>${fim    ? fim.slice(0, 10)    : "—"}</td>
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

    AdminUtils.setVal("valorTitulo",      SP.pick(v, "Title")                             || "");
    AdminUtils.setVal("valorDataInicio",  (SP.pick(v, "Data_Inicio") || "").slice(0, 10));
    AdminUtils.setVal("valorDataFim",     (SP.pick(v, "Data_Fim")    || "").slice(0, 10));
    AdminUtils.setVal("valorVascon",      SP.pick(v, "Valor_Vascon")                      || "");
    AdminUtils.setVal("valorDesconto",    SP.pick(v, "Valor_Desconto_Funcionario")        || "");
    AdminUtils.setVal("valorObs",         SP.pick(v, "Observacao")                        || "");
    AdminUtils.setVal("valorAtivo",       SP.isTrue(SP.pick(v, "Ativo")) ? "sim" : "nao");

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
    const titulo  = AdminUtils.getVal("valorTitulo")     || "Valor refeição";
    const inicio  = AdminUtils.getVal("valorDataInicio");
    const fim     = AdminUtils.getVal("valorDataFim");
    const vascon  = AdminUtils.moeda(AdminUtils.getVal("valorVascon"));
    const desconto = AdminUtils.moeda(AdminUtils.getVal("valorDesconto"));
    const obs     = AdminUtils.getVal("valorObs");
    const ativo   = AdminUtils.getVal("valorAtivo") !== "nao";

    if (!inicio || !fim)   { AdminUtils.toast("Informe data início e fim.", "error"); return; }
    if (vascon === null)    { AdminUtils.toast("Informe o valor Vascon.", "error"); return; }
    if (desconto === null)  { AdminUtils.toast("Informe o valor descontado do funcionário.", "error"); return; }

    try {
      await SP.init();
      const dados = { title: titulo, dataInicio: inicio, dataFim: fim, valorVascon: vascon, valorDesconto: desconto, observacao: obs, ativo };

      if (this._editandoId) {
        await SP.updateValorRefeicao(this._editandoId, dados);
        AdminUtils.toast("Valor atualizado.", "success");
      } else {
        await SP.createValorRefeicao(dados);
        AdminUtils.toast("Valor criado.", "success");
      }

      AdminUtils.closeModal("modalValorRefeicao");
      this._editandoId = null;
      await this._carregar();
    } catch (e) {
      AdminUtils.toast("Erro ao salvar valor: " + e.message, "error");
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

  // ── Bindings ─────────────────────────────────────────────────
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
