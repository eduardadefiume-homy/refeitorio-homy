// ============================================================
// admin-valores.js — Valores de Refeição do Admin Homy
// Responsabilidade: tela e interação do módulo Valores
// A resolução dos nomes internos do SharePoint fica no sharepoint.js
// ============================================================

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

    tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      this._lista = await SP.getValoresRefeicao(false);
      this._render();
    } catch (e) {
      console.error("[AdminValores] carregar:", e);
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
      const id = AdminUtils.esc(v.id || "");
      const titulo = AdminUtils.esc(SP.pick(v, "Title", "Titulo") || "Valor refeição");

      const inicio = String(v.ValorDataInicio || "").slice(0, 10);
      const fim = String(v.ValorDataFim || "").slice(0, 10);

      const vascon = this._formatMoney(v.ValorVascon);
      const desconto = this._formatMoney(v.ValorDescontoFuncionario);

      const ativo = SP.isTrue(v.ValorAtivo);

      return `<tr>
        <td>${titulo}</td>
        <td>${inicio || ""}</td>
        <td>${fim || ""}</td>
        <td>${vascon}</td>
        <td>${desconto}</td>
        <td><span class="badge ${ativo ? "badge-green" : "badge-red"}">${ativo ? "Ativo" : "Inativo"}</span></td>
        <td>
          <div class="table-actions">
            <button class="btn-icon" title="Editar" onclick="AdminValores.abrirEdicao('${id}')">✏️</button>
            <button class="btn-icon danger" title="Excluir" onclick="AdminValores.excluir('${id}')">🗑️</button>
          </div>
        </td>
      </tr>`;
    }).join("");
  },

  _formatMoney(value) {
    const n = Number(value || 0);
    return n.toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL"
    });
  },

  abrirNovo() {
    this._editandoId = null;
    this._limparModal();

    const titulo = document.querySelector("#modalValorRefeicao .modal-title");
    if (titulo) titulo.textContent = "Novo valor de refeição";

    AdminUtils.openModal("modalValorRefeicao");
  },

  abrirEdicao(id) {
    this._editandoId = id;

    const item = this._lista.find(x => String(x.id) === String(id));
    if (!item) {
      AdminUtils.toast("Valor não encontrado.", "error");
      return;
    }

    AdminUtils.setVal("valorTitulo", SP.pick(item, "Title", "Titulo") || "Valor refeição");
    AdminUtils.setVal("valorDataInicio", String(item.ValorDataInicio || "").slice(0, 10));
    AdminUtils.setVal("valorDataFim", String(item.ValorDataFim || "").slice(0, 10));
    AdminUtils.setVal("valorVascon", item.ValorVascon ?? "");
    AdminUtils.setVal("valorDesconto", item.ValorDescontoFuncionario ?? "");
    AdminUtils.setVal("valorObs", item.ValorObservacao || "");
    AdminUtils.setVal("valorAtivo", SP.isTrue(item.ValorAtivo) ? "sim" : "nao");

    const titulo = document.querySelector("#modalValorRefeicao .modal-title");
    if (titulo) titulo.textContent = "Editar valor de refeição";

    AdminUtils.openModal("modalValorRefeicao");
  },

  _limparModal() {
    [
      "valorTitulo",
      "valorDataInicio",
      "valorDataFim",
      "valorVascon",
      "valorDesconto",
      "valorObs"
    ].forEach(id => AdminUtils.setVal(id, ""));

    AdminUtils.setVal("valorAtivo", "sim");
  },

  async salvar() {
    const dados = {
      titulo: AdminUtils.getVal("valorTitulo") || "Valor refeição",
      dataInicio: AdminUtils.getVal("valorDataInicio"),
      dataFim: AdminUtils.getVal("valorDataFim"),
      valorVascon: AdminUtils.getVal("valorVascon"),
      valorDesconto: AdminUtils.getVal("valorDesconto"),
      observacao: AdminUtils.getVal("valorObs"),
      ativo: AdminUtils.getVal("valorAtivo") || "sim"
    };

    if (!dados.dataInicio) {
      AdminUtils.toast("Informe a data de início.", "error");
      return;
    }

    if (!dados.dataFim) {
      AdminUtils.toast("Informe a data fim.", "error");
      return;
    }

    if (AdminUtils.moeda(dados.valorVascon) === null) {
      AdminUtils.toast("Informe o valor Vascon corretamente.", "error");
      return;
    }

    if (AdminUtils.moeda(dados.valorDesconto) === null) {
      AdminUtils.toast("Informe o desconto do funcionário corretamente.", "error");
      return;
    }

    try {
      await SP.init();

      if (this._editandoId) {
        await SP.updateValorRefeicao(this._editandoId, dados);
        AdminUtils.toast("Valor atualizado com sucesso.", "success");
      } else {
        await SP.createValorRefeicao(dados);
        AdminUtils.toast("Valor cadastrado com sucesso.", "success");
      }

      AdminUtils.closeModal("modalValorRefeicao");
      this._editandoId = null;
      await this._carregar();

    } catch (e) {
      console.error("[AdminValores] salvar:", e);
      AdminUtils.toast("Erro ao salvar: " + e.message, "error");
    }
  },

  async excluir(id) {
    if (!id) return;

    const confirmar = confirm("Deseja excluir este valor de refeição?");
    if (!confirmar) return;

    try {
      await SP.init();
      await SP.deleteValorRefeicao(id);
      AdminUtils.toast("Valor excluído com sucesso.", "success");
      await this._carregar();
    } catch (e) {
      console.error("[AdminValores] excluir:", e);
      AdminUtils.toast("Erro ao excluir: " + e.message, "error");
    }
  },

  _bindBotoes() {
    const btnNovo = document.getElementById("btnNovoValor");
    if (btnNovo && !btnNovo.dataset.boundValores) {
      btnNovo.dataset.boundValores = "1";
      btnNovo.addEventListener("click", () => this.abrirNovo());
    }

    const btnSalvar = document.getElementById("btnSalvarValor");
    if (btnSalvar && !btnSalvar.dataset.boundValores) {
      btnSalvar.dataset.boundValores = "1";
      btnSalvar.addEventListener("click", () => this.salvar());
    }
  }
};
