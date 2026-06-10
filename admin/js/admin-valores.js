// admin-valores.js — Valores de Refeição do Admin Homy

const AdminValores = window.AdminValores = {

  _lista: [],
  _editandoId: null,

  async load() {
    this._bindBotoes();
    await this._carregar();
  },

  async _carregar() {
    const tbody = document.getElementById("valoresTable");
    if (!tbody) return;

    tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();

      this._lista = await SP.getValoresRefeicao(false);

      this._lista.sort((a, b) => {
        const da = new Date(SP.pick(a, "Data_Inicio") || 0);
        const db = new Date(SP.pick(b, "Data_Inicio") || 0);
        return db - da;
      });

      this._render();

    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message || e)}</td></tr>`;
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
      const titulo = AdminUtils.esc(SP.pick(v, "Title", "Titulo") || "Valor de refeição");
      const inicio = AdminUtils.esc(this._fmtData(SP.pick(v, "Data_Inicio")) || "—");
      const fim = AdminUtils.esc(this._fmtData(SP.pick(v, "Data_Fim")) || "—");
      const valores = SP._resolveColunasValores(v);
      const ativo = SP.isTrue(SP.pick(v, "Ativo"));

      return `<tr>
        <td>${titulo}</td>
        <td>${inicio}</td>
        <td>${fim}</td>
        <td>${this._fmtMoeda(valores.valorVascon)}</td>
        <td>${this._fmtMoeda(valores.descontoFuncionario)}</td>
        <td><span class="badge ${ativo ? "badge-green" : "badge-red"}">${ativo ? "Ativo" : "Inativo"}</span></td>
        <td><div class="table-actions">
          <button class="btn-icon" title="Editar" onclick="AdminValores.abrirEdicao('${id}')">✏️</button>
          <button class="btn-icon danger" title="Excluir" onclick="AdminValores.excluir('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  abrirNovo() {
    this._editandoId = null;
    this._limparModal();

    const t = document.querySelector("#modalValor .modal-title");
    if (t) t.textContent = "Novo valor de refeição";

    AdminUtils.openModal("modalValor");
  },

  abrirEdicao(id) {
    const v = this._lista.find(x => String(x.id) === String(id));

    if (!v) {
      AdminUtils.toast("Registro não encontrado.", "error");
      return;
    }

    this._editandoId = id;

    const valores = SP._resolveColunasValores(v);

    AdminUtils.setVal("valorTitulo", SP.pick(v, "Title", "Titulo") || "Valor de refeição");
    AdminUtils.setVal("valorInicio", this._dateInput(SP.pick(v, "Data_Inicio")));
    AdminUtils.setVal("valorFim", this._dateInput(SP.pick(v, "Data_Fim")));
    AdminUtils.setVal("valorVascon", this._numeroInput(valores.valorVascon));
    AdminUtils.setVal("valorDesconto", this._numeroInput(valores.descontoFuncionario));
    AdminUtils.setVal("valorObs", SP.pick(v, "Observacao") || "");

    const ativoEl = document.getElementById("valorAtivo");
    if (ativoEl) ativoEl.value = SP.isTrue(SP.pick(v, "Ativo")) ? "sim" : "nao";

    const t = document.querySelector("#modalValor .modal-title");
    if (t) t.textContent = "Editar valor de refeição";

    AdminUtils.openModal("modalValor");
  },

  _limparModal() {
    AdminUtils.setVal("valorTitulo", "Valor de refeição");
    AdminUtils.setVal("valorInicio", "");
    AdminUtils.setVal("valorFim", "");
    AdminUtils.setVal("valorVascon", "");
    AdminUtils.setVal("valorDesconto", "");
    AdminUtils.setVal("valorObs", "");

    const ativoEl = document.getElementById("valorAtivo");
    if (ativoEl) ativoEl.value = "sim";
  },

  async salvar() {
    const titulo = AdminUtils.getVal("valorTitulo") || "Valor de refeição";
    const inicio = AdminUtils.getVal("valorInicio");
    const fim = AdminUtils.getVal("valorFim");
    const valorVascon = SP.moneyToNumber(AdminUtils.getVal("valorVascon"));
    const descontoFuncionario = SP.moneyToNumber(AdminUtils.getVal("valorDesconto"));
    const observacao = AdminUtils.getVal("valorObs");
    const ativo = (AdminUtils.getVal("valorAtivo") || "sim") === "sim";

    if (!inicio) {
      AdminUtils.toast("Informe a data de início.", "error");
      return;
    }

    if (!fim) {
      AdminUtils.toast("Informe a data de fim.", "error");
      return;
    }

    if (!valorVascon) {
      AdminUtils.toast("Informe o valor Vascon.", "error");
      return;
    }

    try {
      await SP.init();

      const fields = {
        Title: titulo,
        Data_Inicio: inicio,
        Data_Fim: fim,
        Valor_Vascon: valorVascon,
        Valor_Desconto_Funcionario: descontoFuncionario,
        Observacao: observacao,
        Ativo: ativo
      };

      if (ativo) {
        await this._desativarOutros(this._editandoId);
      }

      if (this._editandoId) {
        await SP.updateValorRefeicao(this._editandoId, fields);
        AdminUtils.toast("Valor atualizado.", "success");
      } else {
        await SP.createValorRefeicao(fields);
        AdminUtils.toast("Valor cadastrado.", "success");
      }

      AdminUtils.closeModal("modalValor");
      this._editandoId = null;

      await this._carregar();

    } catch (e) {
      AdminUtils.toast("Erro ao salvar: " + (e.message || e), "error");
    }
  },

  async _desativarOutros(idAtual = null) {
    const lista = await SP.getValoresRefeicao(false).catch(() => []);

    for (const item of lista) {
      if (idAtual && String(item.id) === String(idAtual)) continue;
      if (!SP.isTrue(SP.pick(item, "Ativo"))) continue;

      try {
        await SP.updateValorRefeicao(item.id, { Ativo: false });
      } catch (e) {
        console.warn("[valores] falha ao desativar valor antigo:", e);
      }
    }
  },

  async excluir(id) {
    if (!confirm("Excluir este valor?")) return;

    try {
      await SP.init();
      await SP.deleteValorRefeicao(id);

      this._lista = this._lista.filter(v => String(v.id) !== String(id));
      this._render();

      AdminUtils.toast("Valor removido.", "success");

    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + (e.message || e), "error");
    }
  },

  _fmtData(v) {
    if (!v) return "";

    const d = new Date(v);
    if (isNaN(d)) return v;

    return d.toLocaleDateString("pt-BR", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric"
    });
  },

  _dateInput(v) {
    if (!v) return "";

    const d = new Date(v);
    if (isNaN(d)) return String(v).slice(0, 10);

    return d.toISOString().slice(0, 10);
  },

  _fmtMoeda(v) {
    return Number(v || 0).toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL"
    });
  },

  _numeroInput(v) {
    return Number(v || 0).toFixed(2).replace(".", ",");
  },

  _bindBotoes() {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundValores) {
        el.dataset.boundValores = "1";
        el.addEventListener(ev, fn);
      }
    };

    bind("btnNovoValor", "click", () => this.abrirNovo());
    bind("salvarValor", "click", () => this.salvar());
    bind("cancelarValor", "click", () => AdminUtils.closeModal("modalValor"));
    bind("closeModalValor", "click", () => AdminUtils.closeModal("modalValor"));

    bind("btnUploadNF", "click", () => {
      const inp = document.getElementById("nfPdfInput");
      if (inp) {
        inp.value = "";
        inp.click();
      }
    });

    bind("btnReconciliarNF", "click", () => this.reconciliarNF());
  },

  async reconciliarNF() {
    const inicio = AdminUtils.getVal("nfPeriodoInicio");
    const fim = AdminUtils.getVal("nfPeriodoFim");
    const totalNF = SP.moneyToNumber(AdminUtils.getVal("nfTotal"));
    const valorAtivoEl = document.getElementById("nfValorUnitario");

    try {
      const valor = await SP.getValorRefeicaoAtivo();

      if (valor && valorAtivoEl) {
        const col = SP._resolveColunasValores(valor);
        valorAtivoEl.value = this._fmtMoeda(col.valorVascon);
      }

      if (!inicio || !fim || !totalNF) {
        AdminUtils.toast("Informe período e total da NF para reconciliar.", "error");
        return;
      }

      AdminUtils.toast("Reconciliação preparada. O cruzamento final sai no relatório.", "success");

    } catch (e) {
      AdminUtils.toast("Erro na reconciliação: " + (e.message || e), "error");
    }
  }
};
