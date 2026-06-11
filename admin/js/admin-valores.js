// admin-valores.js — Valores de Refeição + NF Vascon
// Correções: botão de upload, valor unitário automático e reconciliação por período

const AdminValores = window.AdminValores = {
  _lista: [],
  _pedidosPeriodo: [],
  _editandoId: null,
  _nfFile: null,

  COL_DESCONTO: "Valor_Desconto_Funcionário",
  COL_DESCONTO_ASCII: "Valor_Desconto_Funcionario",
  COL_VASCON: "Valor_Vascon",
  COL_INICIO: "Data_Inicio",
  COL_FIM: "Data_Fim",
  COL_OBS: "Observacao",
  COL_ATIVO: "Ativo",

  async load() {
    this._bindBotoes();
    await this._carregar();
    await this._preencherPainelNFPadrao();
  },

  _getEl(...ids) {
    for (const id of ids) {
      const el = document.getElementById(id);
      if (el) return el;
    }
    return null;
  },

  _setTxtAny(valor, ...ids) {
    const el = this._getEl(...ids);
    if (el) el.textContent = valor ?? "—";
  },

  _setValAny(valor, ...ids) {
    const el = this._getEl(...ids);
    if (el) el.value = valor ?? "";
  },

  _valAny(...ids) {
    const el = this._getEl(...ids);
    return (el?.value || "").trim();
  },

  _pick(obj, ...keys) {
    for (const k of keys) {
      const v = SP.pick ? SP.pick(obj, k) : obj?.[k];
      if (v !== undefined && v !== null && String(v) !== "") return v;
    }
    return "";
  },

  _dateISO(v) {
    if (!v) return "";
    const s = String(v);
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const d = new Date(s);
    return isNaN(d) ? "" : d.toISOString().slice(0, 10);
  },

  _brDate(iso) {
    const d = this._dateISO(iso);
    if (!d) return "—";
    const [a, m, dia] = d.split("-");
    return `${dia}/${m}/${a}`;
  },

  _money(v) {
    const n = Number(v || 0);
    return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  },

  _moedaInput(v) {
    if (AdminUtils.moeda) return AdminUtils.moeda(v);
    if (v === null || v === undefined || String(v).trim() === "") return null;
    let s = String(v).replace(/R\$/gi, "").replace(/\s/g, "").trim();
    if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
    s = s.replace(/[^0-9.-]/g, "");
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  },

  _numValor(v, ...keys) {
    const bruto = keys.length ? this._pick(v, ...keys) : v;
    const n = this._moedaInput(bruto);
    return n === null ? 0 : n;
  },

  _isAtivo(v) {
    return SP.isTrue ? SP.isTrue(this._pick(v, this.COL_ATIVO, "Ativo", "ativo")) : !!this._pick(v, this.COL_ATIVO, "Ativo", "ativo");
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
      console.error("[Valores]", e);
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

    const lista = [...this._lista].sort((a, b) => {
      const aa = this._dateISO(this._pick(a, this.COL_INICIO, "DataInicio"));
      const bb = this._dateISO(this._pick(b, this.COL_INICIO, "DataInicio"));
      return bb.localeCompare(aa);
    });

    tbody.innerHTML = lista.map(v => {
      const id = AdminUtils.esc(v.id || "");
      const titulo = AdminUtils.esc(this._pick(v, "Title", "Titulo", "Título") || "—");
      const inicio = this._dateISO(this._pick(v, this.COL_INICIO, "DataInicio"));
      const fim = this._dateISO(this._pick(v, this.COL_FIM, "DataFim"));
      const vascon = this._numValor(v, this.COL_VASCON, "ValorVascon");
      const desc = this._numValor(v, this.COL_DESCONTO, this.COL_DESCONTO_ASCII, "Valor_Desconto", "Desconto");
      const ativo = this._isAtivo(v);

      return `<tr>
        <td>${titulo}</td>
        <td>${this._brDate(inicio)}</td>
        <td>${this._brDate(fim)}</td>
        <td>${this._money(vascon)}</td>
        <td>${this._money(desc)}</td>
        <td><span class="badge ${ativo ? "badge-green" : "badge-red"}">${ativo ? "Ativo" : "Inativo"}</span></td>
        <td><div class="table-actions">
          <button class="btn-icon" title="Editar" onclick="AdminValores.abrirEdicao('${id}')">✏️</button>
          <button class="btn-icon" title="Usar na NF" onclick="AdminValores.usarValorNaNF('${id}')">🧾</button>
          <button class="btn-icon danger" title="Excluir" onclick="AdminValores.excluir('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  _valorObjNormalizado(v) {
    if (!v) return null;
    return {
      id: v.id || "",
      titulo: this._pick(v, "Title", "Titulo", "Título") || "Valor refeição",
      inicio: this._dateISO(this._pick(v, this.COL_INICIO, "DataInicio")),
      fim: this._dateISO(this._pick(v, this.COL_FIM, "DataFim")),
      vascon: this._numValor(v, this.COL_VASCON, "ValorVascon"),
      desconto: this._numValor(v, this.COL_DESCONTO, this.COL_DESCONTO_ASCII, "Valor_Desconto", "Desconto"),
      ativo: this._isAtivo(v),
      raw: v
    };
  },

  _periodosSobrepoem(iniA, fimA, iniB, fimB) {
    if (!iniA || !fimA || !iniB || !fimB) return false;
    return iniA <= fimB && fimA >= iniB;
  },

  _valorParaPeriodo(ini, fim) {
    const valores = this._lista.map(v => this._valorObjNormalizado(v)).filter(Boolean);
    if (!valores.length) return null;

    if (ini && fim) {
      const ativoPeriodo = valores.find(v => v.ativo && this._periodosSobrepoem(ini, fim, v.inicio, v.fim));
      if (ativoPeriodo) return ativoPeriodo;

      const periodo = valores.find(v => this._periodosSobrepoem(ini, fim, v.inicio, v.fim));
      if (periodo) return periodo;
    }

    return valores.find(v => v.ativo) || valores[0];
  },

  _validarSobreposicao(inicio, fim, idAtual = null) {
    return this._lista
      .map(v => this._valorObjNormalizado(v))
      .filter(v => v && String(v.id) !== String(idAtual || ""))
      .find(v => this._periodosSobrepoem(inicio, fim, v.inicio, v.fim));
  },

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

    AdminUtils.setVal("valorTitulo", this._pick(v, "Title", "Titulo", "Título") || "");
    AdminUtils.setVal("valorDataInicio", this._dateISO(this._pick(v, this.COL_INICIO, "DataInicio")));
    AdminUtils.setVal("valorDataFim", this._dateISO(this._pick(v, this.COL_FIM, "DataFim")));
    AdminUtils.setVal("valorVascon", String(this._numValor(v, this.COL_VASCON, "ValorVascon")).replace(".", ","));
    AdminUtils.setVal("valorDesconto", String(this._numValor(v, this.COL_DESCONTO, this.COL_DESCONTO_ASCII, "Valor_Desconto", "Desconto")).replace(".", ","));
    AdminUtils.setVal("valorObs", this._pick(v, this.COL_OBS, "Observação", "Obs") || "");
    AdminUtils.setVal("valorAtivo", this._isAtivo(v) ? "sim" : "nao");

    const t = document.querySelector("#modalValorRefeicao .modal-title");
    if (t) t.textContent = "Editar valor de refeição";
    AdminUtils.openModal("modalValorRefeicao");
  },

  _limparModal() {
    ["valorTitulo", "valorDataInicio", "valorDataFim", "valorVascon", "valorDesconto", "valorObs"]
      .forEach(id => AdminUtils.setVal(id, ""));
    AdminUtils.setVal("valorAtivo", "sim");
  },

  async _desativarOutrosAtivos(idAtual = null) {
    const ativos = this._lista.filter(v => this._isAtivo(v) && String(v.id) !== String(idAtual || ""));
    await Promise.all(ativos.map(v => SP.updateItem("Valores de Refeição", v.id, { [this.COL_ATIVO]: false, Ativo: false }).catch(() => null)));
  },

  async salvar() {
    const titulo = AdminUtils.getVal("valorTitulo") || "Valor refeição";
    const inicio = AdminUtils.getVal("valorDataInicio");
    const fim = AdminUtils.getVal("valorDataFim");
    const vascon = this._moedaInput(AdminUtils.getVal("valorVascon"));
    const desconto = this._moedaInput(AdminUtils.getVal("valorDesconto"));
    const obs = AdminUtils.getVal("valorObs");
    const ativo = AdminUtils.getVal("valorAtivo") !== "nao";

    if (!inicio || !fim) { AdminUtils.toast("Informe data início e fim.", "error"); return; }
    if (inicio > fim) { AdminUtils.toast("Data início maior que data fim.", "error"); return; }
    if (vascon === null) { AdminUtils.toast("Informe o valor Vascon.", "error"); return; }
    if (desconto === null) { AdminUtils.toast("Informe o desconto do funcionário.", "error"); return; }

    const sobreposto = this._validarSobreposicao(inicio, fim, this._editandoId);
    if (sobreposto) {
      const v = this._valorObjNormalizado(sobreposto.raw || sobreposto);
      AdminUtils.toast(`Período sobreposto com '${v.titulo}' (${this._brDate(v.inicio)} a ${this._brDate(v.fim)}).`, "error");
      return;
    }

    const fields = {
      Title: titulo,
      [this.COL_INICIO]: inicio,
      [this.COL_FIM]: fim,
      [this.COL_VASCON]: vascon,
      [this.COL_DESCONTO]: desconto,
      [this.COL_OBS]: obs,
      [this.COL_ATIVO]: ativo
    };

    try {
      await SP.init();
      if (ativo) await this._desativarOutrosAtivos(this._editandoId);

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
      await this._preencherPainelNFPadrao();
    } catch (e) {
      AdminUtils.toast("Erro ao salvar: " + (e.message || e), "error");
    }
  },

  async excluir(id) {
    if (!confirm("Excluir este valor?")) return;
    try {
      await SP.init();
      await SP.deleteItem("Valores de Refeição", id);
      this._lista = this._lista.filter(v => String(v.id) !== String(id));
      this._render();
      await this._preencherPainelNFPadrao();
      AdminUtils.toast("Valor excluído.", "success");
    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + (e.message || e), "error");
    }
  },

  usarValorNaNF(id) {
    const valor = this._valorObjNormalizado(this._lista.find(v => String(v.id) === String(id)));
    if (!valor) return;
    this._setValAny(valor.id, "nfValorId");
    this._setValAny(valor.inicio, "nfInicio");
    this._setValAny(valor.fim, "nfFim");
    this._setTxtAny(this._money(valor.vascon), "nfVasconUnit");
    AdminUtils.toast(`Valor ${this._money(valor.vascon)} carregado para reconciliação.`, "success");
  },

  async _preencherPainelNFPadrao() {
    const iniEl = this._getEl("nfInicio");
    const fimEl = this._getEl("nfFim");
    if (!iniEl && !fimEl && !this._getEl("nfVasconUnit")) return;

    const hoje = new Date().toISOString().slice(0, 10);
    const valor = this._valorParaPeriodo(iniEl?.value || hoje, fimEl?.value || hoje) || this._valorParaPeriodo();
    if (!valor) {
      this._setTxtAny("—", "nfVasconUnit");
      return;
    }

    if (iniEl && !iniEl.value) iniEl.value = valor.inicio;
    if (fimEl && !fimEl.value) fimEl.value = valor.fim;
    this._setValAny(valor.id, "nfValorId");
    this._setTxtAny(this._money(valor.vascon), "nfVasconUnit");
  },

  async _atualizarValorUnitarioNF() {
    const ini = this._valAny("nfInicio");
    const fim = this._valAny("nfFim");
    const valor = this._valorParaPeriodo(ini, fim);
    if (valor) {
      this._setValAny(valor.id, "nfValorId");
      this._setTxtAny(this._money(valor.vascon), "nfVasconUnit");
    } else {
      this._setValAny("", "nfValorId");
      this._setTxtAny("—", "nfVasconUnit");
    }
  },

  _dataPedido(p) {
    const direta = this._dateISO(this._pick(p, "Data_Hora", "Data", "Data_Referencia"));
    if (direta) return direta;
    const semana = this._pick(p, "Semana_id", "Semana");
    const dia = this._pick(p, "Dia");
    if (semana && dia && SP.getDataRefBySemanaDia) return SP.getDataRefBySemanaDia(semana, dia);
    return "";
  },

  _isConf(p) {
    const s = AdminUtils.norm(this._pick(p, "Status") || "");
    return ["confirmado", "extra", "aprovado", "travado"].includes(s) || (SP.isTrue && SP.isTrue(this._pick(p, "Confirmado")));
  },

  async _pedidosConfirmadosPeriodo(ini, fim) {
    const todos = await SP.getItems("Pedidos");
    return todos.filter(p => {
      const d = this._dataPedido(p);
      return d && d >= ini && d <= fim && this._isConf(p);
    });
  },

  _abrirUpload() {
    const input = document.getElementById("inputNFPDF");
    if (!input) {
      AdminUtils.toast("Input de PDF não encontrado no HTML.", "error");
      return;
    }
    input.click();
  },

  _onPDFSelecionado(ev) {
    const file = ev?.target?.files?.[0];
    this._nfFile = file || null;
    const status = document.getElementById("nfUploadStatus");

    if (!file) {
      if (status) status.textContent = "";
      return;
    }

    const mb = file.size / 1024 / 1024;
    if (!/\.pdf$/i.test(file.name)) {
      AdminUtils.toast("Selecione um arquivo PDF.", "error");
      ev.target.value = "";
      this._nfFile = null;
      if (status) status.textContent = "";
      return;
    }

    if (status) status.textContent = `PDF selecionado: ${file.name} (${mb.toFixed(2)} MB)`;
    AdminUtils.toast("PDF selecionado para auditoria.", "success");
  },

  async reconciliar() {
    const ini = this._valAny("nfInicio");
    const fim = this._valAny("nfFim");
    const totalNF = this._moedaInput(this._valAny("nfTotalDigitado"));
    const out = document.getElementById("nfResultado");

    if (!ini || !fim) { AdminUtils.toast("Informe o período da NF.", "error"); return; }
    if (ini > fim) { AdminUtils.toast("Data início maior que data fim.", "error"); return; }
    if (totalNF === null) { AdminUtils.toast("Informe o total da NF Vascon.", "error"); return; }

    try {
      await SP.init();
      if (!this._lista.length) this._lista = await SP.getValoresRefeicao(false);

      const valor = this._valorParaPeriodo(ini, fim);
      if (!valor) {
        AdminUtils.toast("Nenhum valor Vascon cadastrado para o período.", "error");
        return;
      }

      this._setValAny(valor.id, "nfValorId");
      this._setTxtAny(this._money(valor.vascon), "nfVasconUnit");

      const pedidos = await this._pedidosConfirmadosPeriodo(ini, fim);
      const qtd = pedidos.length;
      const esperado = qtd * valor.vascon;
      const unitNF = qtd > 0 ? totalNF / qtd : 0;
      const diferenca = totalNF - esperado;
      const cls = Math.abs(diferenca) < 0.01 ? "alert-info" : "alert-warning";
      const pdfTxt = this._nfFile ? `<br>📎 PDF: <b>${AdminUtils.esc(this._nfFile.name)}</b>` : "";

      if (out) {
        out.innerHTML = `
          <div class="alert ${cls}" style="line-height:1.7">
            <b>Reconciliação concluída</b><br>
            Período: <b>${this._brDate(ini)} a ${this._brDate(fim)}</b><br>
            Refeições confirmadas no sistema: <b>${qtd}</b><br>
            Valor unitário cadastrado Vascon: <b>${this._money(valor.vascon)}</b><br>
            Total esperado pelo sistema: <b>${this._money(esperado)}</b><br>
            Total informado na NF: <b>${this._money(totalNF)}</b><br>
            Valor unitário calculado pela NF: <b>${this._money(unitNF)}</b><br>
            Diferença NF x sistema: <b>${this._money(diferenca)}</b>
            ${pdfTxt}
          </div>`;
      }

      AdminUtils.toast("Reconciliação calculada.", "success");
    } catch (e) {
      AdminUtils.toast("Erro ao reconciliar: " + (e.message || e), "error");
    }
  },

  async reconciliarModal() {
    const ini = this._valAny("nfInicioModal");
    const fim = this._valAny("nfFimModal");
    const totalNF = this._moedaInput(this._valAny("nfTotalModal"));
    const out = document.getElementById("nfResultadoModal");

    if (!ini || !fim) { AdminUtils.toast("Informe o período.", "error"); return; }
    if (totalNF === null) { AdminUtils.toast("Informe o total da NF.", "error"); return; }

    try {
      const valor = this._valorParaPeriodo(ini, fim);
      const pedidos = await this._pedidosConfirmadosPeriodo(ini, fim);
      const qtd = pedidos.length;
      const esperado = qtd * (valor?.vascon || 0);
      const diferenca = totalNF - esperado;

      if (out) {
        out.innerHTML = `<div class="alert alert-info" style="line-height:1.7">
          Refeições: <b>${qtd}</b><br>
          Valor Vascon: <b>${this._money(valor?.vascon || 0)}</b><br>
          Total sistema: <b>${this._money(esperado)}</b><br>
          NF informada: <b>${this._money(totalNF)}</b><br>
          Diferença: <b>${this._money(diferenca)}</b>
        </div>`;
      }
    } catch (e) {
      AdminUtils.toast("Erro ao reconciliar: " + (e.message || e), "error");
    }
  },

  _bindBotoes() {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundVal) {
        el.dataset.boundVal = "1";
        el.addEventListener(ev, fn);
      }
    };

    bind("btnNovoValor", "click", () => this.abrirNovo());
    bind("salvarValorRefeicao", "click", () => this.salvar());
    bind("cancelarValorRefeicao", "click", () => AdminUtils.closeModal("modalValorRefeicao"));

    bind("btnUploadNF", "click", () => this._abrirUpload());
    bind("inputNFPDF", "change", ev => this._onPDFSelecionado(ev));
    bind("btnReconciliar", "click", () => this.reconciliar());
    bind("nfInicio", "change", () => this._atualizarValorUnitarioNF());
    bind("nfFim", "change", () => this._atualizarValorUnitarioNF());

    bind("btnReconciliarModal", "click", () => this.reconciliarModal());
    bind("fecharModalNF", "click", () => AdminUtils.closeModal("modalNFVascon"));
  }
};
