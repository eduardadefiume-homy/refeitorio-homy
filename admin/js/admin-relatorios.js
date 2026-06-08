// ============================================================
// admin-relatorios.js — Relatórios do Admin Homy
// Correção: bind robusto dos botões Gerar relatório e Exportar Excel
// ============================================================

const AdminRelatorios = window.AdminRelatorios = {

  _pedidos: [],
  _periodo: { ini: "", fim: "" },
  _bound: false,

  async load(semanaId) {
    this._bindControles();

    const iniAtual = AdminUtils.getVal("relDataIni");
    const fimAtual = AdminUtils.getVal("relDataFim");

    if (!iniAtual || !fimAtual) {
      const datas = SP.getWeekDates(semanaId);
      const ini = datas[0].toISOString().slice(0, 10);
      const fim = datas[4].toISOString().slice(0, 10);

      AdminUtils.setVal("relDataIni", ini);
      AdminUtils.setVal("relDataFim", fim);
    }

    await this.gerar();
  },

  _getEl(...ids) {
    for (const id of ids) {
      const el = document.getElementById(id);
      if (el) return el;
    }

    return null;
  },

  _findButtonByText(texto) {
    const alvo = String(texto || "").toLowerCase();

    return Array.from(document.querySelectorAll("button"))
      .find(btn => String(btn.textContent || "").toLowerCase().includes(alvo));
  },

  _bindControles() {
    if (this._bound) return;
    this._bound = true;

    const btnGerar =
      this._getEl("btnRelFiltrar", "btnGerarRelatorio", "btnRelGerar", "btnGerarRel") ||
      this._findButtonByText("gerar relatório");

    const btnExportar =
      this._getEl("btnRelExportar", "btnExportarExcel", "btnExportarRelatorio", "btnExportExcel") ||
      this._findButtonByText("exportar excel");

    const tipo =
      this._getEl("relTipo", "tipoRelatorio", "selectTipoRelatorio");

    if (btnGerar) {
      btnGerar.addEventListener("click", () => this.gerar());
      console.log("[AdminRelatorios] Botão gerar vinculado:", btnGerar);
    } else {
      console.warn("[AdminRelatorios] Botão Gerar relatório não encontrado.");
    }

    if (btnExportar) {
      btnExportar.addEventListener("click", () => this.exportar());
      console.log("[AdminRelatorios] Botão exportar vinculado:", btnExportar);
    } else {
      console.warn("[AdminRelatorios] Botão Exportar Excel não encontrado.");
    }

    if (tipo) {
      tipo.addEventListener("change", () => this._renderTipo());
    }
  },

  async gerar() {
    const ini =
      AdminUtils.getVal("relDataIni") ||
      AdminUtils.getVal("dataInicioRelatorio") ||
      AdminUtils.getVal("relInicio");

    const fim =
      AdminUtils.getVal("relDataFim") ||
      AdminUtils.getVal("dataFimRelatorio") ||
      AdminUtils.getVal("relFim");

    if (!ini || !fim) {
      AdminUtils.toast("Informe data início e data fim.", "error");
      return;
    }

    if (ini > fim) {
      AdminUtils.toast("A data início não pode ser maior que a data fim.", "error");
      return;
    }

    await this._buscar(ini, fim);
  },

  async _buscar(ini, fim) {
    this._periodo = { ini, fim };

    const wrap = document.getElementById("relConteudo");
    if (wrap) {
      wrap.innerHTML = `<div class="alert alert-info">Carregando relatório...</div>`;
    }

    try {
      await SP.init();

      const todos = await SP.getItems("Pedidos");

      this._pedidos = todos.filter(p => {
        const dataPedido = String(SP.pick(p, "Data_Hora", "Data") || "").slice(0, 10);

        if (!dataPedido) return false;

        return dataPedido >= ini && dataPedido <= fim;
      });

      this._renderCards();
      this._renderTipo();

      AdminUtils.toast("Relatório gerado com sucesso.", "success");

    } catch (e) {
      console.error("[AdminRelatorios] buscar:", e);

      if (wrap) {
        wrap.innerHTML = `
          <div class="alert" style="background:rgba(220,50,50,.1);color:#ff8080">
            Erro ao gerar relatório: ${AdminUtils.esc(e.message || e)}
          </div>
        `;
      }

      AdminUtils.toast("Erro ao gerar relatório: " + (e.message || e), "error");
    }
  },

  _getTipo() {
    return (
      AdminUtils.getVal("relTipo") ||
      AdminUtils.getVal("tipoRelatorio") ||
      AdminUtils.getVal("selectTipoRelatorio") ||
      "dia"
    );
  },

  _isConfirmado(p) {
    const status = AdminUtils.norm(SP.pick(p, "Status") || "");

    return status === "confirmado" ||
      status === "extra" ||
      status === "aprovado" ||
      SP.isTrue(SP.pick(p, "Confirmado"));
  },

  _countOp(lista, opcao) {
    return lista.filter(p => AdminUtils.norm(SP.pick(p, "Opcao")) === opcao).length;
  },

  _setCard(id, value) {
    AdminUtils.setTxt(id, value);
  },

  _renderCards() {
    const confirmados = this._pedidos.filter(p => this._isConfirmado(p));

    this._setCard("rel-principal", this._countOp(confirmados, "principal"));
    this._setCard("rel-light", this._countOp(confirmados, "light"));
    this._setCard("rel-carne", this._countOp(confirmados, "carne"));
    this._setCard("rel-massa", this._countOp(confirmados, "massa"));
    this._setCard("rel-lanche", this._countOp(confirmados, "lanche"));
    this._setCard("rel-total", confirmados.length);
  },

  _renderTipo() {
    const tipo = this._getTipo();
    const wrap = document.getElementById("relConteudo");

    if (!wrap) {
      console.warn("[AdminRelatorios] relConteudo não encontrado.");
      return;
    }

    const confirmados = this._pedidos.filter(p => this._isConfirmado(p));

    if (tipo === "dia") {
      this._renderPorDia(confirmados, wrap);
      return;
    }

    if (tipo === "cc") {
      this._renderPorCentroCusto(confirmados, wrap);
      return;
    }

    if (tipo === "ccfunc" || tipo === "ccfuncionario" || tipo.includes("func")) {
      this._renderPorCentroCustoFuncionario(confirmados, wrap);
      return;
    }

    this._renderPorCentroCustoFuncionario(confirmados, wrap);
  },

  _renderPorDia(confirmados, wrap) {
    const linhas = this._dadosPorDia(confirmados);

    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">Por dia</div>
      <div class="table-wrap">
        <table class="table">
          <thead>
            <tr>
              <th>Data</th>
              <th>Principal</th>
              <th>Light</th>
              <th>Carne</th>
              <th>Massa</th>
              <th>Lanche</th>
              <th>Total</th>
            </tr>
          </thead>
          <tbody>
            ${linhas.length ? linhas.map(l => `
              <tr>
                <td>${AdminUtils.esc(l.Data)}</td>
                <td>${l.Principal}</td>
                <td>${l.Light}</td>
                <td>${l.Carne}</td>
                <td>${l.Massa}</td>
                <td>${l.Lanche}</td>
                <td><strong>${l.Total}</strong></td>
              </tr>
            `).join("") : `<tr><td colspan="7" class="empty-cell">Nenhum pedido no período.</td></tr>`}
          </tbody>
        </table>
      </div>
    `;
  },

  _renderPorCentroCusto(confirmados, wrap) {
    const linhas = this._dadosPorCentroCusto(confirmados);

    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">Por centro de custo</div>
      <div class="table-wrap">
        <table class="table">
          <thead>
            <tr>
              <th>Centro de Custo</th>
              <th>Principal</th>
              <th>Light</th>
              <th>Carne</th>
              <th>Massa</th>
              <th>Lanche</th>
              <th>Total</th>
            </tr>
          </thead>
          <tbody>
            ${linhas.length ? linhas.map(l => `
              <tr>
                <td>${AdminUtils.esc(l.Centro_Custo)}</td>
                <td>${l.Principal}</td>
                <td>${l.Light}</td>
                <td>${l.Carne}</td>
                <td>${l.Massa}</td>
                <td>${l.Lanche}</td>
                <td><strong>${l.Total}</strong></td>
              </tr>
            `).join("") : `<tr><td colspan="7" class="empty-cell">Nenhum pedido no período.</td></tr>`}
          </tbody>
        </table>
      </div>
    `;
  },

  _renderPorCentroCustoFuncionario(confirmados, wrap) {
    const linhas = this._dadosPorCentroCustoFuncionario(confirmados);

    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">Por centro de custo e funcionário</div>
      <div class="table-wrap">
        <table class="table">
          <thead>
            <tr>
              <th>Centro de Custo</th>
              <th>Colaborador</th>
              <th>Total Refeições</th>
              <th>Período Início</th>
              <th>Período Fim</th>
            </tr>
          </thead>
          <tbody>
            ${linhas.length ? linhas.map(l => `
              <tr>
                <td>${AdminUtils.esc(l.Centro_Custo)}</td>
                <td>${AdminUtils.esc(l.Colaborador)}</td>
                <td><strong>${l.Total_Refeicoes}</strong></td>
                <td>${AdminUtils.esc(l.Periodo_Inicio)}</td>
                <td>${AdminUtils.esc(l.Periodo_Fim)}</td>
              </tr>
            `).join("") : `<tr><td colspan="5" class="empty-cell">Nenhum pedido no período.</td></tr>`}
          </tbody>
        </table>
      </div>
    `;
  },

  _dadosPorDia(confirmados) {
    const mapa = {};

    confirmados.forEach(p => {
      const data = String(SP.pick(p, "Data_Hora", "Data") || "").slice(0, 10);
      if (!data) return;

      if (!mapa[data]) mapa[data] = [];
      mapa[data].push(p);
    });

    return Object.entries(mapa)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([data, lista]) => ({
        Data: data,
        Principal: this._countOp(lista, "principal"),
        Light: this._countOp(lista, "light"),
        Carne: this._countOp(lista, "carne"),
        Massa: this._countOp(lista, "massa"),
        Lanche: this._countOp(lista, "lanche"),
        Total: lista.length
      }));
  },

  _dadosPorCentroCusto(confirmados) {
    const mapa = {};

    confirmados.forEach(p => {
      const cc = this._limparTexto(SP.pick(p, "Centro_Custo")) || "Sem CC";

      if (!mapa[cc]) mapa[cc] = [];
      mapa[cc].push(p);
    });

    return Object.entries(mapa)
      .sort((a, b) => b[1].length - a[1].length)
      .map(([cc, lista]) => ({
        Centro_Custo: cc,
        Principal: this._countOp(lista, "principal"),
        Light: this._countOp(lista, "light"),
        Carne: this._countOp(lista, "carne"),
        Massa: this._countOp(lista, "massa"),
        Lanche: this._countOp(lista, "lanche"),
        Total: lista.length
      }));
  },

  _dadosPorCentroCustoFuncionario(confirmados) {
    const mapa = {};

    confirmados.forEach(p => {
      const cc = this._limparTexto(SP.pick(p, "Centro_Custo")) || "Sem CC";

      const colaborador =
        this._limparTexto(SP.pick(p, "Colaborador_nome", "Nome", "Title")) ||
        "Não informado";

      const chave = `${cc}||${colaborador}`;

      if (!mapa[chave]) {
        mapa[chave] = {
          Centro_Custo: cc,
          Colaborador: colaborador,
          Total_Refeicoes: 0,
          Periodo_Inicio: this._periodo.ini,
          Periodo_Fim: this._periodo.fim
        };
      }

      mapa[chave].Total_Refeicoes += 1;
    });

    return Object.values(mapa).sort((a, b) => {
      const cc = a.Centro_Custo.localeCompare(b.Centro_Custo);
      if (cc !== 0) return cc;
      return a.Colaborador.localeCompare(b.Colaborador);
    });
  },

  exportar() {
    const tipo = this._getTipo();
    const confirmados = this._pedidos.filter(p => this._isConfirmado(p));

    if (!this._pedidos.length) {
      AdminUtils.toast("Gere o relatório antes de exportar.", "error");
      return;
    }

    if (tipo === "dia") {
      this._exportarPorDia(confirmados);
      return;
    }

    if (tipo === "cc") {
      this._exportarPorCentroCusto(confirmados);
      return;
    }

    this._exportarPorCentroCustoFuncionario(confirmados);
  },

  _exportarPorDia(confirmados) {
    const linhas = this._dadosPorDia(confirmados);

    this._exportarExcelSimples({
      nomeArquivo: this._nomeArquivo("por-dia"),
      colunas: ["Data", "Principal", "Light", "Carne", "Massa", "Lanche", "Total"],
      linhas
    });
  },

  _exportarPorCentroCusto(confirmados) {
    const linhas = this._dadosPorCentroCusto(confirmados);

    this._exportarExcelSimples({
      nomeArquivo: this._nomeArquivo("por-centro-custo"),
      colunas: ["Centro_Custo", "Principal", "Light", "Carne", "Massa", "Lanche", "Total"],
      linhas
    });
  },

  _exportarPorCentroCustoFuncionario(confirmados) {
    const linhas = this._dadosPorCentroCustoFuncionario(confirmados);

    this._exportarExcelSimples({
      nomeArquivo: this._nomeArquivo("por-centro-custo-funcionario"),
      colunas: ["Centro_Custo", "Colaborador", "Total_Refeicoes", "Periodo_Inicio", "Periodo_Fim"],
      linhas
    });
  },

  _exportarExcelSimples({ nomeArquivo, colunas, linhas }) {
    if (!window.XLSX) {
      AdminUtils.toast("Biblioteca XLSX não carregada.", "error");
      return;
    }

    const dados = linhas.map(l => {
      const obj = {};
      colunas.forEach(c => {
        obj[c] = l[c] ?? "";
      });
      return obj;
    });

    const ws = XLSX.utils.json_to_sheet(dados);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Relatorio");
    XLSX.writeFile(wb, nomeArquivo);

    AdminUtils.toast("Excel exportado com sucesso.", "success");
  },

  _nomeArquivo(tipo) {
    return `relatorio-${tipo}-${this._periodo.ini}-${this._periodo.fim}.xlsx`;
  },

  _limparTexto(valor) {
    const texto = String(valor ?? "").trim();
    if (!texto || texto === "-" || texto === "—") return "";
    return texto;
  }
};
