// ============================================================
// admin-relatorios.js — Relatórios inteligentes do Admin Homy
// Filtro por período, por dia, por centro de custo e por funcionário
// Exportação Excel formatada padrão Homy
// ============================================================

const AdminRelatorios = window.AdminRelatorios = {

  _pedidos: [],
  _periodo: { ini: "", fim: "" },

  async load(semanaId) {
    const datas = SP.getWeekDates(semanaId);
    const ini = datas[0].toISOString().slice(0, 10);
    const fim = datas[4].toISOString().slice(0, 10);

    AdminUtils.setVal("relDataIni", ini);
    AdminUtils.setVal("relDataFim", fim);

    this._bindControles();

    await this._buscar(ini, fim);
  },

  async _buscar(ini, fim) {
    this._periodo = { ini, fim };

    const wrap = document.getElementById("relConteudo");
    if (wrap) wrap.innerHTML = `<div class="alert alert-info">Carregando...</div>`;

    try {
      await SP.init();

      const todos = await SP.getItems("Pedidos");

      this._pedidos = todos.filter(p => {
        const dataPedido = String(SP.pick(p, "Data_Hora", "Data") || "").slice(0, 10);
        return dataPedido && dataPedido >= ini && dataPedido <= fim;
      });

      this._renderCards();
      this._renderTipo();

    } catch (e) {
      console.error("[AdminRelatorios] buscar:", e);
      if (wrap) {
        wrap.innerHTML = `<div class="alert" style="background:rgba(220,50,50,.1);color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</div>`;
      }
    }
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

  _renderCards() {
    const confirmados = this._pedidos.filter(p => this._isConfirmado(p));

    AdminUtils.setTxt("rel-principal", this._countOp(confirmados, "principal"));
    AdminUtils.setTxt("rel-light", this._countOp(confirmados, "light"));
    AdminUtils.setTxt("rel-carne", this._countOp(confirmados, "carne"));
    AdminUtils.setTxt("rel-massa", this._countOp(confirmados, "massa"));
    AdminUtils.setTxt("rel-lanche", this._countOp(confirmados, "lanche"));
    AdminUtils.setTxt("rel-total", confirmados.length);
  },

  _renderTipo() {
    const tipo = AdminUtils.getVal("relTipo") || "dia";
    const wrap = document.getElementById("relConteudo");
    if (!wrap) return;

    const confirmados = this._pedidos.filter(p => this._isConfirmado(p));

    if (tipo === "dia") {
      this._renderPorDia(confirmados, wrap);
      return;
    }

    if (tipo === "cc") {
      this._renderPorCentroCusto(confirmados, wrap);
      return;
    }

    if (tipo === "ccfunc") {
      this._renderPorCentroCustoFuncionario(confirmados, wrap);
      return;
    }

    wrap.innerHTML = `<div class="alert alert-warning">Tipo de relatório inválido.</div>`;
  },

  _renderPorDia(confirmados, wrap) {
    const linhas = this._dadosPorDia(confirmados);

    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">📅 Por dia</div>
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
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">🏢 Por centro de custo</div>
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
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">👤 Por centro de custo e funcionário</div>
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
      const cc = AdminUtils.limparTextoRelatorio(SP.pick(p, "Centro_Custo")) || "Sem CC";

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
      const cc = AdminUtils.limparTextoRelatorio(SP.pick(p, "Centro_Custo")) || "Sem CC";
      const colaborador =
        AdminUtils.limparTextoRelatorio(SP.pick(p, "Colaborador_nome", "Nome", "Title")) ||
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
    const tipo = AdminUtils.getVal("relTipo") || "dia";
    const confirmados = this._pedidos.filter(p => this._isConfirmado(p));

    if (tipo === "dia") {
      this._exportarPorDia(confirmados);
      return;
    }

    if (tipo === "cc") {
      this._exportarPorCentroCusto(confirmados);
      return;
    }

    if (tipo === "ccfunc") {
      this._exportarPorCentroCustoFuncionario(confirmados);
      return;
    }

    AdminUtils.toast("Tipo de relatório inválido para exportação.", "error");
  },

  _exportarPorDia(confirmados) {
    const linhas = this._dadosPorDia(confirmados);

    const totais = {
      Data: "Total",
      Principal: linhas.reduce((s, l) => s + l.Principal, 0),
      Light: linhas.reduce((s, l) => s + l.Light, 0),
      Carne: linhas.reduce((s, l) => s + l.Carne, 0),
      Massa: linhas.reduce((s, l) => s + l.Massa, 0),
      Lanche: linhas.reduce((s, l) => s + l.Lanche, 0),
      Total: linhas.reduce((s, l) => s + l.Total, 0)
    };

    AdminUtils.exportarExcelFormatado({
      nomeArquivo: this._nomeArquivo("por-dia"),
      nomeAba: "Por dia",
      titulo: "Relatório de Refeições por Dia",
      periodo: this._periodoTexto(),
      colunas: [
        { key: "Data", label: "Data", width: 16 },
        { key: "Principal", label: "Principal", width: 14 },
        { key: "Light", label: "Light", width: 14 },
        { key: "Carne", label: "Carne", width: 14 },
        { key: "Massa", label: "Massa", width: 14 },
        { key: "Lanche", label: "Lanche", width: 14 },
        { key: "Total", label: "Total", width: 14 }
      ],
      linhas,
      totais
    });
  },

  _exportarPorCentroCusto(confirmados) {
    const linhas = this._dadosPorCentroCusto(confirmados);

    const totais = {
      Centro_Custo: "Total",
      Principal: linhas.reduce((s, l) => s + l.Principal, 0),
      Light: linhas.reduce((s, l) => s + l.Light, 0),
      Carne: linhas.reduce((s, l) => s + l.Carne, 0),
      Massa: linhas.reduce((s, l) => s + l.Massa, 0),
      Lanche: linhas.reduce((s, l) => s + l.Lanche, 0),
      Total: linhas.reduce((s, l) => s + l.Total, 0)
    };

    AdminUtils.exportarExcelFormatado({
      nomeArquivo: this._nomeArquivo("por-centro-custo"),
      nomeAba: "Por CC",
      titulo: "Relatório de Refeições por Centro de Custo",
      periodo: this._periodoTexto(),
      colunas: [
        { key: "Centro_Custo", label: "Centro de Custo", width: 26 },
        { key: "Principal", label: "Principal", width: 14 },
        { key: "Light", label: "Light", width: 14 },
        { key: "Carne", label: "Carne", width: 14 },
        { key: "Massa", label: "Massa", width: 14 },
        { key: "Lanche", label: "Lanche", width: 14 },
        { key: "Total", label: "Total", width: 14 }
      ],
      linhas,
      totais
    });
  },

  _exportarPorCentroCustoFuncionario(confirmados) {
    const linhas = this._dadosPorCentroCustoFuncionario(confirmados);

    const totais = {
      Centro_Custo: "Total",
      Colaborador: "",
      Total_Refeicoes: linhas.reduce((s, l) => s + l.Total_Refeicoes, 0),
      Periodo_Inicio: this._periodo.ini,
      Periodo_Fim: this._periodo.fim
    };

    AdminUtils.exportarExcelFormatado({
      nomeArquivo: this._nomeArquivo("por-centro-custo-funcionario"),
      nomeAba: "Por funcionário",
      titulo: "Relatório para Desconto em Folha",
      periodo: this._periodoTexto(),
      colunas: [
        { key: "Centro_Custo", label: "Centro de Custo", width: 26 },
        { key: "Colaborador", label: "Colaborador", width: 34 },
        { key: "Total_Refeicoes", label: "Total Refeições", width: 18 },
        { key: "Periodo_Inicio", label: "Período Início", width: 18 },
        { key: "Periodo_Fim", label: "Período Fim", width: 18 }
      ],
      linhas,
      totais
    });
  },

  _periodoTexto() {
    return `Período: ${this._periodo.ini} a ${this._periodo.fim}`;
  },

  _nomeArquivo(tipo) {
    return `relatorio-${tipo}-${this._periodo.ini}-${this._periodo.fim}.xlsx`;
  },

  _bindControles() {
    const btnFiltrar = document.getElementById("btnRelFiltrar");
    if (btnFiltrar && !btnFiltrar.dataset.boundRel) {
      btnFiltrar.dataset.boundRel = "1";
      btnFiltrar.addEventListener("click", async () => {
        const ini = AdminUtils.getVal("relDataIni");
        const fim = AdminUtils.getVal("relDataFim");

        if (!ini || !fim) {
          AdminUtils.toast("Informe data início e data fim.", "error");
          return;
        }

        if (ini > fim) {
          AdminUtils.toast("A data início não pode ser maior que a data fim.", "error");
          return;
        }

        await this._buscar(ini, fim);
      });
    }

    const relTipo = document.getElementById("relTipo");
    if (relTipo && !relTipo.dataset.boundRel) {
      relTipo.dataset.boundRel = "1";
      relTipo.addEventListener("change", () => this._renderTipo());
    }

    const btnExportar = document.getElementById("btnRelExportar");
    if (btnExportar && !btnExportar.dataset.boundRel) {
      btnExportar.dataset.boundRel = "1";
      btnExportar.addEventListener("click", () => this.exportar());
    }
  }
};
