// ============================================================
// admin-relatorios.js — Relatórios do Admin Homy
// Exportação Excel no padrão visual Homy
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

  async exportar() {
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

    this._exportarExcelHomy({
      nomeArquivo: this._nomeArquivo("por-dia"),
      titulo: "RELATÓRIO REFEITÓRIO HOMY  QUANTIDADE POR DIA",
      periodo: this._periodoTexto(),
      resumo: [
        ["Total de refeições", linhas.reduce((s, l) => s + l.Total, 0)]
      ],
      colunas: [
        { key: "Data", label: "Data", width: 18 },
        { key: "Principal", label: "Principal", width: 15 },
        { key: "Light", label: "Light", width: 15 },
        { key: "Carne", label: "Carne", width: 15 },
        { key: "Massa", label: "Massa", width: 15 },
        { key: "Lanche", label: "Lanche", width: 15 },
        { key: "Total", label: "Total", width: 15 }
      ],
      linhas
    });
  },

  _exportarPorCentroCusto(confirmados) {
    const linhas = this._dadosPorCentroCusto(confirmados);

    this._exportarExcelHomy({
      nomeArquivo: this._nomeArquivo("por-centro-custo"),
      titulo: "RELATÓRIO REFEITÓRIO HOMY  QUANTIDADE POR CENTRO DE CUSTO",
      periodo: this._periodoTexto(),
      resumo: [
        ["Total de refeições", linhas.reduce((s, l) => s + l.Total, 0)]
      ],
      colunas: [
        { key: "Centro_Custo", label: "Centro de custo", width: 30 },
        { key: "Principal", label: "Principal", width: 15 },
        { key: "Light", label: "Light", width: 15 },
        { key: "Carne", label: "Carne", width: 15 },
        { key: "Massa", label: "Massa", width: 15 },
        { key: "Lanche", label: "Lanche", width: 15 },
        { key: "Total", label: "Total", width: 15 }
      ],
      linhas
    });
  },

  _exportarPorCentroCustoFuncionario(confirmados) {
    const linhas = this._dadosPorCentroCustoFuncionario(confirmados);

    this._exportarExcelHomy({
      nomeArquivo: this._nomeArquivo("por-centro-custo-funcionario"),
      titulo: "RELATÓRIO REFEITÓRIO HOMY  QUANTIDADE POR COLABORADOR",
      periodo: this._periodoTexto(),
      resumo: [
        ["Gerado em", this._hojeBR()],
        ["Total de refeições", linhas.reduce((s, l) => s + l.Total_Refeicoes, 0)],
        ["Valor unitário Vascon", 0],
        ["Valor unitário descontado", 0]
      ],
      colunas: [
        { key: "Colaborador", label: "Colaborador", width: 34 },
        { key: "Centro_Custo", label: "Centro de custo", width: 22 },
        { key: "Total_Refeicoes", label: "Quantidade", width: 18 },
        { key: "ValorUnitario", label: "Valor de cada refeição", width: 22, money: true },
        { key: "DescontoFolha", label: "Desconto em folha", width: 22, money: true },
        { key: "ValorVasconEstimado", label: "Valor Vascon estimado", width: 24, money: true }
      ],
      linhas: linhas.map(l => ({
        ...l,
        ValorUnitario: 0,
        DescontoFolha: 0,
        ValorVasconEstimado: 0
      }))
    });
  },

  _exportarExcelHomy({ nomeArquivo, titulo, periodo, resumo, colunas, linhas }) {
    if (!window.XLSX) {
      AdminUtils.toast("Biblioteca XLSX não carregada.", "error");
      return;
    }

    const aoa = [];

    aoa.push([titulo]);
    aoa.push([periodo]);
    aoa.push([]);

    (resumo || []).forEach(item => {
      aoa.push([item[0], item[1]]);
    });

    aoa.push([]);
    aoa.push(colunas.map(c => c.label));

    linhas.forEach(linha => {
      aoa.push(colunas.map(c => {
        const valor = linha[c.key];
        return valor === null || valor === undefined ? "" : valor;
      }));
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    const totalCols = colunas.length;
    const headerRow = aoa.findIndex(row =>
      row.length &&
      row[0] === colunas[0].label
    );

    ws["!merges"] = [
      {
        s: { r: 0, c: 0 },
        e: { r: 0, c: totalCols - 1 }
      },
      {
        s: { r: 1, c: 0 },
        e: { r: 1, c: totalCols - 1 }
      }
    ];

    ws["!cols"] = colunas.map(c => ({ wch: c.width || 18 }));

    ws["!freeze"] = {
      xSplit: 0,
      ySplit: headerRow + 1,
      topLeftCell: `A${headerRow + 2}`,
      activePane: "bottomLeft",
      state: "frozen"
    };

    ws["!autofilter"] = {
      ref: XLSX.utils.encode_range({
        s: { r: headerRow, c: 0 },
        e: { r: Math.max(headerRow, aoa.length - 1), c: totalCols - 1 }
      })
    };

    const range = XLSX.utils.decode_range(ws["!ref"]);

    const border = {
      top: { style: "thin", color: { rgb: "D9E2F3" } },
      bottom: { style: "thin", color: { rgb: "D9E2F3" } },
      left: { style: "thin", color: { rgb: "D9E2F3" } },
      right: { style: "thin", color: { rgb: "D9E2F3" } }
    };

    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= Math.max(range.e.c, totalCols - 1); C++) {
        const ref = XLSX.utils.encode_cell({ r: R, c: C });

        if (!ws[ref]) {
          ws[ref] = { t: "s", v: "" };
        }

        ws[ref].s = {
          font: {
            name: "Calibri",
            sz: 11,
            color: { rgb: "000000" }
          },
          alignment: {
            vertical: "center",
            horizontal: "left",
            wrapText: true
          },
          border
        };

        if (R === 0) {
          ws[ref].s = {
            font: {
              name: "Calibri",
              sz: 14,
              bold: true,
              color: { rgb: "FFFFFF" }
            },
            fill: {
              fgColor: { rgb: "071A33" }
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            }
          };
        }

        if (R === 1) {
          ws[ref].s = {
            font: {
              name: "Calibri",
              sz: 11,
              bold: true,
              color: { rgb: "FFFFFF" }
            },
            fill: {
              fgColor: { rgb: "C9281D" }
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            }
          };
        }

        if (R === headerRow) {
          ws[ref].s = {
            font: {
              name: "Calibri",
              sz: 11,
              bold: true,
              color: { rgb: "FFFFFF" }
            },
            fill: {
              fgColor: { rgb: "071A33" }
            },
            alignment: {
              vertical: "center",
              horizontal: "center",
              wrapText: true
            },
            border
          };
        }

        if (R > headerRow && C >= 0) {
          const col = colunas[C];

          if (col?.money) {
            ws[ref].t = "n";
            ws[ref].z = '"R$" #,##0.00';
            ws[ref].s.alignment = {
              vertical: "center",
              horizontal: "right"
            };
          }

          if (typeof ws[ref].v === "number" && !col?.money) {
            ws[ref].s.alignment = {
              vertical: "center",
              horizontal: "center"
            };
          }
        }
      }
    }

    ws["!rows"] = aoa.map((_, idx) => {
      if (idx === 0) return { hpt: 22 };
      if (idx === 1) return { hpt: 18 };
      if (idx === headerRow) return { hpt: 20 };
      return { hpt: 18 };
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Relatorio");

    XLSX.writeFile(wb, nomeArquivo, {
      bookType: "xlsx",
      cellStyles: true
    });

    AdminUtils.toast("Excel exportado com sucesso.", "success");
  },

  _periodoTexto() {
    return `Período: ${this._dataBR(this._periodo.ini)} a ${this._dataBR(this._periodo.fim)}`;
  },

  _nomeArquivo(tipo) {
    return `relatorio-${tipo}-${this._periodo.ini}-${this._periodo.fim}.xlsx`;
  },

  _dataBR(dataIso) {
    if (!dataIso) return "";
    const [a, m, d] = String(dataIso).slice(0, 10).split("-");
    if (!a || !m || !d) return dataIso;
    return `${d}/${m}/${a}`;
  },

  _hojeBR() {
    const d = new Date();
    return d.toLocaleDateString("pt-BR");
  },

  _limparTexto(valor) {
    const texto = String(valor ?? "").trim();
    if (!texto || texto === "-" || texto === "—") return "";
    return texto;
  }
};
