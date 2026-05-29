// ============================================================
// admin-relatorios.js
// DEFINITIVO: sobrescreve loadRelatorios do admin-extensions
// e controla a tela Relatórios inteira.
// ============================================================

(function () {
  const DIAS = ["segunda", "terca", "quarta", "quinta", "sexta", "sabado", "domingo"];
  const DIA_LABEL = {
    segunda: "Segunda-feira",
    terca: "Terça-feira",
    quarta: "Quarta-feira",
    quinta: "Quinta-feira",
    sexta: "Sexta-feira",
    sabado: "Sábado",
    domingo: "Domingo"
  };

  const TIPOS = {
    "resumo-dia": "Resumo por dia",
    "centro-custo": "Valor total por centro de custo",
    "colaborador": "Quantidade por colaborador"
  };

  const state = {
    pedidos: [],
    colaboradores: [],
    valores: [],
    resultado: null
  };

  function $(id) {
    return document.getElementById(id);
  }

  function normalizar(valor) {
    return String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  }

  function moeda(valor) {
    return Number(valor || 0).toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL"
    });
  }

  function dataBR(data) {
    if (!data) return "";

    if (typeof data === "string" && /^\d{4}-\d{2}-\d{2}$/.test(data)) {
      const [ano, mes, dia] = data.split("-");
      return `${dia}/${mes}/${ano}`;
    }

    const d = new Date(data);
    if (Number.isNaN(d.getTime())) return "";
    return d.toLocaleDateString("pt-BR");
  }

  function toInputDate(data) {
    const d = new Date(data);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }

  function slug(valor) {
    return normalizar(valor).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  }

  function isTrue(valor) {
    if (window.SP && typeof SP.isTrue === "function") return SP.isTrue(valor);
    return ["true", "sim", "1", "yes"].includes(normalizar(valor));
  }

  function semanaAtualFallback() {
    if (typeof window.getSemanaIdAtualSelecionada === "function") return window.getSemanaIdAtualSelecionada();
    if (typeof window.getSemanaId === "function") return window.getSemanaId();

    const d = new Date();
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() + 3 - (d.getDay() + 6) % 7);
    const w1 = new Date(d.getFullYear(), 0, 4);
    const wn = 1 + Math.round(((d - w1) / 86400000 - 3 + (w1.getDay() + 6) % 7) / 7);
    return d.getFullYear() + "-W" + String(wn).padStart(2, "0");
  }

  function inicioSemanaISO(ano, semana) {
    const jan4 = new Date(ano, 0, 4);
    const dia = jan4.getDay() || 7;
    const segunda = new Date(jan4);
    segunda.setDate(jan4.getDate() - dia + 1 + (semana - 1) * 7);
    return segunda;
  }

  function dataPedido(pedido) {
    if (pedido.Data_Hora) {
      const d = new Date(pedido.Data_Hora);
      if (!Number.isNaN(d.getTime())) return d;
    }

    if (pedido.Data) {
      const d = new Date(pedido.Data);
      if (!Number.isNaN(d.getTime())) return d;
    }

    const semana = String(pedido.Semana_id || pedido.Semana || pedido.semana || "");
    const dia = String(pedido.Dia || pedido.dia || "");
    const match = semana.match(/(\d{4})-?W?(\d{1,2})/i);

    if (match && dia) {
      const segunda = inicioSemanaISO(Number(match[1]), Number(match[2]));
      const mapa = {
        segunda: 0,
        terca: 1,
        terça: 1,
        quarta: 2,
        quinta: 3,
        sexta: 4,
        sabado: 5,
        sábado: 5,
        domingo: 6
      };

      const idx = mapa[normalizar(dia)];

      if (idx !== undefined) {
        const d = new Date(segunda);
        d.setDate(d.getDate() + idx);
        return d;
      }
    }

    return null;
  }

  function diaChave(data) {
    const mapa = ["domingo", "segunda", "terca", "quarta", "quinta", "sexta", "sabado"];
    return mapa[new Date(data).getDay()];
  }

  function getStatusPedido(pedido) {
    return normalizar(pedido.Status || pedido.status || "");
  }

  function pedidoConta(pedido) {
    const status = getStatusPedido(pedido);
    if (["cancelado", "bloqueado", "afastado", "ferias", "férias", "nao vai almocar", "não vai almoçar"].includes(status)) {
      return false;
    }

    if (pedido.Confirmado === false || pedido.confirmado === false) return false;
    return true;
  }

  function getOpcao(pedido) {
    const raw = normalizar(pedido.Opcao || pedido.opcao || pedido.Nome_Prato || pedido.nomePrato || "");

    if (raw.includes("principal")) return "principal";
    if (raw.includes("light") || raw.includes("ligth")) return "light";
    if (raw.includes("carne")) return "carne";
    if (raw.includes("massa")) return "massa";
    if (raw.includes("lanche")) return "lanche";

    return raw || "principal";
  }

  function getNomeColaborador(pedido) {
    return pedido.Colaborador_nome ||
      pedido.colaborador_nome ||
      pedido.Colaborador ||
      pedido.colaborador ||
      pedido.Title ||
      "Sem nome";
  }

  function getColaboradorId(pedido) {
    return String(
      pedido.Colaborador_id ||
      pedido.colaborador_id ||
      pedido.ColaboradorId ||
      pedido.colaboradorId ||
      ""
    );
  }

  function getCentroCusto(pedido) {
    const direto = pedido.Centro_Custo || pedido.centro_custo || pedido.CentroCusto || pedido.centroCusto;
    if (direto) return direto;

    const id = getColaboradorId(pedido);
    const nome = normalizar(getNomeColaborador(pedido));

    const colaborador = (state.colaboradores || []).find(c =>
      String(c.id) === id ||
      String(c.ID) === id ||
      normalizar(c.Nome || c.Title || "") === nome
    );

    return colaborador?.Centro_Custo || colaborador?.centro_custo || "Sem centro de custo";
  }

  function valorNumero(valor) {
    return Number(String(valor || "0").replace(",", ".")) || 0;
  }

  async function safe(fn, fallback = []) {
    try {
      return await fn();
    } catch (erro) {
      console.warn("Falha no relatório:", erro);
      return fallback;
    }
  }

  async function carregarDados(semanaId) {
    if (!window.SP) throw new Error("SP não carregado.");

    state.colaboradores = await safe(() =>
      SP.getTodosColaboradores ? SP.getTodosColaboradores() : SP.getColaboradores(), []
    );

    if (SP.getItems) {
      state.pedidos = await safe(() => SP.getItems("Pedidos"), []);
    } else {
      state.pedidos = await safe(() => SP.getPedidos(semanaId || semanaAtualFallback()), []);
    }

    if (SP.getValoresRefeicao) {
      state.valores = await safe(() => SP.getValoresRefeicao(), []);
    } else if (SP.getItems) {
      state.valores = await safe(() => SP.getItems("Valores de Refeição"), []);
    } else {
      state.valores = [];
    }
  }

  function obterValorPeriodo(dataInicio, dataFim) {
    const ini = new Date(`${dataInicio}T00:00:00`);
    const fim = new Date(`${dataFim}T23:59:59`);

    const ativos = (state.valores || []).filter(v => {
      const ativo = v.Ativo === undefined ? true : isTrue(v.Ativo);
      const vi = v.Data_Inicio ? new Date(v.Data_Inicio) : null;
      const vf = v.Data_Fim ? new Date(v.Data_Fim) : null;

      if (!ativo) return false;
      if (!vi || !vf) return true;

      return vi <= fim && vf >= ini;
    });

    const v = ativos[0] || state.valores[0] || {};

    return {
      valorVascon: valorNumero(v.Valor_Vascon || v.valorVascon),
      valorDesconto: valorNumero(v.Valor_Desconto_Funcionario || v.valorDescontoFuncionario || v.valorDesconto)
    };
  }

  function criarTelaRelatorios() {
    const mod = $("mod-relatorios");
    if (!mod) return;

    mod.dataset.advanced = "final";

    mod.innerHTML = `
      <div class="section-header">
        <div>
          <div class="section-title">📈 Relatórios Gerenciais</div>
          <div style="font-size:0.78rem;color:rgba(143,170,210,0.58);margin-top:0.25rem">
            Quantidade, centro de custo, colaborador, valores e comparação com NF Vascon.
          </div>
        </div>
        <button class="btn-secondary" type="button" id="btnExportarRelatorioFinal">📥 Exportar Excel</button>
      </div>

      <div class="form-grid" style="margin-bottom:1rem;grid-template-columns:minmax(180px,1fr) minmax(180px,1fr) minmax(210px,1fr) minmax(240px,1fr) minmax(200px,260px);align-items:end">
        <div class="form-group">
          <label class="form-label">Data inicial</label>
          <input class="form-input" type="date" id="relDataInicio">
        </div>

        <div class="form-group">
          <label class="form-label">Data final</label>
          <input class="form-input" type="date" id="relDataFim">
        </div>

        <div class="form-group">
          <label class="form-label">NF Vascon recebida</label>
          <input class="form-input" type="number" step="0.01" id="relValorNF" placeholder="Ex.: 1234,56">
        </div>

        <div class="form-group">
          <label class="form-label">Tipo de relatório</label>
          <select class="form-select" id="tipoRelatorio">
            <option value="resumo-dia">Resumo por dia</option>
            <option value="centro-custo">Valor total por centro de custo</option>
            <option value="colaborador">Quantidade por colaborador</option>
          </select>
        </div>

        <div class="form-group" style="margin:0;align-self:end">
          <button class="btn-primary" type="button" id="btnGerarRelatorioFinal" style="width:100%;height:42px;justify-content:center">🔎 Gerar relatório</button>
        </div>
      </div>

      <div class="stats-grid">
        <div class="stat-card"><div class="stat-icon">🍽️</div><div class="stat-value" id="relTotalGeral">0</div><div class="stat-label">Refeições no período</div></div>
        <div class="stat-card"><div class="stat-icon">🏢</div><div class="stat-value" id="relCustoVascon">R$ 0,00</div><div class="stat-label">Custo Vascon estimado</div></div>
        <div class="stat-card"><div class="stat-icon">👤</div><div class="stat-value" id="relDescontoFolha">R$ 0,00</div><div class="stat-label">Desconto funcionários</div></div>
        <div class="stat-card"><div class="stat-icon">🧾</div><div class="stat-value" id="relDiferencaNF">—</div><div class="stat-label">Diferença NF x calculado</div></div>
      </div>

      <div class="stats-grid">
        <div class="stat-card"><div class="stat-icon">🍗</div><div class="stat-value" id="rel-principal">0</div><div class="stat-label">Principal</div></div>
        <div class="stat-card"><div class="stat-icon">🥗</div><div class="stat-value" id="rel-light">0</div><div class="stat-label">Light</div></div>
        <div class="stat-card"><div class="stat-icon">🥩</div><div class="stat-value" id="rel-carne">0</div><div class="stat-label">Carne</div></div>
        <div class="stat-card"><div class="stat-icon">🍝</div><div class="stat-value" id="rel-massa">0</div><div class="stat-label">Massa</div></div>
        <div class="stat-card"><div class="stat-icon">🍔</div><div class="stat-value" id="rel-lanche">0</div><div class="stat-label">Lanche</div></div>
      </div>

      <div style="margin-top:1rem">
        <div class="section-title" id="relTituloTabela" style="font-size:1rem;margin-bottom:0.8rem">Resumo por dia</div>
        <div class="table-wrap">
          <table class="table">
            <thead id="relHead"></thead>
            <tbody id="relTableFinal"></tbody>
          </table>
        </div>
      </div>
    `;

    const hoje = new Date();
    const primeiroDia = new Date(hoje.getFullYear(), hoje.getMonth(), 1);

    $("relDataInicio").value = toInputDate(primeiroDia);
    $("relDataFim").value = toInputDate(hoje);

    $("btnGerarRelatorioFinal").onclick = () => gerarRelatorioFinal(semanaAtualFallback());
    $("btnExportarRelatorioFinal").onclick = exportarRelatorioExcelFinal;
    $("tipoRelatorio").onchange = () => gerarRelatorioFinal(semanaAtualFallback());
  }

  function getFiltros() {
    return {
      dataInicio: $("relDataInicio")?.value || toInputDate(new Date()),
      dataFim: $("relDataFim")?.value || toInputDate(new Date()),
      nf: valorNumero($("relValorNF")?.value || 0),
      tipo: $("tipoRelatorio")?.value || "resumo-dia"
    };
  }

  function filtrarPedidos(filtros) {
    const ini = new Date(`${filtros.dataInicio}T00:00:00`);
    const fim = new Date(`${filtros.dataFim}T23:59:59`);

    return (state.pedidos || []).filter(p => {
      if (!pedidoConta(p)) return false;
      const data = dataPedido(p);
      return data && data >= ini && data <= fim;
    });
  }

  function montarResumoDia(filtros, pedidos) {
    const linhas = [];
    const ini = new Date(`${filtros.dataInicio}T00:00:00`);
    const fim = new Date(`${filtros.dataFim}T00:00:00`);

    for (let d = new Date(ini); d <= fim; d.setDate(d.getDate() + 1)) {
      const chave = toInputDate(d);
      const pedidosDia = pedidos.filter(p => {
        const data = dataPedido(p);
        return data && toInputDate(data) === chave;
      });

      const cont = { principal: 0, light: 0, carne: 0, massa: 0, lanche: 0 };

      pedidosDia.forEach(p => {
        const op = getOpcao(p);
        if (cont[op] !== undefined) cont[op]++;
      });

      linhas.push({
        "Dia": DIA_LABEL[diaChave(d)] || diaSemanaCompleto(d),
        "Data": dataBR(d),
        "Principal": cont.principal,
        "Light": cont.light,
        "Carne": cont.carne,
        "Massa": cont.massa,
        "Lanche": cont.lanche,
        "Total": pedidosDia.length
      });
    }

    return linhas;
  }

  function montarCentroCusto(pedidos, valores, nf) {
    const mapa = {};

    pedidos.forEach(p => {
      const cc = getCentroCusto(p);

      if (!mapa[cc]) {
        mapa[cc] = {
          "Centro de custo": cc,
          "Quantidade": 0,
          "Valor Vascon": 0,
          "Desconto funcionários": 0,
          "Rateio NF": 0
        };
      }

      mapa[cc]["Quantidade"]++;
      mapa[cc]["Valor Vascon"] += valores.valorVascon;
      mapa[cc]["Desconto funcionários"] += valores.valorDesconto;
    });

    const total = pedidos.length || 1;

    Object.values(mapa).forEach(l => {
      l["Rateio NF"] = nf ? nf * (l["Quantidade"] / total) : 0;
    });

    return Object.values(mapa).sort((a, b) => b.Quantidade - a.Quantidade);
  }

  function montarColaborador(pedidos, valores) {
    const mapa = {};

    pedidos.forEach(p => {
      const nome = getNomeColaborador(p);
      const cc = getCentroCusto(p);
      const key = `${nome}|${cc}`;

      if (!mapa[key]) {
        mapa[key] = {
          "Colaborador": nome,
          "Centro de custo": cc,
          "Quantidade": 0,
          "Valor de cada refeição": valores.valorDesconto,
          "Desconto em folha": 0,
          "Valor Vascon estimado": 0
        };
      }

      mapa[key]["Quantidade"]++;
      mapa[key]["Desconto em folha"] += valores.valorDesconto;
      mapa[key]["Valor Vascon estimado"] += valores.valorVascon;
    });

    return Object.values(mapa).sort((a, b) => String(a.Colaborador).localeCompare(String(b.Colaborador)));
  }

  function atualizarCards(pedidos, valores, filtros) {
    const total = pedidos.length;
    const custo = total * valores.valorVascon;
    const desconto = total * valores.valorDesconto;
    const cont = { principal: 0, light: 0, carne: 0, massa: 0, lanche: 0 };

    pedidos.forEach(p => {
      const op = getOpcao(p);
      if (cont[op] !== undefined) cont[op]++;
    });

    $("relTotalGeral").textContent = total;
    $("relCustoVascon").textContent = moeda(custo);
    $("relDescontoFolha").textContent = moeda(desconto);
    $("relDiferencaNF").textContent = filtros.nf ? moeda(filtros.nf - custo) : "—";

    $("rel-principal").textContent = cont.principal;
    $("rel-light").textContent = cont.light;
    $("rel-carne").textContent = cont.carne;
    $("rel-massa").textContent = cont.massa;
    $("rel-lanche").textContent = cont.lanche;
  }

  function linhasDoTipo(resultado) {
    if (resultado.filtros.tipo === "centro-custo") return resultado.centroCusto;
    if (resultado.filtros.tipo === "colaborador") return resultado.colaborador;
    return resultado.resumoDia;
  }

  function valorTela(valor) {
    if (typeof valor === "number") return String(valor).replace(".", ",");
    return valor ?? "";
  }

  function renderTabela(resultado) {
    const linhas = linhasDoTipo(resultado);
    const colunas = linhas[0] ? Object.keys(linhas[0]) : ["Sem dados"];
    const titulo = TIPOS[resultado.filtros.tipo] || "Relatório";

    $("relTituloTabela").textContent = titulo;
    $("relHead").innerHTML = `<tr>${colunas.map(c => `<th>${c}</th>`).join("")}</tr>`;

    if (!linhas.length) {
      $("relTableFinal").innerHTML = `<tr><td colspan="${colunas.length}" style="text-align:center;padding:2rem;color:rgba(143,170,210,0.4)">Sem dados no período.</td></tr>`;
      return;
    }

    $("relTableFinal").innerHTML = linhas.map(linha => `
      <tr>${colunas.map(c => `<td>${valorTela(linha[c])}</td>`).join("")}</tr>
    `).join("");
  }

  async function gerarRelatorioFinal(semanaId) {
    try {
      await carregarDados(semanaId);

      const filtros = getFiltros();
      const pedidos = filtrarPedidos(filtros);
      const valores = obterValorPeriodo(filtros.dataInicio, filtros.dataFim);

      const resultado = {
        filtros,
        pedidos,
        valores,
        resumoDia: montarResumoDia(filtros, pedidos),
        centroCusto: montarCentroCusto(pedidos, valores, filtros.nf),
        colaborador: montarColaborador(pedidos, valores)
      };

      state.resultado = resultado;

      atualizarCards(pedidos, valores, filtros);
      renderTabela(resultado);
    } catch (erro) {
      console.error("Erro ao gerar relatório final:", erro);
      if (typeof window.toast === "function") {
        toast("Erro ao gerar relatório: " + (erro.message || erro), "error");
      } else {
        alert("Erro ao gerar relatório: " + (erro.message || erro));
      }
    }
  }

  async function carregarExcelJS() {
    if (window.ExcelJS) return;

    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js";
      s.onload = resolve;
      s.onerror = () => reject(new Error("Não foi possível carregar ExcelJS."));
      document.head.appendChild(s);
    });
  }

  function borda(color = "FFCBD5E1") {
    return {
      top: { style: "thin", color: { argb: color } },
      left: { style: "thin", color: { argb: color } },
      bottom: { style: "thin", color: { argb: color } },
      right: { style: "thin", color: { argb: color } }
    };
  }

  async function exportarRelatorioExcelFinal() {
    if (!state.resultado) {
      await gerarRelatorioFinal(semanaAtualFallback());
    }

    const resultado = state.resultado;
    if (!resultado) return;

    await carregarExcelJS();

    const titulo = TIPOS[resultado.filtros.tipo] || "Relatório";
    const linhas = linhasDoTipo(resultado);
    const colunas = linhas[0] ? Object.keys(linhas[0]) : ["Sem dados"];
    const totalCols = Math.max(colunas.length, 6);

    const wb = new ExcelJS.Workbook();
    wb.creator = "Homy Refeitório";

    const ws = wb.addWorksheet(titulo.substring(0, 31));

    ws.mergeCells(1, 1, 1, totalCols);
    ws.getCell(1, 1).value = `RELATÓRIO REFEITÓRIO HOMY — ${titulo.toUpperCase()}`;
    ws.getCell(1, 1).font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    ws.getCell(1, 1).alignment = { horizontal: "center" };
    ws.getCell(1, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1F3C" } };

    ws.mergeCells(2, 1, 2, totalCols);
    ws.getCell(2, 1).value = `Período: ${dataBR(resultado.filtros.dataInicio)} a ${dataBR(resultado.filtros.dataFim)}`;
    ws.getCell(2, 1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    ws.getCell(2, 1).alignment = { horizontal: "center" };
    ws.getCell(2, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC0281C" } };

    ws.mergeCells(3, 1, 3, totalCols);
    ws.getCell(3, 1).value = `Gerado em: ${new Date().toLocaleDateString("pt-BR")}`;
    ws.getCell(3, 1).font = { bold: true, color: { argb: "FF0B1F3C" } };
    ws.getCell(3, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F1FF" } };

    ws.addRow([]);
    ws.addRow(["Total de refeições", resultado.pedidos.length]);
    ws.addRow(["Valor unitário Vascon", resultado.valores.valorVascon]);
    ws.addRow(["Valor unitário descontado", resultado.valores.valorDesconto]);
    ws.addRow([]);

    const header = ws.addRow(colunas);

    header.eachCell(cell => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1F3C" } };
      cell.alignment = { horizontal: "center" };
      cell.border = borda();
    });

    const moneyCols = [
      "Valor Vascon",
      "Desconto funcionários",
      "Rateio NF",
      "Valor de cada refeição",
      "Desconto em folha",
      "Valor Vascon estimado"
    ];

    if (linhas.length) {
      linhas.forEach((linha, idx) => {
        const row = ws.addRow(colunas.map(c => linha[c]));

        row.eachCell((cell, colNumber) => {
          const colName = colunas[colNumber - 1];
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? "FFF3F6FB" : "FFFFFFFF" } };
          cell.border = borda("FFD9E2F3");
          if (moneyCols.includes(colName)) cell.numFmt = '"R$" #,##0.00';
        });
      });
    } else {
      ws.addRow(["Sem dados no período."]);
    }

    ws.columns.forEach(col => {
      let max = 12;
      col.eachCell({ includeEmpty: true }, cell => {
        max = Math.max(max, String(cell.value || "").length + 2);
      });
      col.width = Math.min(max, 38);
    });

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const nomeArquivo = `relatorio-refeitorio-${slug(titulo)}-${resultado.filtros.dataInicio}-a-${resultado.filtros.dataFim}.xlsx`;

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = nomeArquivo;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  // ======================================================================
  // ESTE É O PONTO PRINCIPAL:
  // o admin-extensions.js define loadRelatorios e recria a tela antiga.
  // Aqui nós sobrescrevemos de novo, depois dele.
  // ======================================================================
  window.loadRelatorios = async function (semanaId) {
    criarTelaRelatorios();
    await gerarRelatorioFinal(semanaId || semanaAtualFallback());
  };

  window.loadRelatoriosAvancados = window.loadRelatorios;
  window.exportarRelatorioCSV = exportarRelatorioExcelFinal;
  window.exportarCSV = exportarRelatorioExcelFinal;
  window.exportarExcel = exportarRelatorioExcelFinal;

  window.AdminRelatorios = {
    carregar: window.loadRelatorios,
    gerar: gerarRelatorioFinal,
    exportarExcel: exportarRelatorioExcelFinal
  };

  document.addEventListener("click", event => {
    const item = event.target.closest("[data-module='relatorios'], .nav-item, button, a");
    if (!item) return;

    const texto = normalizar(`${item.innerText || ""} ${item.dataset?.module || ""}`);

    if (texto.includes("relatorios")) {
      setTimeout(() => window.loadRelatorios(semanaAtualFallback()), 120);
    }
  }, true);

  document.addEventListener("DOMContentLoaded", () => {
    setTimeout(() => {
      const ativo = document.querySelector(".nav-item.active");
      if (normalizar(ativo?.dataset?.module || ativo?.innerText || "") === "relatorios") {
        window.loadRelatorios(semanaAtualFallback());
      }
    }, 300);
  });
})();
