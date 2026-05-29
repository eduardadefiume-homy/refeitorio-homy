// ============================================================
// admin-relatorios.js
// Versão consolidada final
// - Tipo de relatório no próprio HTML do painel
// - Exportação Excel .xlsx real com layout Homy
// - Sem arquivos force/layout/remendos
// ============================================================

(function () {
  const state = { pedidos: [], colaboradores: [], valores: [], ultimoResultado: null };

  const tipos = {
    "resumo-dia": "Resumo por dia",
    "centro-custo": "Valor total por centro de custo",
    "colaborador": "Quantidade por colaborador"
  };

  function normalizar(valor) {
    return String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  }

  function estaNaTelaRelatorios() {
    const modulo = document.getElementById("mod-relatorios");
    if (modulo && modulo.classList.contains("active")) return true;

    const ativo = document.querySelector(".nav-item.active, .module.active");
    return normalizar(ativo?.innerText || "").includes("relatorios");
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

    const dt = new Date(data);
    if (Number.isNaN(dt.getTime())) return "";
    return dt.toLocaleDateString("pt-BR");
  }

  function toInputDate(data) {
    const d = new Date(data);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }

  function diaSemanaCompleto(data) {
    return ["Domingo", "Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado"][new Date(data).getDay()];
  }

  function slug(valor) {
    return normalizar(valor).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  }

  function valorTela(valor) {
    if (typeof valor === "number") return String(valor).replace(".", ",");
    return valor ?? "";
  }

  function borda(color = "FFCBD5E1") {
    return {
      top: { style: "thin", color: { argb: color } },
      left: { style: "thin", color: { argb: color } },
      bottom: { style: "thin", color: { argb: color } },
      right: { style: "thin", color: { argb: color } }
    };
  }

  function setTexto(ids, valor) {
    for (const id of ids) {
      const el = document.getElementById(id);
      if (el) {
        el.textContent = valor;
        return;
      }
    }
  }

  function inicializarDatas() {
    const inicio = document.getElementById("relDataInicio");
    const fim = document.getElementById("relDataFim");
    if (!inicio || !fim) return;

    const hoje = new Date();
    const diaSemana = hoje.getDay() || 7;
    const segunda = new Date(hoje);
    segunda.setDate(hoje.getDate() - diaSemana + 1);
    const sexta = new Date(segunda);
    sexta.setDate(segunda.getDate() + 4);

    if (!inicio.value) inicio.value = toInputDate(segunda);
    if (!fim.value) fim.value = toInputDate(sexta);
  }

  function getFiltros() {
    inicializarDatas();

    const dataInicio = document.getElementById("relDataInicio")?.value || toInputDate(new Date());
    const dataFim = document.getElementById("relDataFim")?.value || dataInicio;
    const tipo = document.getElementById("tipoRelatorio")?.value || "resumo-dia";
    const nfVascon = Number(String(document.getElementById("relNfVascon")?.value || "0").replace(",", ".")) || 0;

    return { dataInicio, dataFim, tipo, nfVascon };
  }

  async function safe(fn) {
    try {
      return await fn();
    } catch (erro) {
      console.warn("Falha ao carregar dados do relatório:", erro);
      return [];
    }
  }

  async function carregarDados() {
    if (!window.SP) throw new Error("SP não encontrado. Verifique se sharepoint.js foi carregado.");

    state.pedidos = await safe(() => SP.getItems("Pedidos"));
    state.colaboradores = await safe(() => SP.getTodosColaboradores ? SP.getTodosColaboradores() : SP.getColaboradores());
    state.valores = await safe(() => SP.getValoresRefeicao ? SP.getValoresRefeicao() : []);
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

    const semana = String(pedido.Semana_id || "");
    const dia = String(pedido.Dia || "");
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

  function filtrarPedidos(dataInicio, dataFim) {
    const inicio = new Date(`${dataInicio}T00:00:00`);
    const fim = new Date(`${dataFim}T23:59:59`);

    return (state.pedidos || []).filter(p => {
      const status = normalizar(p.Status || "");
      if (["cancelado", "bloqueado", "afastado", "ferias", "nao vai almocar", "não vai almoçar"].includes(status)) return false;

      const d = dataPedido(p);
      return d && d >= inicio && d <= fim;
    });
  }

  function obterValoresPeriodo(dataInicio, dataFim) {
    const inicio = new Date(`${dataInicio}T00:00:00`);
    const fim = new Date(`${dataFim}T23:59:59`);

    const ativos = (state.valores || []).filter(v => {
      const ativo = SP.isTrue ? SP.isTrue(v.Ativo) : String(v.Ativo).toLowerCase() !== "false";
      const vi = v.Data_Inicio ? new Date(v.Data_Inicio) : null;
      const vf = v.Data_Fim ? new Date(v.Data_Fim) : null;

      if (!ativo) return false;
      if (!vi || !vf) return true;
      return vi <= fim && vf >= inicio;
    });

    const valor = ativos[0] || state.valores[0] || {};

    return {
      valorVascon: Number(valor.Valor_Vascon || 0),
      valorDesconto: Number(valor.Valor_Desconto_Funcionario || 0),
      titulo: valor.Title || ""
    };
  }

  function contarOpcoes(pedidos) {
    const total = { Principal: 0, Light: 0, Carne: 0, Massa: 0, Lanche: 0 };

    pedidos.forEach(p => {
      const opcao = normalizar(p.Opcao || p.Nome_Prato || "");
      if (opcao.includes("principal")) total.Principal++;
      else if (opcao.includes("light") || opcao.includes("ligth")) total.Light++;
      else if (opcao.includes("carne")) total.Carne++;
      else if (opcao.includes("massa")) total.Massa++;
      else if (opcao.includes("lanche")) total.Lanche++;
    });

    return total;
  }

  function centroCustoColaborador(pedido) {
    const id = String(pedido.Colaborador_id || "");
    const nome = normalizar(pedido.Colaborador_nome || "");

    const colaborador = (state.colaboradores || []).find(c =>
      String(c.id) === id ||
      String(c.ID) === id ||
      normalizar(c.Nome || c.Title || "") === nome
    );

    return colaborador?.Centro_Custo || "";
  }

  function montarResumoDia(dataInicio, dataFim, pedidos) {
    const linhas = [];
    const inicio = new Date(`${dataInicio}T00:00:00`);
    const fim = new Date(`${dataFim}T00:00:00`);

    for (let d = new Date(inicio); d <= fim; d.setDate(d.getDate() + 1)) {
      const chave = toInputDate(d);
      const pedidosDia = pedidos.filter(p => {
        const dp = dataPedido(p);
        return dp && toInputDate(dp) === chave;
      });

      const op = contarOpcoes(pedidosDia);

      linhas.push({
        "Dia": diaSemanaCompleto(d),
        "Data": dataBR(d),
        "Principal": op.Principal,
        "Light": op.Light,
        "Carne": op.Carne,
        "Massa": op.Massa,
        "Lanche": op.Lanche,
        "Total": pedidosDia.length
      });
    }

    return linhas;
  }

  function montarCentroCusto(pedidos, valores) {
    const mapa = {};

    pedidos.forEach(p => {
      const centro = p.Centro_Custo || centroCustoColaborador(p) || "Sem centro de custo";

      if (!mapa[centro]) {
        mapa[centro] = {
          "Centro de Custo": centro,
          "Quantidade": 0,
          "Valor Vascon": 0,
          "Desconto Funcionários": 0,
          "Rateio NF": 0
        };
      }

      mapa[centro]["Quantidade"]++;
      mapa[centro]["Valor Vascon"] += valores.valorVascon;
      mapa[centro]["Desconto Funcionários"] += valores.valorDesconto;
      mapa[centro]["Rateio NF"] += valores.valorVascon;
    });

    return Object.values(mapa);
  }

  function montarColaborador(pedidos, valores) {
    const mapa = {};

    pedidos.forEach(p => {
      const nome = p.Colaborador_nome || p.Colaborador || "Sem nome";
      const centro = p.Centro_Custo || centroCustoColaborador(p) || "Sem centro de custo";
      const chave = `${nome}|${centro}`;

      if (!mapa[chave]) {
        mapa[chave] = {
          "Colaborador": nome,
          "Centro de Custo": centro,
          "Quantidade": 0,
          "Valor de cada refeição": valores.valorDesconto,
          "Desconto em folha": 0
        };
      }

      mapa[chave]["Quantidade"]++;
      mapa[chave]["Desconto em folha"] += valores.valorDesconto;
    });

    return Object.values(mapa).sort((a, b) => String(a.Colaborador).localeCompare(String(b.Colaborador)));
  }

  function linhasDoTipo(resultado) {
    if (resultado.filtros.tipo === "centro-custo") return resultado.centroCusto;
    if (resultado.filtros.tipo === "colaborador") return resultado.colaborador;
    return resultado.resumoDia;
  }

  function tituloDoTipo(tipo) {
    return tipos[tipo] || "Relatório";
  }

  function atualizarCards(resultado) {
    const op = resultado.totais.opcoes;

    setTexto(["rel-total", "relTotalRefeicoes", "statRelRefeicoes"], resultado.totais.total);
    setTexto(["rel-vascon", "relCustoVascon", "statRelVascon"], moeda(resultado.totais.valorVascon));
    setTexto(["rel-desconto", "relDescontoFuncionarios", "statRelDesconto"], moeda(resultado.totais.desconto));
    setTexto(["rel-diferenca", "relDiferencaNf", "statRelDiferenca"], resultado.totais.diferencaNf === null ? "-" : moeda(resultado.totais.diferencaNf));

    setTexto(["rel-principal"], op.Principal);
    setTexto(["rel-light", "relLigth"], op.Light);
    setTexto(["rel-carne"], op.Carne);
    setTexto(["rel-massa"], op.Massa);
    setTexto(["rel-lanche"], op.Lanche);
  }

  function renderTabela(resultado) {
    const linhas = linhasDoTipo(resultado);
    const colunas = linhas[0] ? Object.keys(linhas[0]) : (resultado.filtros.tipo === "resumo-dia"
      ? ["Dia", "Data", "Principal", "Light", "Carne", "Massa", "Lanche", "Total"]
      : resultado.filtros.tipo === "centro-custo"
        ? ["Centro de Custo", "Quantidade", "Valor Vascon", "Desconto Funcionários", "Rateio NF"]
        : ["Colaborador", "Centro de Custo", "Quantidade", "Valor de cada refeição", "Desconto em folha"]);

    const titulo = tituloDoTipo(resultado.filtros.tipo);
    const tituloEl = document.getElementById("relTituloTabela");
    if (tituloEl) tituloEl.textContent = `📌 ${titulo}`;

    const thead = document.getElementById("relTableHead");
    const tbody = document.getElementById("relTable");
    if (!thead || !tbody) return;

    thead.innerHTML = `<tr>${colunas.map(c => `<th>${c}</th>`).join("")}</tr>`;

    if (!linhas.length) {
      tbody.innerHTML = `<tr><td colspan="${colunas.length}" style="text-align:center;color:rgba(143,170,210,0.4);padding:2rem">Sem dados no período.</td></tr>`;
      return;
    }

    tbody.innerHTML = linhas.map(linha => `
      <tr>${colunas.map(c => `<td>${valorTela(linha[c])}</td>`).join("")}</tr>
    `).join("");
  }

  async function gerar() {
    if (!estaNaTelaRelatorios()) return;

    try {
      await carregarDados();

      const filtros = getFiltros();
      const pedidos = filtrarPedidos(filtros.dataInicio, filtros.dataFim);
      const valores = obterValoresPeriodo(filtros.dataInicio, filtros.dataFim);

      const resultado = {
        filtros,
        pedidos,
        valoresPeriodo: valores,
        resumoDia: montarResumoDia(filtros.dataInicio, filtros.dataFim, pedidos),
        centroCusto: montarCentroCusto(pedidos, valores),
        colaborador: montarColaborador(pedidos, valores)
      };

      const total = pedidos.length;
      resultado.totais = {
        total,
        valorVascon: total * valores.valorVascon,
        desconto: total * valores.valorDesconto,
        diferencaNf: filtros.nfVascon ? filtros.nfVascon - (total * valores.valorVascon) : null,
        opcoes: contarOpcoes(pedidos)
      };

      state.ultimoResultado = resultado;
      atualizarCards(resultado);
      renderTabela(resultado);
    } catch (erro) {
      console.error("Erro ao gerar relatório:", erro);
      alert(`Erro ao gerar relatório: ${erro.message || erro}`);
    }
  }

  async function carregarExcelJS() {
    if (window.ExcelJS) return;

    await new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js";
      script.onload = resolve;
      script.onerror = () => reject(new Error("Não foi possível carregar ExcelJS."));
      document.head.appendChild(script);
    });
  }

  function aplicarEstiloResumo(ws) {
    for (let row = 5; row <= 9; row++) {
      ws.getCell(row, 1).font = { bold: true };
      ws.getCell(row, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F1FF" } };
      ws.getCell(row, 1).border = borda();
      ws.getCell(row, 2).border = borda();
      if (row > 5) ws.getCell(row, 2).numFmt = '"R$" #,##0.00';
    }
  }

  async function exportarExcel() {
    if (!state.ultimoResultado) await gerar();
    const resultado = state.ultimoResultado;
    if (!resultado) return;

    await carregarExcelJS();

    const titulo = tituloDoTipo(resultado.filtros.tipo);
    const linhas = linhasDoTipo(resultado);
    const colunas = linhas[0] ? Object.keys(linhas[0]) : (resultado.filtros.tipo === "resumo-dia"
      ? ["Dia", "Data", "Principal", "Light", "Carne", "Massa", "Lanche", "Total"]
      : resultado.filtros.tipo === "centro-custo"
        ? ["Centro de Custo", "Quantidade", "Valor Vascon", "Desconto Funcionários", "Rateio NF"]
        : ["Colaborador", "Centro de Custo", "Quantidade", "Valor de cada refeição", "Desconto em folha"]);

    const wb = new ExcelJS.Workbook();
    wb.creator = "Refeitório Homy";
    wb.created = new Date();

    const ws = wb.addWorksheet(titulo.substring(0, 31));
    const totalCols = Math.max(colunas.length, 6);

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
    ws.addRow(["Total de refeições", resultado.totais.total]);
    ws.addRow(["Custo Vascon estimado", resultado.totais.valorVascon]);
    ws.addRow(["Desconto funcionários", resultado.totais.desconto]);
    ws.addRow(["Valor unitário Vascon", resultado.valoresPeriodo.valorVascon]);
    ws.addRow(["Valor unitário descontado", resultado.valoresPeriodo.valorDesconto]);
    ws.addRow([]);
    aplicarEstiloResumo(ws);

    const header = ws.addRow(colunas);
    header.eachCell(cell => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1F3C" } };
      cell.alignment = { horizontal: "center" };
      cell.border = borda();
    });

    const colunasMoeda = ["Valor Vascon", "Desconto Funcionários", "Rateio NF", "Valor de cada refeição", "Desconto em folha"];

    if (linhas.length) {
      linhas.forEach((linha, idx) => {
        const row = ws.addRow(colunas.map(c => linha[c]));
        row.eachCell((cell, colNumber) => {
          const colName = colunas[colNumber - 1];
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? "FFF3F6FB" : "FFFFFFFF" } };
          cell.border = borda("FFD9E2F3");
          if (colunasMoeda.includes(colName)) cell.numFmt = '"R$" #,##0.00';
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

    ws.views = [{ state: "frozen", ySplit: 11 }];

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const nome = `relatorio-refeitorio-${slug(titulo)}-${resultado.filtros.dataInicio}-a-${resultado.filtros.dataFim}.xlsx`;
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = nome;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  function inicializar() {
    inicializarDatas();

    document.getElementById("btnGerarRelatorio")?.addEventListener("click", gerar);
    document.getElementById("btnExportarRelatorio")?.addEventListener("click", exportarExcel);
    document.getElementById("tipoRelatorio")?.addEventListener("change", gerar);
    document.getElementById("relDataInicio")?.addEventListener("change", gerar);
    document.getElementById("relDataFim")?.addEventListener("change", gerar);

    document.addEventListener("click", event => {
      const alvo = event.target.closest(".nav-item, [data-module], [onclick]");
      if (!alvo) return;

      const texto = normalizar(`${alvo.innerText || ""} ${alvo.dataset?.module || ""} ${alvo.getAttribute("onclick") || ""}`);
      if (texto.includes("relatorios")) {
        setTimeout(() => {
          inicializarDatas();
          gerar();
        }, 350);
      }
    });

    if (estaNaTelaRelatorios()) gerar();
  }

  window.AdminRelatorios = { gerar, exportarExcel, inicializar };
  window.gerarRelatorio = gerar;
  window.exportarCSV = exportarExcel;
  window.exportarRelatorioCSV = exportarExcel;
  window.exportarExcel = exportarExcel;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", inicializar);
  } else {
    inicializar();
  }
})();
