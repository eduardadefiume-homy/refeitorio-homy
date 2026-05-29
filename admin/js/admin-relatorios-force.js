// ============================================================
// admin-relatorios-force-v2.js
// Correção forçada: seletor de relatório + exportação Excel real
// ============================================================

(function () {
  const BOX_ID = "tipoRelatorioForceBoxV2";

  const state = {
    pedidos: [],
    colaboradores: [],
    valores: [],
    ultimoResultado: null
  };

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
    const titulo = document.querySelector(".topbar-title, h1, h2");
    if (normalizar(titulo?.innerText || "").includes("relatorios")) return true;

    const ativo = document.querySelector(".nav-item.active, .module.active");
    return normalizar(ativo?.innerText || "").includes("relatorios");
  }

  function containerRelatorios() {
    return document.querySelector(".module.active") ||
      document.querySelector(".content") ||
      document.querySelector("main") ||
      document.body;
  }

  function limparElementosAntigos() {
    document.querySelectorAll("#tipoRelatorioBox, #relatorioSelecionadoPreview, #relatorioSelecionadoPreviewForce").forEach(el => el.remove());
  }

  function prepararTela() {
    if (!estaNaTelaRelatorios()) return;

    limparElementosAntigos();
    forcarBotaoExcel();
    inserirSeletor();
    gerar();
  }

  function forcarBotaoExcel() {
    document.querySelectorAll("button, a").forEach(btn => {
      const txt = normalizar(btn.innerText || "");
      if (txt.includes("exportar csv") || txt.includes("exportar excel") || txt.includes("exportar")) {
        btn.innerHTML = "📥 Exportar Excel";
      }
    });
  }

  function inserirSeletor() {
    if (document.getElementById(BOX_ID)) return;

    const container = containerRelatorios();

    const box = document.createElement("div");
    box.id = BOX_ID;
    box.style.cssText = `
      width:100%;
      margin:14px 0 16px 0;
      padding:14px;
      border:1px solid rgba(60,140,255,.35);
      border-radius:12px;
      background:rgba(10,31,60,.55);
      display:grid;
      grid-template-columns:300px 1fr;
      gap:14px;
      align-items:end;
      box-sizing:border-box;
    `;

    box.innerHTML = `
      <div class="form-group" style="margin:0;">
        <label class="form-label">TIPO DE RELATÓRIO</label>
        <select id="tipoRelatorioForceV2" class="form-select">
          <option value="resumo-dia">Resumo por dia</option>
          <option value="centro-custo">Valor total por centro de custo</option>
          <option value="colaborador">Quantidade por colaborador</option>
        </select>
      </div>
      <div style="font-size:12px;color:#75b7ff;line-height:1.4;">
        Escolha o tipo de relatório, selecione o período e clique em <b>Gerar relatório</b>.
        A exportação baixa somente o relatório selecionado em Excel formatado.
      </div>
    `;

    const botaoGerar = Array.from(container.querySelectorAll("button, a")).find(b =>
      normalizar(b.innerText || "").includes("gerar relatorio")
    );

    const linhaFiltros = botaoGerar?.closest("div") || container.querySelector(".form-grid");

    if (linhaFiltros) linhaFiltros.insertAdjacentElement("afterend", box);
    else container.prepend(box);

    document.getElementById("tipoRelatorioForceV2").addEventListener("change", gerar);
  }

  async function carregarDados() {
    if (!window.SP) throw new Error("SP não encontrado.");

    const safe = async fn => {
      try { return await fn(); }
      catch (e) { console.warn("Falha ao buscar relatório:", e); return []; }
    };

    state.pedidos = await safe(() => SP.getItems("Pedidos"));
    state.colaboradores = await safe(() => SP.getTodosColaboradores ? SP.getTodosColaboradores() : SP.getColaboradores());
    state.valores = await safe(() => SP.getValoresRefeicao ? SP.getValoresRefeicao() : []);
  }

  function filtros() {
    const container = containerRelatorios();
    const datas = Array.from(container.querySelectorAll("input[type='date']"));

    const dataInicio = datas[0]?.value || toInputDate(new Date());
    const dataFim = datas[1]?.value || toInputDate(new Date());
    const tipo = document.getElementById("tipoRelatorioForceV2")?.value || "resumo-dia";

    const nfInput = Array.from(container.querySelectorAll("input")).find(i =>
      normalizar(i.placeholder || "").includes("1234") ||
      normalizar(i.previousElementSibling?.innerText || "").includes("nf vascon")
    );

    const nfVascon = Number(String(nfInput?.value || "0").replace(",", ".")) || 0;

    return { dataInicio, dataFim, tipo, nfVascon };
  }

  async function gerar() {
    if (!estaNaTelaRelatorios()) return;

    try {
      await carregarDados();

      const f = filtros();
      const pedidos = filtrarPedidos(f.dataInicio, f.dataFim);
      const valores = obterValoresPeriodo(f.dataInicio, f.dataFim);

      const resultado = {
        filtros: f,
        pedidos,
        valoresPeriodo: valores,
        resumoDia: montarResumoDia(f.dataInicio, f.dataFim, pedidos),
        centroCusto: montarCentroCusto(pedidos, valores),
        colaborador: montarColaborador(pedidos, valores)
      };

      resultado.totais = montarTotais(resultado, f);
      state.ultimoResultado = resultado;

      atualizarCards(resultado);
      renderPreview(resultado);
    } catch (e) {
      console.error(e);
      alert(`Erro ao gerar relatório: ${e.message || e}`);
    }
  }

  function filtrarPedidos(dataInicio, dataFim) {
    const ini = new Date(`${dataInicio}T00:00:00`);
    const fim = new Date(`${dataFim}T23:59:59`);

    return (state.pedidos || []).filter(p => {
      const status = normalizar(p.Status || "");
      if (["cancelado", "bloqueado", "afastado", "ferias", "nao vai almocar"].includes(status)) return false;

      const d = dataPedido(p);
      return d && d >= ini && d <= fim;
    });
  }

  function dataPedido(p) {
    if (p.Data_Hora) {
      const d = new Date(p.Data_Hora);
      if (!Number.isNaN(d.getTime())) return d;
    }

    if (p.Data) {
      const d = new Date(p.Data);
      if (!Number.isNaN(d.getTime())) return d;
    }

    const semana = String(p.Semana_id || "");
    const dia = String(p.Dia || "");
    const match = semana.match(/(\d{4})-?W?(\d{1,2})/i);

    if (match && dia) {
      const segunda = inicioSemanaISO(Number(match[1]), Number(match[2]));
      const mapa = { segunda: 0, terca: 1, terça: 1, quarta: 2, quinta: 3, sexta: 4, sabado: 5, sábado: 5, domingo: 6 };
      const idx = mapa[normalizar(dia)];

      if (idx !== undefined) {
        const d = new Date(segunda);
        d.setDate(d.getDate() + idx);
        return d;
      }
    }

    return null;
  }

  function inicioSemanaISO(ano, semana) {
    const jan4 = new Date(ano, 0, 4);
    const dia = jan4.getDay() || 7;
    const segunda = new Date(jan4);
    segunda.setDate(jan4.getDate() - dia + 1 + (semana - 1) * 7);
    return segunda;
  }

  function obterValoresPeriodo(dataInicio, dataFim) {
    const ini = new Date(`${dataInicio}T00:00:00`);
    const fim = new Date(`${dataFim}T23:59:59`);

    const ativos = (state.valores || []).filter(v => {
      const ativo = SP.isTrue ? SP.isTrue(v.Ativo) : String(v.Ativo).toLowerCase() !== "false";
      const vi = v.Data_Inicio ? new Date(v.Data_Inicio) : null;
      const vf = v.Data_Fim ? new Date(v.Data_Fim) : null;

      if (!ativo) return false;
      if (!vi || !vf) return true;

      return vi <= fim && vf >= ini;
    });

    const v = ativos[0] || state.valores[0] || {};

    return {
      valorVascon: Number(v.Valor_Vascon || 0),
      valorDesconto: Number(v.Valor_Desconto_Funcionario || 0)
    };
  }

  function montarTotais(resultado, f) {
    const total = resultado.pedidos.length;
    const valorVascon = total * resultado.valoresPeriodo.valorVascon;
    const desconto = total * resultado.valoresPeriodo.valorDesconto;

    return {
      total,
      valorVascon,
      desconto,
      diferencaNf: f.nfVascon ? f.nfVascon - valorVascon : null,
      opcoes: contarOpcoes(resultado.pedidos)
    };
  }

  function montarResumoDia(dataInicio, dataFim, pedidos) {
    const linhas = [];
    const ini = new Date(`${dataInicio}T00:00:00`);
    const fim = new Date(`${dataFim}T00:00:00`);

    for (let d = new Date(ini); d <= fim; d.setDate(d.getDate() + 1)) {
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
      const centro = p.Centro_Custo || centroCustoDoColaborador(p) || "Sem centro de custo";

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
      const centro = p.Centro_Custo || centroCustoDoColaborador(p) || "Sem centro de custo";
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

    return Object.values(mapa);
  }

  function centroCustoDoColaborador(pedido) {
    const id = String(pedido.Colaborador_id || "");
    const nome = normalizar(pedido.Colaborador_nome || "");

    const c = (state.colaboradores || []).find(x =>
      String(x.id) === id ||
      String(x.ID) === id ||
      normalizar(x.Nome || x.Title || "") === nome
    );

    return c?.Centro_Custo || "";
  }

  function contarOpcoes(pedidos) {
    const r = { Principal: 0, Light: 0, Carne: 0, Massa: 0, Lanche: 0 };

    pedidos.forEach(p => {
      const o = normalizar(p.Opcao || p.Nome_Prato || "");
      if (o.includes("principal")) r.Principal++;
      else if (o.includes("light") || o.includes("ligth")) r.Light++;
      else if (o.includes("carne")) r.Carne++;
      else if (o.includes("massa")) r.Massa++;
      else if (o.includes("lanche")) r.Lanche++;
    });

    return r;
  }

  function atualizarCards(resultado) {
    const op = resultado.totais.opcoes;
    setTexto(["relTotalRefeicoes", "statRelRefeicoes"], resultado.totais.total);
    setTexto(["relCustoVascon", "statRelVascon"], moeda(resultado.totais.valorVascon));
    setTexto(["relDescontoFuncionarios", "statRelDesconto"], moeda(resultado.totais.desconto));
    setTexto(["relDiferencaNf", "statRelDiferenca"], resultado.totais.diferencaNf === null ? "-" : moeda(resultado.totais.diferencaNf));
    setTexto(["relPrincipal"], op.Principal);
    setTexto(["relLight", "relLigth"], op.Light);
    setTexto(["relCarne"], op.Carne);
    setTexto(["relMassa"], op.Massa);
    setTexto(["relLanche"], op.Lanche);
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

  function linhasDoTipo(resultado) {
    if (resultado.filtros.tipo === "centro-custo") return resultado.centroCusto;
    if (resultado.filtros.tipo === "colaborador") return resultado.colaborador;
    return resultado.resumoDia;
  }

  function renderPreview(resultado) {
    let box = document.getElementById("relatorioSelecionadoPreviewForce");
    const container = containerRelatorios();

    if (!box) {
      box = document.createElement("div");
      box.id = "relatorioSelecionadoPreviewForce";
      box.style.marginTop = "18px";
      container.appendChild(box);
    }

    const linhas = linhasDoTipo(resultado);
    const colunas = linhas[0] ? Object.keys(linhas[0]) : ["Sem dados"];
    const titulo = tipos[resultado.filtros.tipo] || "Relatório";

    box.innerHTML = `
      <div class="section-title" style="margin-bottom:10px;">📌 ${titulo}</div>
      <div class="table-wrap">
        <table class="table">
          <thead><tr>${colunas.map(c => `<th>${c}</th>`).join("")}</tr></thead>
          <tbody>
            ${
              linhas.length
                ? linhas.map(l => `<tr>${colunas.map(c => `<td>${valorTela(l[c])}</td>`).join("")}</tr>`).join("")
                : `<tr><td colspan="${colunas.length}" style="text-align:center;padding:18px;">Sem dados no período.</td></tr>`
            }
          </tbody>
        </table>
      </div>
    `;
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

  async function exportarExcel() {
    if (!state.ultimoResultado) await gerar();
    const resultado = state.ultimoResultado;

    await carregarExcelJS();

    const titulo = tipos[resultado.filtros.tipo] || "Relatório";
    const linhas = linhasDoTipo(resultado);
    const colunas = linhas[0] ? Object.keys(linhas[0]) : ["Sem dados"];

    const wb = new ExcelJS.Workbook();
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

    ws.addRow([]);
    ws.addRow(["Total de refeições", resultado.totais.total]);
    ws.addRow(["Custo Vascon estimado", resultado.totais.valorVascon]);
    ws.addRow(["Desconto funcionários", resultado.totais.desconto]);
    ws.addRow(["Valor unitário Vascon", resultado.valoresPeriodo.valorVascon]);
    ws.addRow(["Valor unitário descontado", resultado.valoresPeriodo.valorDesconto]);
    ws.addRow([]);

    const header = ws.addRow(colunas);
    header.eachCell(cell => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1F3C" } };
      cell.alignment = { horizontal: "center" };
      cell.border = borda();
    });

    if (linhas.length) {
      linhas.forEach((linha, idx) => {
        const row = ws.addRow(colunas.map(c => linha[c]));
        row.eachCell(cell => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? "FFF3F6FB" : "FFFFFFFF" } };
          cell.border = borda("FFD9E2F3");
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

    const nome = `relatorio-refeitorio-${slug(titulo)}-${resultado.filtros.dataInicio}-a-${resultado.filtros.dataFim}.xlsx`;
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = nome;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  function borda(color = "FFCBD5E1") {
    return {
      top: { style: "thin", color: { argb: color } },
      left: { style: "thin", color: { argb: color } },
      bottom: { style: "thin", color: { argb: color } },
      right: { style: "thin", color: { argb: color } }
    };
  }

  function moeda(v) {
    return Number(v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  }

  function valorTela(v) {
    if (typeof v === "number") return String(v).replace(".", ",");
    return v ?? "";
  }

  function dataBR(data) {
    if (!data) return "";
    if (typeof data === "string" && /^\d{4}-\d{2}-\d{2}$/.test(data)) {
      const [a, m, d] = data.split("-");
      return `${d}/${m}/${a}`;
    }
    const dt = new Date(data);
    return Number.isNaN(dt.getTime()) ? "" : dt.toLocaleDateString("pt-BR");
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

  window.AdminRelatoriosForceV2 = { prepararTela, gerar, exportarExcel };
  window.exportarCSV = exportarExcel;
  window.exportarRelatorioCSV = exportarExcel;
  window.exportarExcel = exportarExcel;
  window.gerarRelatorio = gerar;

  document.addEventListener("click", e => {
    const el = e.target.closest("button, a, .nav-item, [onclick]");
    if (!el) return;

    const texto = normalizar(`${el.innerText || ""} ${el.getAttribute("onclick") || ""}`);

    if (texto.includes("relatorios")) setTimeout(prepararTela, 500);

    if (estaNaTelaRelatorios() && texto.includes("gerar relatorio")) {
      e.preventDefault();
      gerar();
    }

    if (estaNaTelaRelatorios() && texto.includes("exportar")) {
      e.preventDefault();
      exportarExcel();
    }
  }, true);

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => setTimeout(prepararTela, 800));
  } else {
    setTimeout(prepararTela, 800);
  }
})();
