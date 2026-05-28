// ============================================================
// admin-relatorios.js — Relatórios somente na aba Relatórios + Excel .xlsx real
// ============================================================

window.AdminRelatorios = {
  pedidos: [],
  colaboradores: [],
  valores: [],
  ultimoResultado: null,

  tipos: {
    "resumo-dia": "Resumo por dia",
    "centro-custo": "Por centro de custo",
    "colaborador": "Quantidade por colaborador"
  },

  estaNaAbaRelatorios() {
    const ativo = document.querySelector(".module.active, .nav-item.active");
    const textoAtivo = (ativo?.innerText || "").toLowerCase();

    if (textoAtivo.includes("relatórios") || textoAtivo.includes("relatorios")) return true;

    const titulo = document.querySelector(".topbar-title, h1, h2");
    const textoTitulo = (titulo?.innerText || "").toLowerCase();

    return textoTitulo.includes("relatórios") || textoTitulo.includes("relatorios");
  },

  async init() {
    this.configurarBotoesGlobais();

    if (!this.estaNaAbaRelatorios()) return;

    this.criarSeletorForcado();
    await this.gerar();
  },

  configurarBotoesGlobais() {
    document.addEventListener("click", e => {
      const btn = e.target.closest("button");
      if (!btn) return;

      const texto = (btn.innerText || "").toLowerCase();

      if (texto.includes("exportar") && this.estaNaAbaRelatorios()) {
        e.preventDefault();
        this.exportarExcel();
      }

      if ((texto.includes("gerar relatório") || texto.includes("gerar relatorio")) && this.estaNaAbaRelatorios()) {
        e.preventDefault();
        this.gerar();
      }
    }, true);

    document.addEventListener("click", e => {
      const item = e.target.closest(".nav-item, [onclick]");
      if (!item) return;

      const txt = `${item.innerText || ""} ${item.getAttribute("onclick") || ""}`.toLowerCase();
      if (txt.includes("relatórios") || txt.includes("relatorios")) {
        setTimeout(() => this.init(), 300);
      }
    });
  },

  criarSeletorForcado() {
    if (!this.estaNaAbaRelatorios()) return;
    if (document.getElementById("tipoRelatorioBox")) return;

    const titulo = this.encontrarTituloRelatorios();
    const container = titulo?.closest(".module") || document.querySelector(".module.active") || document.querySelector(".content") || document.body;

    const box = document.createElement("div");
    box.id = "tipoRelatorioBox";
    box.style.cssText = `
      margin: 1rem 0 1.2rem 0;
      padding: 1rem;
      border: 1px solid rgba(60,140,255,.28);
      border-radius: 12px;
      background: rgba(10,31,60,.55);
      display: grid;
      grid-template-columns: minmax(240px, 360px) 1fr;
      gap: 1rem;
      align-items: end;
    `;

    box.innerHTML = `
      <div class="form-group">
        <label class="form-label">TIPO DE RELATÓRIO</label>
        <select id="tipoRelatorio" class="form-select" onchange="AdminRelatorios.gerar()">
          <option value="resumo-dia">Resumo por dia</option>
          <option value="centro-custo">Por centro de custo</option>
          <option value="colaborador">Quantidade por colaborador</option>
        </select>
      </div>
      <div style="font-size:.78rem;color:#75b7ff;line-height:1.45;">
        Escolha o tipo, defina o período e clique em <b>Gerar relatório</b>. A exportação baixa somente o tipo escolhido em Excel real.
      </div>
    `;

    const formGrid = Array.from(container.querySelectorAll(".form-grid")).find(el =>
      el.innerText.toLowerCase().includes("data inicial") || el.querySelector('input[type="date"]')
    );

    if (formGrid) formGrid.insertAdjacentElement("afterend", box);
    else container.prepend(box);

    document.querySelectorAll("button").forEach(btn => {
      const texto = (btn.innerText || "").toLowerCase();
      if (texto.includes("csv") || texto.includes("exportar")) btn.innerHTML = "📥 Exportar Excel";
    });
  },

  encontrarTituloRelatorios() {
    return Array.from(document.querySelectorAll("h1,h2,h3,.section-title,.topbar-title")).find(el => {
      const t = (el.innerText || "").toLowerCase();
      return t.includes("relatórios") || t.includes("relatorios");
    });
  },

  async carregarDados() {
    if (!window.SP) throw new Error("SP não encontrado.");

    const [pedidos, colaboradores, valores] = await Promise.all([
      this.safe(() => SP.getItems("Pedidos"), []),
      this.safe(() => SP.getTodosColaboradores ? SP.getTodosColaboradores() : SP.getColaboradores(), []),
      this.safe(() => SP.getValoresRefeicao ? SP.getValoresRefeicao() : [], [])
    ]);

    this.pedidos = pedidos || [];
    this.colaboradores = colaboradores || [];
    this.valores = valores || [];
  },

  async safe(fn, fallback) {
    try { return await fn(); }
    catch (e) { console.warn("Falha ao carregar relatório:", e); return fallback; }
  },

  getFiltros() {
    const container = document.querySelector(".module.active") || document.body;
    const inputsDate = Array.from(container.querySelectorAll('input[type="date"]'));

    const dataInicio =
      document.getElementById("relDataInicio")?.value ||
      document.getElementById("dataInicioRelatorio")?.value ||
      inputsDate[0]?.value ||
      this.toInputDate(new Date());

    const dataFim =
      document.getElementById("relDataFim")?.value ||
      document.getElementById("dataFimRelatorio")?.value ||
      inputsDate[1]?.value ||
      this.toInputDate(new Date());

    const tipo = document.getElementById("tipoRelatorio")?.value || "resumo-dia";

    const nfCampo =
      document.getElementById("nfVasconRecebida") ||
      document.getElementById("valorNfVascon") ||
      Array.from(container.querySelectorAll("input")).find(i => (i.placeholder || "").toLowerCase().includes("1234"));

    const nfVascon = Number(String(nfCampo?.value || "0").replace(",", ".")) || 0;

    return { dataInicio, dataFim, tipo, nfVascon };
  },

  async gerar() {
    if (!this.estaNaAbaRelatorios()) return;

    try {
      this.criarSeletorForcado();
      await this.carregarDados();

      const filtros = this.getFiltros();
      const pedidosPeriodo = this.filtrarPedidos(filtros.dataInicio, filtros.dataFim);
      const valoresPeriodo = this.obterValoresPeriodo(filtros.dataInicio, filtros.dataFim);

      const resultado = {
        filtros,
        valoresPeriodo,
        pedidos: pedidosPeriodo,
        resumoDia: this.gerarResumoDia(filtros.dataInicio, filtros.dataFim, pedidosPeriodo),
        centroCusto: this.gerarCentroCusto(pedidosPeriodo, valoresPeriodo),
        colaborador: this.gerarColaborador(pedidosPeriodo, valoresPeriodo)
      };

      resultado.totais = this.gerarTotais(resultado, filtros);

      this.ultimoResultado = resultado;
      this.renderizar(resultado);
      return resultado;
    } catch (erro) {
      console.error("Erro ao gerar relatório:", erro);
      alert(`Erro ao gerar relatório: ${erro.message || erro}`);
    }
  },

  filtrarPedidos(dataInicio, dataFim) {
    const ini = new Date(`${dataInicio}T00:00:00`);
    const fim = new Date(`${dataFim}T23:59:59`);

    return this.pedidos.filter(p => {
      const status = this.normalizar(p.Status || "");
      if (["cancelado", "bloqueado", "afastado", "ferias", "férias", "nao vai almocar", "não vai almoçar"].includes(status)) return false;

      const data = this.dataPedido(p);
      if (!data) return false;

      return data >= ini && data <= fim;
    });
  },

  dataPedido(p) {
    if (p.Data_Hora) {
      const d = new Date(p.Data_Hora);
      if (!Number.isNaN(d.getTime())) return d;
    }

    if (p.Data) {
      const d = new Date(p.Data);
      if (!Number.isNaN(d.getTime())) return d;
    }

    const semana = p.Semana_id || "";
    const dia = p.Dia || "";
    const match = String(semana).match(/(\d{4})-?W?(\d{1,2})/i);

    if (match && dia) {
      const ano = Number(match[1]);
      const semanaNum = Number(match[2]);
      const segunda = this.dataInicioSemanaISO(ano, semanaNum);

      const mapa = { segunda: 0, terca: 1, terça: 1, quarta: 2, quinta: 3, sexta: 4, sabado: 5, sábado: 5, domingo: 6 };
      const idx = mapa[this.normalizar(dia)];

      if (idx !== undefined) {
        const d = new Date(segunda);
        d.setDate(d.getDate() + idx);
        return d;
      }
    }

    return null;
  },

  dataInicioSemanaISO(ano, semana) {
    const jan4 = new Date(ano, 0, 4);
    const diaSemana = jan4.getDay() || 7;
    const segundaSemana1 = new Date(jan4);
    segundaSemana1.setDate(jan4.getDate() - diaSemana + 1);

    const segunda = new Date(segundaSemana1);
    segunda.setDate(segundaSemana1.getDate() + (semana - 1) * 7);
    return segunda;
  },

  obterValoresPeriodo(dataInicio, dataFim) {
    const ini = new Date(`${dataInicio}T00:00:00`);
    const fim = new Date(`${dataFim}T23:59:59`);

    const candidatos = this.valores.filter(v => {
      const ativo = typeof SP?.isTrue === "function" ? SP.isTrue(v.Ativo) : String(v.Ativo).toLowerCase() !== "false";
      const vi = v.Data_Inicio ? new Date(v.Data_Inicio) : null;
      const vf = v.Data_Fim ? new Date(v.Data_Fim) : null;

      if (!ativo) return false;
      if (!vi || !vf) return true;
      return vi <= fim && vf >= ini;
    });

    const v = candidatos[0] || this.valores[0] || {};

    return {
      valorVascon: Number(v.Valor_Vascon || 0),
      valorDesconto: Number(v.Valor_Desconto_Funcionario || 0),
      titulo: v.Title || ""
    };
  },

  gerarTotais(resultado, filtros) {
    const total = resultado.pedidos.length;
    const valorVascon = total * resultado.valoresPeriodo.valorVascon;
    const desconto = total * resultado.valoresPeriodo.valorDesconto;

    return {
      total,
      valorVascon,
      desconto,
      diferencaNf: filtros.nfVascon ? filtros.nfVascon - valorVascon : null,
      opcoes: this.contarOpcoes(resultado.pedidos)
    };
  },

  gerarResumoDia(dataInicio, dataFim, pedidos) {
    const linhas = [];
    const ini = new Date(`${dataInicio}T00:00:00`);
    const fim = new Date(`${dataFim}T00:00:00`);

    for (let d = new Date(ini); d <= fim; d.setDate(d.getDate() + 1)) {
      const chave = this.toInputDate(d);
      const pedidosDia = pedidos.filter(p => {
        const data = this.dataPedido(p);
        return data && this.toInputDate(data) === chave;
      });

      const op = this.contarOpcoes(pedidosDia);

      linhas.push({
        "Dia": this.diaSemanaCompleto(d),
        "Data": this.dataBR(d),
        "Principal": op.Principal,
        "Light": op.Light,
        "Carne": op.Carne,
        "Massa": op.Massa,
        "Lanche": op.Lanche,
        "Total": pedidosDia.length
      });
    }

    return linhas;
  },

  gerarCentroCusto(pedidos, valores) {
    const mapa = {};

    for (const p of pedidos) {
      const centro = p.Centro_Custo || this.centroCustoDoColaborador(p) || "Sem centro de custo";

      if (!mapa[centro]) {
        mapa[centro] = { "Centro de Custo": centro, "Quantidade": 0, "Valor Vascon": 0, "Desconto Funcionários": 0, "Rateio NF": 0 };
      }

      mapa[centro]["Quantidade"] += 1;
      mapa[centro]["Valor Vascon"] += valores.valorVascon;
      mapa[centro]["Desconto Funcionários"] += valores.valorDesconto;
      mapa[centro]["Rateio NF"] += valores.valorVascon;
    }

    return Object.values(mapa);
  },

  gerarColaborador(pedidos, valores) {
    const mapa = {};

    for (const p of pedidos) {
      const nome = p.Colaborador_nome || p.Colaborador || "Sem nome";
      const centro = p.Centro_Custo || this.centroCustoDoColaborador(p) || "Sem centro de custo";
      const chave = `${nome}|${centro}`;

      if (!mapa[chave]) {
        mapa[chave] = { "Colaborador": nome, "Centro de Custo": centro, "Quantidade": 0, "Valor de cada refeição": valores.valorDesconto, "Desconto em folha": 0 };
      }

      mapa[chave]["Quantidade"] += 1;
      mapa[chave]["Desconto em folha"] += valores.valorDesconto;
    }

    return Object.values(mapa).sort((a, b) => String(a.Colaborador).localeCompare(String(b.Colaborador)));
  },

  centroCustoDoColaborador(pedido) {
    const id = String(pedido.Colaborador_id || "");
    const nome = this.normalizar(pedido.Colaborador_nome || "");

    const colab = this.colaboradores.find(c =>
      String(c.id) === id ||
      String(c.ID) === id ||
      this.normalizar(c.Nome || c.Title || "") === nome
    );

    return colab?.Centro_Custo || "";
  },

  contarOpcoes(pedidos) {
    const base = { Principal: 0, Light: 0, Carne: 0, Massa: 0, Lanche: 0 };

    for (const p of pedidos) {
      const opcao = this.normalizar(p.Opcao || "");
      if (opcao.includes("principal")) base.Principal++;
      else if (opcao.includes("light") || opcao.includes("ligth")) base.Light++;
      else if (opcao.includes("carne")) base.Carne++;
      else if (opcao.includes("massa")) base.Massa++;
      else if (opcao.includes("lanche")) base.Lanche++;
    }

    return base;
  },

  renderizar(resultado) {
    const op = resultado.totais.opcoes;

    this.setTextoPossivel(["relTotalRefeicoes", "statRelRefeicoes"], resultado.totais.total);
    this.setTextoPossivel(["relCustoVascon", "statRelVascon"], this.moeda(resultado.totais.valorVascon));
    this.setTextoPossivel(["relDescontoFuncionarios", "statRelDesconto"], this.moeda(resultado.totais.desconto));
    this.setTextoPossivel(["relDiferencaNf", "statRelDiferenca"], resultado.totais.diferencaNf === null ? "-" : this.moeda(resultado.totais.diferencaNf));
    this.setTextoPossivel(["relPrincipal"], op.Principal);
    this.setTextoPossivel(["relLight", "relLigth"], op.Light);
    this.setTextoPossivel(["relCarne"], op.Carne);
    this.setTextoPossivel(["relMassa"], op.Massa);
    this.setTextoPossivel(["relLanche"], op.Lanche);

    this.renderizarPreviewSelecionado(resultado);
  },

  renderizarPreviewSelecionado(resultado) {
    if (!this.estaNaAbaRelatorios()) return;

    const area = document.querySelector(".module.active") || document.querySelector(".content") || document.body;
    let bloco = document.getElementById("relatorioSelecionadoPreview");

    if (!bloco) {
      bloco = document.createElement("div");
      bloco.id = "relatorioSelecionadoPreview";
      bloco.style.marginTop = "1.5rem";
      area.appendChild(bloco);
    }

    const tipo = resultado.filtros.tipo;
    const titulo = this.tipos[tipo] || "Relatório";
    const linhas = this.linhasDoTipo(resultado);
    const colunas = linhas[0] ? Object.keys(linhas[0]) : ["Sem dados"];

    bloco.innerHTML = `
      <div class="section-title" style="margin-bottom:.8rem;">📌 ${titulo}</div>
      <div class="table-wrap">
        <table class="table">
          <thead><tr>${colunas.map(c => `<th>${c}</th>`).join("")}</tr></thead>
          <tbody>
            ${
              linhas.length
                ? linhas.map(l => `<tr>${colunas.map(c => `<td>${this.valorTela(l[c])}</td>`).join("")}</tr>`).join("")
                : `<tr><td colspan="${colunas.length}" style="text-align:center;padding:1.5rem;">Sem dados no período.</td></tr>`
            }
          </tbody>
        </table>
      </div>
    `;
  },

  linhasDoTipo(resultado) {
    if (resultado.filtros.tipo === "centro-custo") return resultado.centroCusto;
    if (resultado.filtros.tipo === "colaborador") return resultado.colaborador;
    return resultado.resumoDia;
  },

  async carregarExcelJS() {
    if (window.ExcelJS) return;

    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js";
      s.onload = resolve;
      s.onerror = reject;
      document.head.appendChild(s);
    });
  },

  async exportarExcel() {
    if (!this.ultimoResultado) await this.gerar();
    const resultado = this.ultimoResultado;
    if (!resultado) return;

    await this.carregarExcelJS();

    const tipo = resultado.filtros.tipo;
    const titulo = this.tipos[tipo] || "Relatório";
    const linhas = this.linhasDoTipo(resultado);
    const colunas = linhas[0] ? Object.keys(linhas[0]) : ["Sem dados"];

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(titulo.substring(0, 31));

    const totalCols = Math.max(colunas.length, 6);

    ws.mergeCells(1, 1, 1, totalCols);
    ws.getCell("A1").value = `RELATÓRIO REFEITÓRIO HOMY — ${titulo.toUpperCase()}`;
    ws.getCell("A1").font = { bold: true, color: { argb: "FFFFFFFF" }, size: 16 };
    ws.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1F3C" } };
    ws.getCell("A1").alignment = { horizontal: "center" };

    ws.mergeCells(2, 1, 2, totalCols);
    ws.getCell("A2").value = `Período: ${this.dataBR(resultado.filtros.dataInicio)} a ${this.dataBR(resultado.filtros.dataFim)}`;
    ws.getCell("A2").font = { bold: true, color: { argb: "FFFFFFFF" } };
    ws.getCell("A2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC0281C" } };
    ws.getCell("A2").alignment = { horizontal: "center" };

    ws.addRow([]);
    ws.addRow(["Total de refeições", resultado.totais.total]);
    ws.addRow(["Custo Vascon estimado", resultado.totais.valorVascon]);
    ws.addRow(["Desconto funcionários", resultado.totais.desconto]);
    ws.addRow(["Valor unitário Vascon", resultado.valoresPeriodo.valorVascon]);
    ws.addRow(["Valor unitário descontado", resultado.valoresPeriodo.valorDesconto]);
    ws.addRow([]);

    const headerRow = ws.addRow(colunas);
    headerRow.eachCell(cell => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1F3C" } };
      cell.alignment = { horizontal: "center" };
      cell.border = {
        top: { style: "thin" }, left: { style: "thin" },
        bottom: { style: "thin" }, right: { style: "thin" }
      };
    });

    if (linhas.length) {
      linhas.forEach((linha, idx) => {
        const row = ws.addRow(colunas.map(c => linha[c]));
        row.eachCell(cell => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? "FFF3F6FB" : "FFFFFFFF" } };
          cell.border = { bottom: { style: "thin", color: { argb: "FFD9E2F3" } } };
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
      col.width = Math.min(max, 35);
    });

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

    const nomeArquivo = `relatorio-refeitorio-${this.slug(titulo)}-${resultado.filtros.dataInicio}-a-${resultado.filtros.dataFim}.xlsx`;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = nomeArquivo;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  },

  setTextoPossivel(ids, valor) {
    for (const id of ids) {
      const el = document.getElementById(id);
      if (el) { el.textContent = valor; return; }
    }
  },

  valorTela(v) {
    if (typeof v === "number") return String(v).replace(".", ",");
    return v ?? "";
  },

  moeda(v) {
    return Number(v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  },

  dataBR(data) {
    if (!data) return "";
    if (typeof data === "string" && /^\d{4}-\d{2}-\d{2}$/.test(data)) {
      const [a, m, d] = data.split("-");
      return `${d}/${m}/${a}`;
    }
    const dt = new Date(data);
    if (Number.isNaN(dt.getTime())) return "";
    return dt.toLocaleDateString("pt-BR");
  },

  toInputDate(data) {
    const d = new Date(data);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  },

  diaSemanaCompleto(data) {
    return ["Domingo", "Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado"][new Date(data).getDay()];
  },

  normalizar(valor) {
    return String(valor || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();
  },

  slug(valor) {
    return this.normalizar(valor).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  }
};

window.gerarRelatorio = () => AdminRelatorios.gerar();
window.exportarCSV = () => AdminRelatorios.exportarExcel();
window.exportarRelatorioCSV = () => AdminRelatorios.exportarExcel();
window.exportarExcel = () => AdminRelatorios.exportarExcel();

document.addEventListener("DOMContentLoaded", () => {
  setTimeout(() => AdminRelatorios.init(), 800);
});
