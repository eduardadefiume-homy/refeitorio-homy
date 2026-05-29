// ============================================================
// admin-relatorios.js
// Relatórios com seletor de tipo + exportação Excel .xlsx estilizada
// ============================================================

(function () {
  const Relatorios = {
    pedidos: [],
    colaboradores: [],
    valores: [],
    ultimoResultado: null,

    tipos: {
      "resumo-dia": "Resumo por dia",
      "centro-custo": "Por centro de custo",
      "colaborador": "Quantidade por colaborador"
    },

    init() {
      this.configurarEventos();
      setTimeout(() => this.prepararTelaRelatorios(), 600);
    },

    configurarEventos() {
      document.addEventListener("click", e => {
        const el = e.target.closest("button, a, .nav-item, [onclick]");
        if (!el) return;

        const texto = this.normalizar(`${el.innerText || ""} ${el.getAttribute("onclick") || ""}`);

        if (texto.includes("relatorios")) {
          setTimeout(() => this.prepararTelaRelatorios(), 500);
        }

        if (this.estaNaTelaRelatorios() && texto.includes("gerar relatorio")) {
          e.preventDefault();
          this.gerar();
        }

        if (this.estaNaTelaRelatorios() && texto.includes("exportar")) {
          e.preventDefault();
          this.exportarExcel();
        }
      }, true);

      document.addEventListener("change", e => {
        if (e.target?.id === "tipoRelatorio") {
          this.gerar();
        }
      });
    },

    estaNaTelaRelatorios() {
      const titulo = document.querySelector(".topbar-title, h1, h2");
      if (this.normalizar(titulo?.innerText || "").includes("relatorios")) return true;

      const ativo = document.querySelector(".nav-item.active, .module.active");
      return this.normalizar(ativo?.innerText || "").includes("relatorios");
    },

    prepararTelaRelatorios() {
      if (!this.estaNaTelaRelatorios()) return;

      this.forcarBotaoExcel();
      this.inserirSeletorTipo();
      this.gerar();
    },

    forcarBotaoExcel() {
      document.querySelectorAll("button, a").forEach(btn => {
        const texto = this.normalizar(btn.innerText || "");
        if (texto.includes("exportar csv") || texto.includes("exportar excel") || texto === "exportar") {
          btn.innerHTML = "📥 Exportar Excel";
        }
      });
    },

    inserirSeletorTipo() {
      if (document.getElementById("tipoRelatorioBox")) return;

      const container = this.containerRelatorios();
      if (!container) return;

      const box = document.createElement("div");
      box.id = "tipoRelatorioBox";
      box.style.cssText = `
        width: 100%;
        margin: 14px 0 16px 0;
        padding: 14px;
        border: 1px solid rgba(60,140,255,.35);
        border-radius: 12px;
        background: rgba(10,31,60,.55);
        display: grid;
        grid-template-columns: 300px 1fr;
        gap: 14px;
        align-items: end;
        box-sizing: border-box;
      `;

      box.innerHTML = `
        <div class="form-group" style="margin:0;">
          <label class="form-label">TIPO DE RELATÓRIO</label>
          <select id="tipoRelatorio" class="form-select">
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

      const alvo =
        Array.from(container.querySelectorAll("button, a")).find(b => this.normalizar(b.innerText || "").includes("gerar relatorio"))?.closest("div") ||
        Array.from(container.querySelectorAll("input[type='date']")).at(-1)?.closest("div")?.parentElement ||
        container.querySelector(".form-grid");

      if (alvo) alvo.insertAdjacentElement("afterend", box);
      else container.prepend(box);
    },

    containerRelatorios() {
      return document.querySelector(".module.active") ||
             document.querySelector("main") ||
             document.querySelector(".content") ||
             document.body;
    },

    async gerar() {
      if (!this.estaNaTelaRelatorios()) return;

      this.forcarBotaoExcel();
      this.inserirSeletorTipo();

      await this.carregarDados();

      const filtros = this.obterFiltros();
      const pedidosPeriodo = this.filtrarPedidos(filtros.dataInicio, filtros.dataFim);
      const valoresPeriodo = this.obterValoresPeriodo(filtros.dataInicio, filtros.dataFim);

      const resultado = {
        filtros,
        pedidos: pedidosPeriodo,
        valoresPeriodo,
        resumoDia: this.montarResumoDia(filtros.dataInicio, filtros.dataFim, pedidosPeriodo),
        centroCusto: this.montarCentroCusto(pedidosPeriodo, valoresPeriodo),
        colaborador: this.montarColaborador(pedidosPeriodo, valoresPeriodo)
      };

      resultado.totais = this.montarTotais(resultado, filtros);
      this.ultimoResultado = resultado;

      this.atualizarCards(resultado);
      this.renderPreview(resultado);
    },

    async carregarDados() {
      if (!window.SP) throw new Error("SP não encontrado.");

      const get = async fn => {
        try { return await fn(); }
        catch (e) { console.warn("Falha relatório:", e); return []; }
      };

      this.pedidos = await get(() => SP.getItems ? SP.getItems("Pedidos") : SP.getPedidos());
      this.colaboradores = await get(() => SP.getTodosColaboradores ? SP.getTodosColaboradores() : SP.getColaboradores());
      this.valores = await get(() => SP.getValoresRefeicao ? SP.getValoresRefeicao() : []);
    },

    obterFiltros() {
      const container = this.containerRelatorios();
      const datas = Array.from(container.querySelectorAll("input[type='date']"));

      const dataInicio =
        document.getElementById("relDataInicio")?.value ||
        document.getElementById("dataInicioRelatorio")?.value ||
        datas[0]?.value ||
        this.toInputDate(new Date());

      const dataFim =
        document.getElementById("relDataFim")?.value ||
        document.getElementById("dataFimRelatorio")?.value ||
        datas[1]?.value ||
        this.toInputDate(new Date());

      const tipo = document.getElementById("tipoRelatorio")?.value || "resumo-dia";

      const campoNf =
        document.getElementById("nfVasconRecebida") ||
        document.getElementById("valorNfVascon") ||
        Array.from(container.querySelectorAll("input")).find(i => (i.placeholder || "").includes("1234"));

      const nfVascon = Number(String(campoNf?.value || "0").replace(",", ".")) || 0;

      return { dataInicio, dataFim, tipo, nfVascon };
    },

    filtrarPedidos(dataInicio, dataFim) {
      const ini = new Date(`${dataInicio}T00:00:00`);
      const fim = new Date(`${dataFim}T23:59:59`);

      return (this.pedidos || []).filter(p => {
        const status = this.normalizar(p.Status || "");
        if (["cancelado", "bloqueado", "afastado", "ferias", "nao vai almocar"].includes(status)) return false;

        const d = this.dataPedido(p);
        return d && d >= ini && d <= fim;
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

      const semana = String(p.Semana_id || "");
      const dia = String(p.Dia || "");
      const match = semana.match(/(\d{4})-?W?(\d{1,2})/i);

      if (match && dia) {
        const segunda = this.inicioSemanaISO(Number(match[1]), Number(match[2]));
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

    inicioSemanaISO(ano, semana) {
      const jan4 = new Date(ano, 0, 4);
      const dia = jan4.getDay() || 7;
      const segunda = new Date(jan4);
      segunda.setDate(jan4.getDate() - dia + 1 + (semana - 1) * 7);
      return segunda;
    },

    obterValoresPeriodo(dataInicio, dataFim) {
      const ini = new Date(`${dataInicio}T00:00:00`);
      const fim = new Date(`${dataFim}T23:59:59`);

      const ativos = (this.valores || []).filter(v => {
        const ativo = SP.isTrue ? SP.isTrue(v.Ativo) : String(v.Ativo).toLowerCase() !== "false";
        const vi = v.Data_Inicio ? new Date(v.Data_Inicio) : null;
        const vf = v.Data_Fim ? new Date(v.Data_Fim) : null;

        if (!ativo) return false;
        if (!vi || !vf) return true;
        return vi <= fim && vf >= ini;
      });

      const v = ativos[0] || this.valores[0] || {};

      return {
        valorVascon: Number(v.Valor_Vascon || 0),
        valorDesconto: Number(v.Valor_Desconto_Funcionario || 0),
        titulo: v.Title || ""
      };
    },

    montarTotais(resultado, filtros) {
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

    montarResumoDia(dataInicio, dataFim, pedidos) {
      const linhas = [];
      const ini = new Date(`${dataInicio}T00:00:00`);
      const fim = new Date(`${dataFim}T00:00:00`);

      for (let d = new Date(ini); d <= fim; d.setDate(d.getDate() + 1)) {
        const chave = this.toInputDate(d);
        const pedidosDia = pedidos.filter(p => {
          const dp = this.dataPedido(p);
          return dp && this.toInputDate(dp) === chave;
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

    montarCentroCusto(pedidos, valores) {
      const mapa = {};

      pedidos.forEach(p => {
        const centro = p.Centro_Custo || this.centroCustoDoColaborador(p) || "Sem centro de custo";

        if (!mapa[centro]) {
          mapa[centro] = {
            "Centro de Custo": centro,
            "Quantidade": 0,
            "Valor Vascon": 0,
            "Desconto Funcionários": 0,
            "Rateio NF": 0
          };
        }

        mapa[centro]["Quantidade"] += 1;
        mapa[centro]["Valor Vascon"] += valores.valorVascon;
        mapa[centro]["Desconto Funcionários"] += valores.valorDesconto;
        mapa[centro]["Rateio NF"] += valores.valorVascon;
      });

      return Object.values(mapa);
    },

    montarColaborador(pedidos, valores) {
      const mapa = {};

      pedidos.forEach(p => {
        const nome = p.Colaborador_nome || p.Colaborador || "Sem nome";
        const centro = p.Centro_Custo || this.centroCustoDoColaborador(p) || "Sem centro de custo";
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

        mapa[chave]["Quantidade"] += 1;
        mapa[chave]["Desconto em folha"] += valores.valorDesconto;
      });

      return Object.values(mapa).sort((a, b) => String(a.Colaborador).localeCompare(String(b.Colaborador)));
    },

    centroCustoDoColaborador(pedido) {
      const id = String(pedido.Colaborador_id || "");
      const nome = this.normalizar(pedido.Colaborador_nome || "");

      const c = (this.colaboradores || []).find(x =>
        String(x.id) === id ||
        String(x.ID) === id ||
        this.normalizar(x.Nome || x.Title || "") === nome
      );

      return c?.Centro_Custo || "";
    },

    contarOpcoes(pedidos) {
      const r = { Principal: 0, Light: 0, Carne: 0, Massa: 0, Lanche: 0 };

      pedidos.forEach(p => {
        const o = this.normalizar(p.Opcao || p.Nome_Prato || "");
        if (o.includes("principal")) r.Principal++;
        else if (o.includes("light") || o.includes("ligth")) r.Light++;
        else if (o.includes("carne")) r.Carne++;
        else if (o.includes("massa")) r.Massa++;
        else if (o.includes("lanche")) r.Lanche++;
      });

      return r;
    },

    atualizarCards(resultado) {
      const op = resultado.totais.opcoes;

      this.setTexto(["relTotalRefeicoes", "statRelRefeicoes"], resultado.totais.total);
      this.setTexto(["relCustoVascon", "statRelVascon"], this.moeda(resultado.totais.valorVascon));
      this.setTexto(["relDescontoFuncionarios", "statRelDesconto"], this.moeda(resultado.totais.desconto));
      this.setTexto(["relDiferencaNf", "statRelDiferenca"], resultado.totais.diferencaNf === null ? "-" : this.moeda(resultado.totais.diferencaNf));
      this.setTexto(["relPrincipal"], op.Principal);
      this.setTexto(["relLight", "relLigth"], op.Light);
      this.setTexto(["relCarne"], op.Carne);
      this.setTexto(["relMassa"], op.Massa);
      this.setTexto(["relLanche"], op.Lanche);
    },

    renderPreview(resultado) {
      let box = document.getElementById("relatorioSelecionadoPreview");
      const container = this.containerRelatorios();

      if (!box) {
        box = document.createElement("div");
        box.id = "relatorioSelecionadoPreview";
        box.style.marginTop = "18px";
        container.appendChild(box);
      }

      const linhas = this.linhasDoTipo(resultado);
      const colunas = linhas[0] ? Object.keys(linhas[0]) : ["Sem dados"];
      const titulo = this.tipos[resultado.filtros.tipo] || "Relatório";

      box.innerHTML = `
        <div class="section-title" style="margin-bottom:10px;">📌 ${titulo}</div>
        <div class="table-wrap">
          <table class="table">
            <thead><tr>${colunas.map(c => `<th>${c}</th>`).join("")}</tr></thead>
            <tbody>
              ${
                linhas.length
                  ? linhas.map(l => `<tr>${colunas.map(c => `<td>${this.valorTela(l[c])}</td>`).join("")}</tr>`).join("")
                  : `<tr><td colspan="${colunas.length}" style="text-align:center;padding:18px;">Sem dados no período.</td></tr>`
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
        s.onerror = () => reject(new Error("Não foi possível carregar ExcelJS."));
        document.head.appendChild(s);
      });
    },

    async exportarExcel() {
      if (!this.ultimoResultado) await this.gerar();
      const resultado = this.ultimoResultado;

      await this.carregarExcelJS();

      const tipo = resultado.filtros.tipo;
      const titulo = this.tipos[tipo] || "Relatório";
      const linhas = this.linhasDoTipo(resultado);
      const colunas = linhas[0] ? Object.keys(linhas[0]) : ["Sem dados"];

      const wb = new ExcelJS.Workbook();
      wb.creator = "Homy Refeitório";
      const ws = wb.addWorksheet(titulo.substring(0, 31));

      const totalCols = Math.max(colunas.length, 6);

      ws.mergeCells(1, 1, 1, totalCols);
      ws.getCell(1, 1).value = `RELATÓRIO REFEITÓRIO HOMY — ${titulo.toUpperCase()}`;
      ws.getCell(1, 1).font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
      ws.getCell(1, 1).alignment = { horizontal: "center" };
      ws.getCell(1, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1F3C" } };

      ws.mergeCells(2, 1, 2, totalCols);
      ws.getCell(2, 1).value = `Período: ${this.dataBR(resultado.filtros.dataInicio)} a ${this.dataBR(resultado.filtros.dataFim)}`;
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

      for (let i = 5; i <= 9; i++) {
        ws.getCell(i, 1).font = { bold: true };
        ws.getCell(i, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F1FF" } };
        ws.getCell(i, 2).numFmt = i === 5 ? "0" : '"R$" #,##0.00';
        ws.getCell(i, 2).border = this.borda();
      }

      const header = ws.addRow(colunas);
      header.eachCell(cell => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1F3C" } };
        cell.alignment = { horizontal: "center" };
        cell.border = this.borda();
      });

      const moneyCols = ["Valor Vascon", "Desconto Funcionários", "Rateio NF", "Valor de cada refeição", "Desconto em folha"];

      if (linhas.length) {
        linhas.forEach((linha, idx) => {
          const row = ws.addRow(colunas.map(c => linha[c]));
          row.eachCell((cell, colNumber) => {
            const colName = colunas[colNumber - 1];
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? "FFF3F6FB" : "FFFFFFFF" } };
            cell.border = this.borda("FFD9E2F3");
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

      ws.views = [{ state: "frozen", ySplit: 11 }];

      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });

      const nome = `relatorio-refeitorio-${this.slug(titulo)}-${resultado.filtros.dataInicio}-a-${resultado.filtros.dataFim}.xlsx`;
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = nome;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    },

    borda(color = "FFCBD5E1") {
      return {
        top: { style: "thin", color: { argb: color } },
        left: { style: "thin", color: { argb: color } },
        bottom: { style: "thin", color: { argb: color } },
        right: { style: "thin", color: { argb: color } }
      };
    },

    setTexto(ids, valor) {
      for (const id of ids) {
        const el = document.getElementById(id);
        if (el) {
          el.textContent = valor;
          return;
        }
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
      return Number.isNaN(dt.getTime()) ? "" : dt.toLocaleDateString("pt-BR");
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

  window.AdminRelatorios = Relatorios;
  window.gerarRelatorio = () => Relatorios.gerar();
  window.exportarCSV = () => Relatorios.exportarExcel();
  window.exportarRelatorioCSV = () => Relatorios.exportarExcel();
  window.exportarExcel = () => Relatorios.exportarExcel();

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => Relatorios.init());
  } else {
    Relatorios.init();
  }
})();
