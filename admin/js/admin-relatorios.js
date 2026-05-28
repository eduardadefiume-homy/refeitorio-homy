// ============================================================
// admin-relatorios.js — Relatórios por tipo + Excel estilizado
// Correção: seletor visível, exportação bonita e relatórios separados
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

  async init() {
    this.montarControles();
    await this.gerar();
  },

  montarControles() {
    const areaRelatorios = this.encontrarAreaRelatorios();
    if (!areaRelatorios) return;

    // Renomear botão de CSV para Excel
    document.querySelectorAll("button").forEach(btn => {
      const txt = (btn.innerText || "").toLowerCase();
      if (txt.includes("csv")) {
        btn.innerHTML = "📥 Exportar Excel";
        btn.onclick = () => AdminRelatorios.exportarExcel();
      }
    });

    // Criar seletor se ainda não existir
    if (!document.getElementById("tipoRelatorio")) {
      const bloco = document.createElement("div");
      bloco.className = "form-group";
      bloco.innerHTML = `
        <label class="form-label">TIPO DE RELATÓRIO</label>
        <select id="tipoRelatorio" class="form-select" onchange="AdminRelatorios.gerar()">
          <option value="resumo-dia">Resumo por dia</option>
          <option value="centro-custo">Por centro de custo</option>
          <option value="colaborador">Quantidade por colaborador</option>
        </select>
      `;

      const formGrid = areaRelatorios.querySelector(".form-grid") || document.querySelector(".form-grid");
      if (formGrid) {
        formGrid.insertBefore(bloco, formGrid.children[2] || null);
      } else {
        areaRelatorios.prepend(bloco);
      }
    }

    // Botão gerar
    document.querySelectorAll("button").forEach(btn => {
      const txt = (btn.innerText || "").toLowerCase();
      if (txt.includes("gerar relatório") || txt.includes("gerar relatorio")) {
        btn.onclick = () => AdminRelatorios.gerar();
      }
    });
  },

  encontrarAreaRelatorios() {
    const modulos = Array.from(document.querySelectorAll(".module, section, main, div"));
    return modulos.find(el => {
      const texto = (el.innerText || "").toLowerCase();
      return texto.includes("relatórios gerenciais") || texto.includes("relatorios gerenciais");
    }) || document.body;
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
    try {
      return await fn();
    } catch (e) {
      console.warn("Falha ao carregar dado do relatório:", e);
      return fallback;
    }
  },

  getFiltros() {
    const inputsDate = Array.from(document.querySelectorAll('input[type="date"]'));

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
      Array.from(document.querySelectorAll("input")).find(i =>
        (i.placeholder || "").toLowerCase().includes("1234")
      );

    const nfVascon = Number(String(nfCampo?.value || "0").replace(",", ".")) || 0;

    return { dataInicio, dataFim, tipo, nfVascon };
  },

  async gerar() {
    try {
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
      if (["cancelado", "bloqueado", "afastado", "ferias", "férias", "nao vai almocar", "não vai almoçar"].includes(status)) {
        return false;
      }

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

    // Fallback para Semana_id + Dia
    const semana = p.Semana_id || p.semana_id || "";
    const dia = p.Dia || "";

    const match = String(semana).match(/(\d{4})-?W?(\d{1,2})/i);
    if (match && dia) {
      const ano = Number(match[1]);
      const semanaNum = Number(match[2]);
      const segunda = this.dataInicioSemanaISO(ano, semanaNum);

      const mapa = {
        "segunda": 0,
        "terca": 1,
        "terça": 1,
        "quarta": 2,
        "quinta": 3,
        "sexta": 4,
        "sabado": 5,
        "sábado": 5,
        "domingo": 6
      };

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
    this.setTextoPossivel(["relTotalRefeicoes", "statRelRefeicoes"], resultado.totais.total);
    this.setTextoPossivel(["relCustoVascon", "statRelVascon"], this.moeda(resultado.totais.valorVascon));
    this.setTextoPossivel(["relDescontoFuncionarios", "statRelDesconto"], this.moeda(resultado.totais.desconto));
    this.setTextoPossivel(["relDiferencaNf", "statRelDiferenca"], resultado.totais.diferencaNf === null ? "-" : this.moeda(resultado.totais.diferencaNf));

    const op = resultado.totais.opcoes;
    this.setTextoPossivel(["relPrincipal"], op.Principal);
    this.setTextoPossivel(["relLight", "relLigth"], op.Light);
    this.setTextoPossivel(["relCarne"], op.Carne);
    this.setTextoPossivel(["relMassa"], op.Massa);
    this.setTextoPossivel(["relLanche"], op.Lanche);

    this.renderizarTabelaPrincipal(resultado);
  },

  renderizarTabelaPrincipal(resultado) {
    const tipo = resultado.filtros.tipo;
    const area = this.encontrarAreaRelatorios();
    if (!area) return;

    let bloco = document.getElementById("relatorioSelecionadoPreview");

    if (!bloco) {
      bloco = document.createElement("div");
      bloco.id = "relatorioSelecionadoPreview";
      bloco.style.marginTop = "1.5rem";
      area.appendChild(bloco);
    }

    const titulo = this.tipos[tipo] || "Relatório";
    const linhas = this.linhasDoTipo(resultado);

    bloco.innerHTML = `
      <div class="section-title" style="margin-bottom:.8rem;">📌 ${titulo}</div>
      <div class="table-wrap">
        <table class="table">
          <thead>
            <tr>
              ${(linhas[0] ? Object.keys(linhas[0]) : ["Sem dados"]).map(c => `<th>${c}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${
              linhas.length
                ? linhas.map(l => `<tr>${Object.values(l).map(v => `<td>${this.valorTela(v)}</td>`).join("")}</tr>`).join("")
                : `<tr><td style="text-align:center;padding:1.5rem;">Sem dados no período.</td></tr>`
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

  exportarExcel() {
    if (!this.ultimoResultado) {
      this.gerar().then(() => this.exportarExcel());
      return;
    }

    const resultado = this.ultimoResultado;
    const tipo = resultado.filtros.tipo;
    const titulo = this.tipos[tipo] || "Relatório";
    const linhas = this.linhasDoTipo(resultado);

    const html = this.gerarHtmlExcel(titulo, resultado, linhas);
    const blob = new Blob(["\ufeff" + html], { type: "application/vnd.ms-excel;charset=utf-8;" });

    const nomeArquivo =
      `relatorio-refeitorio-${this.slug(titulo)}-${resultado.filtros.dataInicio}-a-${resultado.filtros.dataFim}.xls`;

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = nomeArquivo;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  },

  gerarHtmlExcel(titulo, resultado, linhas) {
    const colunas = linhas[0] ? Object.keys(linhas[0]) : ["Sem dados"];

    const linhaResumo = `
      <tr><td class="resumo-label">Total de refeições</td><td>${resultado.totais.total}</td></tr>
      <tr><td class="resumo-label">Custo Vascon estimado</td><td>${this.moeda(resultado.totais.valorVascon)}</td></tr>
      <tr><td class="resumo-label">Desconto funcionários</td><td>${this.moeda(resultado.totais.desconto)}</td></tr>
      <tr><td class="resumo-label">Valor unitário Vascon</td><td>${this.moeda(resultado.valoresPeriodo.valorVascon)}</td></tr>
      <tr><td class="resumo-label">Valor unitário descontado</td><td>${this.moeda(resultado.valoresPeriodo.valorDesconto)}</td></tr>
    `;

    return `
      <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body {
            font-family: Arial, sans-serif;
          }
          .titulo {
            background: #0B1F3C;
            color: #FFFFFF;
            font-size: 18px;
            font-weight: bold;
            text-align: center;
          }
          .periodo {
            background: #C0281C;
            color: #FFFFFF;
            font-weight: bold;
            text-align: center;
          }
          .sub {
            background: #E8F0FF;
            color: #0B1F3C;
            font-weight: bold;
          }
          .resumo-label {
            font-weight: bold;
            background: #F3F6FB;
          }
          th {
            background: #0B1F3C;
            color: #FFFFFF;
            font-weight: bold;
            border: 1px solid #0B1F3C;
            text-align: center;
          }
          td {
            border: 1px solid #D9E2F3;
            padding: 6px;
          }
          .linha-par {
            background: #F3F6FB;
          }
          .linha-impar {
            background: #FFFFFF;
          }
        </style>
      </head>
      <body>
        <table>
          <tr><td class="titulo" colspan="${Math.max(colunas.length, 6)}">RELATÓRIO REFEITÓRIO HOMY — ${titulo.toUpperCase()}</td></tr>
          <tr><td class="periodo" colspan="${Math.max(colunas.length, 6)}">Período: ${this.dataBR(resultado.filtros.dataInicio)} a ${this.dataBR(resultado.filtros.dataFim)}</td></tr>
          <tr><td class="sub" colspan="${Math.max(colunas.length, 6)}">Gerado em: ${this.dataBR(new Date())}</td></tr>
          <tr></tr>
          ${linhaResumo}
          <tr></tr>
          <tr>${colunas.map(c => `<th>${c}</th>`).join("")}</tr>
          ${
            linhas.length
              ? linhas.map((l, idx) => `
                <tr class="${idx % 2 === 0 ? "linha-par" : "linha-impar"}">
                  ${colunas.map(c => `<td>${this.valorExcel(l[c])}</td>`).join("")}
                </tr>
              `).join("")
              : `<tr><td colspan="${colunas.length}">Sem dados no período.</td></tr>`
          }
        </table>
      </body>
      </html>
    `;
  },

  setTextoPossivel(ids, valor) {
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

  valorExcel(v) {
    if (typeof v === "number") return String(v).replace(".", ",");
    return String(v ?? "");
  },

  moeda(v) {
    return Number(v || 0).toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL"
    });
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
    const dias = [
      "Domingo",
      "Segunda-feira",
      "Terça-feira",
      "Quarta-feira",
      "Quinta-feira",
      "Sexta-feira",
      "Sábado"
    ];

    return dias[new Date(data).getDay()];
  },

  normalizar(valor) {
    return String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  },

  slug(valor) {
    return this.normalizar(valor)
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "");
  }
};

// Compatibilidade com botões antigos
window.gerarRelatorio = () => AdminRelatorios.gerar();
window.exportarCSV = () => AdminRelatorios.exportarExcel();
window.exportarRelatorioCSV = () => AdminRelatorios.exportarExcel();
window.exportarExcel = () => AdminRelatorios.exportarExcel();

document.addEventListener("DOMContentLoaded", () => {
  setTimeout(() => AdminRelatorios.init(), 800);
});
