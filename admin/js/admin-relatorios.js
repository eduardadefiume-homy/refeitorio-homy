// ============================================================
// admin-relatorios.js — Relatórios gerenciais + exportação XLSX
// Prioridade: relatório bonito, por período e por tipo
// ============================================================

window.AdminRelatorios = {
  dados: {
    pedidos: [],
    colaboradores: [],
    valores: [],
    filtrados: []
  },

  ultimoResultado: null,

  tiposValidos: {
    "resumo-dia": "Resumo por dia",
    "centro-custo": "Por centro de custo",
    "colaborador": "Quantidade por colaborador",
    "completo": "Completo gerencial"
  },

  async carregarBase() {
    if (!window.SP) throw new Error("SP não encontrado.");

    const semanaAtual = typeof getSemanaId === "function" ? getSemanaId() : null;

    const [pedidos, colaboradores, valores] = await Promise.all([
      this.buscarPedidos(),
      this.buscarColaboradores(),
      this.buscarValores()
    ]);

    this.dados.pedidos = pedidos || [];
    this.dados.colaboradores = colaboradores || [];
    this.dados.valores = valores || [];

    return this.dados;
  },

  async buscarPedidos() {
    try {
      if (typeof SP.getItems === "function") return await SP.getItems("Pedidos");
      if (typeof SP.getPedidos === "function") {
        const semana = typeof getSemanaId === "function" ? getSemanaId() : "";
        return await SP.getPedidos(semana);
      }
      return [];
    } catch (erro) {
      console.error("Erro ao buscar pedidos:", erro);
      return [];
    }
  },

  async buscarColaboradores() {
    try {
      if (typeof SP.getTodosColaboradores === "function") return await SP.getTodosColaboradores();
      if (typeof SP.getColaboradores === "function") return await SP.getColaboradores();
      return [];
    } catch (erro) {
      console.error("Erro ao buscar colaboradores:", erro);
      return [];
    }
  },

  async buscarValores() {
    try {
      if (typeof SP.getValoresRefeicao === "function") return await SP.getValoresRefeicao();
      return [];
    } catch (erro) {
      console.error("Erro ao buscar valores:", erro);
      return [];
    }
  },

  getFiltros() {
    const hoje = new Date();
    const primeiroDia = new Date(hoje.getFullYear(), hoje.getMonth(), 1);

    const dataInicial =
      document.getElementById("relDataInicio")?.value ||
      document.getElementById("dataInicioRelatorio")?.value ||
      document.querySelector('input[type="date"]')?.value ||
      this.toInputDate(primeiroDia);

    const dateInputs = Array.from(document.querySelectorAll('input[type="date"]'));
    const dataFinal =
      document.getElementById("relDataFim")?.value ||
      document.getElementById("dataFimRelatorio")?.value ||
      dateInputs[1]?.value ||
      this.toInputDate(hoje);

    const tipo =
      document.getElementById("tipoRelatorio")?.value ||
      "resumo-dia";

    const nfVascon =
      Number(String(
        document.getElementById("nfVasconRecebida")?.value ||
        document.getElementById("valorNfVascon")?.value ||
        "0"
      ).replace(",", ".")) || 0;

    return { dataInicial, dataFinal, tipo, nfVascon };
  },

  async gerar() {
    try {
      await this.carregarBase();

      const filtros = this.getFiltros();
      const pedidos = this.filtrarPedidosPorPeriodo(filtros.dataInicial, filtros.dataFinal);

      const valorPeriodo = this.obterValorPeriodo(filtros.dataInicial, filtros.dataFinal);
      const resultado = this.montarResultado(filtros, pedidos, valorPeriodo);

      this.ultimoResultado = resultado;
      this.renderizarTela(resultado);

      return resultado;
    } catch (erro) {
      console.error("Erro ao gerar relatório:", erro);
      alert(`Erro ao gerar relatório: ${erro.message || erro}`);
    }
  },

  filtrarPedidosPorPeriodo(dataInicial, dataFinal) {
    const ini = new Date(`${dataInicial}T00:00:00`);
    const fim = new Date(`${dataFinal}T23:59:59`);

    return this.dados.pedidos.filter(p => {
      const status = String(p.Status || "").toLowerCase();
      if (status === "cancelado" || status === "bloqueado") return false;

      const dataPedido = this.obterDataPedido(p);
      if (!dataPedido) return false;

      return dataPedido >= ini && dataPedido <= fim;
    });
  },

  obterDataPedido(pedido) {
    // Preferir Data_Hora real.
    if (pedido.Data_Hora) {
      const d = new Date(pedido.Data_Hora);
      if (!Number.isNaN(d.getTime())) return d;
    }

    // Fallback: se existir campo Data.
    if (pedido.Data) {
      const d = new Date(pedido.Data);
      if (!Number.isNaN(d.getTime())) return d;
    }

    // Último fallback: tenta usar Dia + Semana_id.
    return null;
  },

  obterValorPeriodo(dataInicial, dataFinal) {
    const ini = new Date(`${dataInicial}T00:00:00`);
    const fim = new Date(`${dataFinal}T23:59:59`);

    const ativos = this.dados.valores.filter(v => {
      const ativo = typeof SP?.isTrue === "function" ? SP.isTrue(v.Ativo) : String(v.Ativo).toLowerCase() !== "false";
      if (!ativo) return false;

      const vi = v.Data_Inicio ? new Date(v.Data_Inicio) : null;
      const vf = v.Data_Fim ? new Date(v.Data_Fim) : null;

      if (!vi || !vf) return true;
      return vi <= fim && vf >= ini;
    });

    const escolhido = ativos[0] || this.dados.valores[0] || {};

    return {
      valorVascon: Number(escolhido.Valor_Vascon || escolhido.valorVascon || 0),
      valorDesconto: Number(escolhido.Valor_Desconto_Funcionario || escolhido.valorDesconto || 0),
      periodo: escolhido.Title || ""
    };
  },

  montarResultado(filtros, pedidos, valorPeriodo) {
    const porDia = this.montarResumoPorDia(filtros.dataInicial, filtros.dataFinal, pedidos);
    const porCentroCusto = this.montarPorCentroCusto(pedidos, valorPeriodo);
    const porColaborador = this.montarPorColaborador(pedidos, valorPeriodo);
    const totaisOpcao = this.contarPorOpcao(pedidos);

    const totalRefeicoes = pedidos.length;
    const custoVascon = totalRefeicoes * valorPeriodo.valorVascon;
    const descontoFuncionarios = totalRefeicoes * valorPeriodo.valorDesconto;

    return {
      filtros,
      valorPeriodo,
      pedidos,
      porDia,
      porCentroCusto,
      porColaborador,
      totaisOpcao,
      totais: {
        totalRefeicoes,
        custoVascon,
        descontoFuncionarios,
        diferencaNf: filtros.nfVascon ? filtros.nfVascon - custoVascon : null
      }
    };
  },

  montarResumoPorDia(dataInicial, dataFinal, pedidos) {
    const linhas = [];
    const ini = new Date(`${dataInicial}T00:00:00`);
    const fim = new Date(`${dataFinal}T00:00:00`);

    for (let d = new Date(ini); d <= fim; d.setDate(d.getDate() + 1)) {
      const dataChave = this.toInputDate(d);
      const pedidosDia = pedidos.filter(p => {
        const dataPedido = this.obterDataPedido(p);
        return dataPedido && this.toInputDate(dataPedido) === dataChave;
      });

      const contagem = this.contarPorOpcao(pedidosDia);

      linhas.push({
        Dia: this.diaSemanaCurto(d),
        Data: this.dataBR(d),
        "Dia/Data": `${this.diaSemanaCurto(d)} - ${this.dataBR(d)}`,
        Principal: contagem.Principal,
        Light: contagem.Light,
        Carne: contagem.Carne,
        Massa: contagem.Massa,
        Lanche: contagem.Lanche,
        Total: pedidosDia.length
      });
    }

    return linhas;
  },

  montarPorCentroCusto(pedidos, valorPeriodo) {
    const mapa = {};

    for (const p of pedidos) {
      const centro = p.Centro_Custo || this.centroCustoColaborador(p.Colaborador_id, p.Colaborador_nome) || "Sem centro de custo";

      if (!mapa[centro]) {
        mapa[centro] = {
          "Centro de Custo": centro,
          Quantidade: 0,
          "Valor Vascon": 0,
          "Desconto Funcionários": 0
        };
      }

      mapa[centro].Quantidade += 1;
      mapa[centro]["Valor Vascon"] += valorPeriodo.valorVascon;
      mapa[centro]["Desconto Funcionários"] += valorPeriodo.valorDesconto;
    }

    return Object.values(mapa);
  },

  montarPorColaborador(pedidos, valorPeriodo) {
    const mapa = {};

    for (const p of pedidos) {
      const nome = p.Colaborador_nome || p.Colaborador || "Sem nome";
      const centro = p.Centro_Custo || this.centroCustoColaborador(p.Colaborador_id, nome) || "Sem centro de custo";
      const chave = `${nome}|${centro}`;

      if (!mapa[chave]) {
        mapa[chave] = {
          Colaborador: nome,
          "Centro de Custo": centro,
          Quantidade: 0,
          "Valor unitário descontado": valorPeriodo.valorDesconto,
          "Desconto em folha": 0,
          "Valor Vascon estimado": 0
        };
      }

      mapa[chave].Quantidade += 1;
      mapa[chave]["Desconto em folha"] += valorPeriodo.valorDesconto;
      mapa[chave]["Valor Vascon estimado"] += valorPeriodo.valorVascon;
    }

    return Object.values(mapa).sort((a, b) => a.Colaborador.localeCompare(b.Colaborador));
  },

  centroCustoColaborador(id, nome) {
    const colab = this.dados.colaboradores.find(c =>
      String(c.id) === String(id) ||
      String(c.ID) === String(id) ||
      this.normalizar(c.Nome || c.Title) === this.normalizar(nome)
    );

    return colab ? colab.Centro_Custo : "";
  },

  contarPorOpcao(pedidos) {
    const base = { Principal: 0, Light: 0, Carne: 0, Massa: 0, Lanche: 0 };

    for (const p of pedidos) {
      const opcao = this.normalizar(p.Opcao || p.opcao || "");
      if (opcao.includes("principal")) base.Principal++;
      else if (opcao.includes("light") || opcao.includes("ligth")) base.Light++;
      else if (opcao.includes("carne")) base.Carne++;
      else if (opcao.includes("massa")) base.Massa++;
      else if (opcao.includes("lanche")) base.Lanche++;
    }

    return base;
  },

  renderizarTela(resultado) {
    this.setTexto(["relTotalRefeicoes", "statRelRefeicoes"], resultado.totais.totalRefeicoes);
    this.setTexto(["relCustoVascon", "statRelVascon"], this.moeda(resultado.totais.custoVascon));
    this.setTexto(["relDescontoFuncionarios", "statRelDesconto"], this.moeda(resultado.totais.descontoFuncionarios));

    if (resultado.totais.diferencaNf !== null) {
      this.setTexto(["relDiferencaNf", "statRelDiferenca"], this.moeda(resultado.totais.diferencaNf));
    }

    const op = resultado.totaisOpcao;
    this.setTexto(["relPrincipal"], op.Principal);
    this.setTexto(["relLight", "relLigth"], op.Light);
    this.setTexto(["relCarne"], op.Carne);
    this.setTexto(["relMassa"], op.Massa);
    this.setTexto(["relLanche"], op.Lanche);

    this.renderTabela("tbodyResumoDia", resultado.porDia);
    this.renderTabela("tbodyCentroCusto", resultado.porCentroCusto);
    this.renderTabela("tbodyColaboradorRelatorio", resultado.porColaborador);
  },

  renderTabela(id, linhas) {
    const tbody = document.getElementById(id);
    if (!tbody) return;

    if (!linhas || !linhas.length) {
      tbody.innerHTML = `<tr><td colspan="10" style="text-align:center;opacity:.55;padding:1.5rem;">Sem dados no período.</td></tr>`;
      return;
    }

    const cols = Object.keys(linhas[0]);
    tbody.innerHTML = linhas.map(l => `
      <tr>
        ${cols.map(c => `<td>${this.formatarValorTabela(l[c])}</td>`).join("")}
      </tr>
    `).join("");
  },

  formatarValorTabela(valor) {
    if (typeof valor === "number") return String(valor).replace(".", ",");
    return String(valor ?? "");
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

  exportarExcel() {
    const resultado = this.ultimoResultado || null;

    if (!resultado) {
      this.gerar().then(r => {
        if (r) this.exportarExcelComResultado(r);
      });
      return;
    }

    this.exportarExcelComResultado(resultado);
  },

  exportarExcelComResultado(resultado) {
    if (!window.XLSX) {
      alert("Biblioteca XLSX não carregada. Verifique a conexão ou o script SheetJS.");
      return;
    }

    const tipo = resultado.filtros.tipo;
    const wb = XLSX.utils.book_new();

    if (tipo === "resumo-dia") {
      this.adicionarAbaBonita(wb, "Resumo por dia", resultado, resultado.porDia);
    } else if (tipo === "centro-custo") {
      this.adicionarAbaBonita(wb, "Centro de custo", resultado, resultado.porCentroCusto);
    } else if (tipo === "colaborador") {
      this.adicionarAbaBonita(wb, "Por colaborador", resultado, resultado.porColaborador);
    } else {
      this.adicionarAbaBonita(wb, "Resumo por dia", resultado, resultado.porDia);
      this.adicionarAbaBonita(wb, "Centro de custo", resultado, resultado.porCentroCusto);
      this.adicionarAbaBonita(wb, "Por colaborador", resultado, resultado.porColaborador);
    }

    const nomeTipo = this.tiposValidos[tipo] || "Relatorio";
    const nomeArquivo = `relatorio-refeitorio-${this.slug(nomeTipo)}-${resultado.filtros.dataInicial}-a-${resultado.filtros.dataFinal}.xlsx`;

    XLSX.writeFile(wb, nomeArquivo);
  },

  adicionarAbaBonita(wb, nomeAba, resultado, linhas) {
    const titulo = `RELATÓRIO REFEITÓRIO HOMY - ${nomeAba.toUpperCase()}`;
    const periodo = `Período: ${this.dataBR(resultado.filtros.dataInicial)} a ${this.dataBR(resultado.filtros.dataFinal)}`;
    const gerado = `Gerado em: ${this.dataBR(new Date())}`;

    const dados = [
      [titulo],
      [periodo],
      [gerado],
      [],
      ["Total de refeições", resultado.totais.totalRefeicoes],
      ["Custo Vascon estimado", resultado.totais.custoVascon],
      ["Desconto funcionários", resultado.totais.descontoFuncionarios],
      ["Valor unitário Vascon", resultado.valorPeriodo.valorVascon],
      ["Valor unitário descontado", resultado.valorPeriodo.valorDesconto],
      [],
      ...(linhas && linhas.length ? [Object.keys(linhas[0])] : [["Sem dados"]]),
      ...(linhas && linhas.length ? linhas.map(l => Object.values(l)) : [])
    ];

    const ws = XLSX.utils.aoa_to_sheet(dados);

    const range = XLSX.utils.decode_range(ws["!ref"]);

    // Larguras automáticas
    const colWidths = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      let max = 10;
      for (let R = range.s.r; R <= range.e.r; ++R) {
        const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
        if (cell && cell.v !== undefined) {
          max = Math.max(max, String(cell.v).length + 2);
        }
      }
      colWidths.push({ wch: Math.min(Math.max(max, 12), 35) });
    }
    ws["!cols"] = colWidths;

    // Merge título
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(4, range.e.c) } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: Math.max(4, range.e.c) } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: Math.max(4, range.e.c) } }
    ];

    // Estilos — podem ser lidos por Excel moderno; se alguma lib ignorar, os dados continuam corretos.
    const azulHomy = "0B1F3C";
    const vermelhoHomy = "C0281C";
    const branco = "FFFFFF";
    const cinzaClaro = "E8F0FF";

    const aplicarEstilo = (addr, style) => {
      if (!ws[addr]) return;
      ws[addr].s = style;
    };

    aplicarEstilo("A1", {
      font: { bold: true, color: { rgb: branco }, sz: 16 },
      fill: { fgColor: { rgb: azulHomy } },
      alignment: { horizontal: "center" }
    });

    aplicarEstilo("A2", {
      font: { bold: true, color: { rgb: branco } },
      fill: { fgColor: { rgb: vermelhoHomy } },
      alignment: { horizontal: "center" }
    });

    aplicarEstilo("A3", {
      font: { italic: true, color: { rgb: azulHomy } }
    });

    const headerRow = 10; // zero-based: linha 11 do Excel
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const addr = XLSX.utils.encode_cell({ r: headerRow, c: C });
      aplicarEstilo(addr, {
        font: { bold: true, color: { rgb: branco } },
        fill: { fgColor: { rgb: azulHomy } },
        alignment: { horizontal: "center" },
        border: {
          top: { style: "thin", color: { rgb: azulHomy } },
          bottom: { style: "thin", color: { rgb: azulHomy } },
          left: { style: "thin", color: { rgb: azulHomy } },
          right: { style: "thin", color: { rgb: azulHomy } }
        }
      });
    }

    for (let R = headerRow + 1; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        aplicarEstilo(addr, {
          fill: { fgColor: { rgb: R % 2 === 0 ? "F3F6FB" : "FFFFFF" } },
          border: {
            bottom: { style: "thin", color: { rgb: "D9E2F3" } }
          }
        });
      }
    }

    XLSX.utils.book_append_sheet(wb, ws, nomeAba.substring(0, 31));
  },

  toInputDate(data) {
    const d = new Date(data);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  },

  dataBR(data) {
    if (!data) return "";
    if (typeof data === "string" && /^\d{4}-\d{2}-\d{2}$/.test(data)) {
      const [y, m, d] = data.split("-");
      return `${d}/${m}/${y}`;
    }

    const dt = new Date(data);
    if (Number.isNaN(dt.getTime())) return "";
    return dt.toLocaleDateString("pt-BR");
  },

  diaSemanaCurto(data) {
    const dias = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SÁB"];
    return dias[new Date(data).getDay()];
  },

  moeda(valor) {
    return Number(valor || 0).toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL"
    });
  },

  normalizar(valor) {
    return String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  },

  slug(valor) {
    return this.normalizar(valor).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  }
};

// Compatibilidade com botões antigos
window.gerarRelatorio = () => AdminRelatorios.gerar();
window.exportarCSV = () => AdminRelatorios.exportarExcel();
window.exportarExcel = () => AdminRelatorios.exportarExcel();
window.exportarRelatorioCSV = () => AdminRelatorios.exportarExcel();

document.addEventListener("DOMContentLoaded", () => {
  setTimeout(() => {
    const texto = document.body.innerText.toLowerCase();
    if (texto.includes("relatórios") || texto.includes("relatorios")) {
      AdminRelatorios.gerar();
    }
  }, 1000);
});
