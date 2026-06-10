// ============================================================
// admin-relatorios.js
// Versão: 2026-06-10
// Melhorias:
//  - CC exibido como "120501 - TI" (busca Departamento do colaborador)
//  - Coluna Conta_Contabil no relatório de CC (vem da lista Colaboradores)
//  - Desconto por colaborador calculado sobre qtd real de refeições
//  - Aba Rateio Vascon no Excel no padrão enviado pela Luana
//  - Exportação com duas abas: Relatório Principal + Rateio Vascon
// ============================================================

(function () {

  // ── Mapa completo CC → { descricao, conta } ─────────────────────────────────
  // Fonte: modelo de rateio fornecido pela Luana. Conta_Contabil vem do SharePoint
  // se preenchida; caso não esteja, usamos este fallback.
  const CC_MAP = {
    "110101": { descricao: "DIRETORIA PRESIDENCIAL",                  conta: "51101015" },
    "110201": { descricao: "DIRETORIA ADMINISTRATIVA",                conta: "51101015" },
    "110202": { descricao: "DIRETORIA DE PRODUTOS",                   conta: "51101015" },
    "120101": { descricao: "ADM GERAL",                               conta: "51101015" },
    "120102": { descricao: "CUSTOS",                                  conta: "51101015" },
    "120103": { descricao: "LEGALIZAÇÃO",                             conta: "51101015" },
    "120201": { descricao: "CONTABILIDADE",                           conta: "51101015" },
    "120202": { descricao: "FISCAL",                                  conta: "51101015" },
    "120301": { descricao: "FINANCEIRO",                              conta: "51101015" },
    "120401": { descricao: "RECURSOS HUMANOS",                        conta: "51101015" },
    "120402": { descricao: "DEPARTAMENTO PESSOAL",                    conta: "51101015" },
    "120501": { descricao: "TI",                                      conta: "51101015" },
    "120601": { descricao: "RECEPÇÃO",                                conta: "51101015" },
    "120602": { descricao: "PORTARIA",                                conta: "51101015" },
    "120603": { descricao: "ASSEIO E CONSERVAÇÃO",                    conta: "51101015" },
    "120604": { descricao: "JARDINAGEM",                              conta: "51101015" },
    "150101": { descricao: "SUPRIMENTOS",                             conta: "51101015" },
    "160101": { descricao: "CONTROLADORIA E COMPLIANCE",              conta: "51101015" },
    "160102": { descricao: "ADM CONTRATOS",                           conta: "51101015" },
    "170101": { descricao: "SGI",                                     conta: "51101015" },
    "180101": { descricao: "P&D",                                     conta: "51101015" },
    "190101": { descricao: "PATIO EXTERNO",                           conta: "51101015" },
    "220101": { descricao: "ADM VENDAS",                              conta: "51101015" },
    "220201": { descricao: "COML INTERNO - SUPORTE",                  conta: "51101015" },
    "220202": { descricao: "COML INTERNO - ATIVO",                    conta: "51101015" },
    "220301": { descricao: "COML EXTERNO - CLT",                      conta: "51101015" },
    "220302": { descricao: "COML EXTERNO - REPRESENTANTE",            conta: "51101015" },
    "230101": { descricao: "SUPORTE TECNICO INDUSTRIAL",              conta: "51101015" },
    "230102": { descricao: "SUPORTE TECNICO OBRAS/INFRA",             conta: "51101015" },
    "240101": { descricao: "MARKETING",                               conta: "51101015" },
    "250101": { descricao: "FATURAMENTO",                             conta: "51101015" },
    "250102": { descricao: "LOGISTICA",                               conta: "51101015" },
    "250103": { descricao: "EXPEDIÇÃO",                               conta: "51101015" },
    "320101": { descricao: "PRODUÇÃO",                                conta: "42101015" },
    "320201": { descricao: "ENVASE MANUAL",                           conta: "42101015" },
    "320202": { descricao: "ENVASE AUTOMATICO",                       conta: "42101015" },
    "320301": { descricao: "LABORATORIO E CONTROLE QUALIDADE",        conta: "42101015" },
    "360101": { descricao: "APOIO A PRODUÇÃO",                        conta: "42101015" },
    "360102": { descricao: "PCP",                                     conta: "42101015" },
    "360201": { descricao: "MANUTENÇÃO",                              conta: "42101015" },
    "360301": { descricao: "ALMOXARIFADO DE INSUMOS",                 conta: "42101015" }
  };

  // Retorna "120501 - TI" ou só "120501" se não encontrado
  function ccComNome(cc, colaboradorData) {
    const cod = String(cc || "").trim();
    if (!cod) return "Sem CC";

    // Tenta pegar departamento do colaborador (mais preciso)
    if (colaboradorData) {
      const dept = colaboradorData.Departamento || colaboradorData.departamento || "";
      if (dept) return `${cod} - ${dept.toUpperCase()}`;
    }

    // Fallback no mapa
    const info = CC_MAP[cod];
    return info ? `${cod} - ${info.descricao}` : cod;
  }

  function contaContabil(cc, colaboradorData) {
    const cod = String(cc || "").trim();
    // Tenta pegar Conta_Contabil do colaborador
    if (colaboradorData) {
      const conta = colaboradorData.Conta_Contabil || colaboradorData.conta_contabil || "";
      if (conta) return String(conta).trim();
    }
    const info = CC_MAP[cod];
    return info ? info.conta : "51101015";
  }

  const DIA_LABEL = {
    segunda: "Segunda-feira",
    terca:   "Terça-feira",
    quarta:  "Quarta-feira",
    quinta:  "Quinta-feira",
    sexta:   "Sexta-feira",
    sabado:  "Sábado",
    domingo: "Domingo"
  };

  const TIPOS = {
    "resumo-dia":   "Resumo por dia",
    "centro-custo": "Valor total por centro de custo",
    "colaborador":  "Quantidade por colaborador"
  };

  const COLUNAS_TIPO = {
    "resumo-dia":   ["Dia", "Data", "Principal", "Light", "Carne", "Massa", "Lanche", "Total"],
    "centro-custo": ["Centro de custo", "Descrição", "Quantidade", "Valor Vascon", "Desconto funcionários", "Rateio NF"],
    "colaborador":  ["Colaborador", "Centro de custo", "Quantidade", "Valor unitário desconto", "Total desconto em folha", "Valor Vascon estimado"]
  };

  const state = {
    pedidos:       [],
    colaboradores: [],
    valores:       [],
    resultado:     null
  };

  // ── Cache de colaboradores por id e por nome ──────────────────────────────
  let _colabById   = {};
  let _colabByNome = {};

  function buildColabCache() {
    _colabById   = {};
    _colabByNome = {};
    (state.colaboradores || []).forEach(c => {
      const id   = String(c.id || c.ID || "");
      const nome = normalizar(c.Nome || c.Title || "");
      if (id)   _colabById[id]   = c;
      if (nome) _colabByNome[nome] = c;
    });
  }

  function getColabData(pedido) {
    const id   = String(pedido.Colaborador_id || pedido.colaborador_id || "");
    const nome = normalizar(pedido.Colaborador_nome || pedido.colaborador_nome || "");
    return _colabById[id] || _colabByNome[nome] || null;
  }

  function $(id) { return document.getElementById(id); }

  function normalizar(valor) {
    return String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  }

  function moeda(valor) {
    return Number(valor || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
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
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
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
    if (window.AdminState && AdminState.getSemanaId) return AdminState.getSemanaId();
    const d = new Date(); d.setHours(0,0,0,0);
    d.setDate(d.getDate()+3-((d.getDay()+6)%7));
    const w1 = new Date(d.getFullYear(),0,4);
    const wn = 1+Math.round(((d-w1)/86400000-3+((w1.getDay()+6)%7))/7);
    return `${d.getFullYear()}-W${String(wn).padStart(2,"0")}`;
  }

  function inicioSemanaISO(ano, semana) {
    const jan4  = new Date(ano, 0, 4);
    const start = new Date(jan4);
    start.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1 + (semana-1)*7);
    return start;
  }

  function diaSemanaCompleto(d) {
    return ["Domingo","Segunda-feira","Terça-feira","Quarta-feira","Quinta-feira","Sexta-feira","Sábado"][d.getDay()];
  }

  function dataPedido(pedido) {
    const dh = pedido.Data_Hora || pedido.data_hora || pedido.DataHora;
    if (dh) {
      const d = new Date(dh);
      if (!Number.isNaN(d.getTime())) return d;
    }

    const semId  = pedido.Semana_id || pedido.semana_id || "";
    const dia    = pedido.Dia || pedido.dia || "";
    const match  = semId.match(/(\d{4})-W(\d{1,2})/i);

    if (match && dia) {
      const segunda = inicioSemanaISO(Number(match[1]), Number(match[2]));
      const mapa = { segunda:0, terca:1, "terça":1, quarta:2, quinta:3, sexta:4, sabado:5, "sábado":5, domingo:6 };
      const idx  = mapa[normalizar(dia)];
      if (idx !== undefined) {
        const d = new Date(segunda);
        d.setDate(d.getDate() + idx);
        return d;
      }
    }
    return null;
  }

  function diaChave(data) {
    return ["domingo","segunda","terca","quarta","quinta","sexta","sabado"][new Date(data).getDay()];
  }

  function getStatusPedido(pedido) {
    return normalizar(pedido.Status || pedido.status || "");
  }

  function pedidoConta(pedido) {
    const status = getStatusPedido(pedido);
    if (["cancelado","bloqueado","afastado","ferias","férias","nao vai almocar","não vai almoçar"].includes(status)) return false;
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
    return pedido.Colaborador_nome || pedido.colaborador_nome || pedido.Colaborador || pedido.colaborador || pedido.Title || "Sem nome";
  }

  function getColaboradorId(pedido) {
    return String(pedido.Colaborador_id || pedido.colaborador_id || pedido.ColaboradorId || pedido.colaboradorId || "");
  }

  function getCentroCustoRaw(pedido) {
    const direto = pedido.Centro_Custo || pedido.centro_custo || pedido.CentroCusto || pedido.centroCusto;
    if (direto) return String(direto).trim();
    const colab = getColabData(pedido);
    return String(colab?.Centro_Custo || colab?.centro_custo || "").trim() || "Sem CC";
  }

  function valorNumero(valor) {
    return Number(String(valor || "0").replace(",", ".")) || 0;
  }

  async function safe(fn, fallback = []) {
    try { return await fn(); }
    catch (erro) { console.warn("Falha no relatório:", erro); return fallback; }
  }

  async function carregarDados(semanaId) {
    if (!window.SP) throw new Error("SP não carregado.");

    state.colaboradores = await safe(() =>
      SP.getTodosColaboradores ? SP.getTodosColaboradores() : SP.getColaboradores(), []
    );

    buildColabCache();

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
      const vf = v.Data_Fim    ? new Date(v.Data_Fim)    : null;
      if (!ativo) return false;
      if (!vi || !vf) return true;
      return vi <= fim && vf >= ini;
    });

    const v = ativos[0] || state.valores[0] || {};
    return {
      valorVascon:   valorNumero(v.Valor_Vascon   || v.valorVascon),
      valorDesconto: valorNumero(v.Valor_Desconto_Funcionario || v.valorDescontoFuncionario || v.valorDesconto)
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

    for (let d = new Date(ini); d <= fim; d.setDate(d.getDate()+1)) {
      const chave    = toInputDate(d);
      const pedidosDia = pedidos.filter(p => {
        const data = dataPedido(p);
        return data && toInputDate(data) === chave;
      });
      const cont = { principal:0, light:0, carne:0, massa:0, lanche:0 };
      pedidosDia.forEach(p => { const op = getOpcao(p); if (cont[op] !== undefined) cont[op]++; });
      linhas.push({
        "Dia":       DIA_LABEL[diaChave(d)] || diaSemanaCompleto(d),
        "Data":      dataBR(d),
        "Principal": cont.principal,
        "Light":     cont.light,
        "Carne":     cont.carne,
        "Massa":     cont.massa,
        "Lanche":    cont.lanche,
        "Total":     pedidosDia.length
      });
    }
    return linhas;
  }

  function montarCentroCusto(pedidos, valores, nf) {
    const mapa = {};

    pedidos.forEach(p => {
      const ccRaw   = getCentroCustoRaw(p);
      const colab   = getColabData(p);
      const ccLabel = ccComNome(ccRaw, colab);
      const conta   = contaContabil(ccRaw, colab);

      if (!mapa[ccRaw]) {
        mapa[ccRaw] = {
          "_ccRaw":              ccRaw,
          "_conta":              conta,
          "Centro de custo":     ccLabel,
          "Descrição":           (CC_MAP[ccRaw]?.descricao || (colab?.Departamento || "").toUpperCase() || ccRaw),
          "Quantidade":          0,
          "Valor Vascon":        0,
          "Desconto funcionários": 0,
          "Rateio NF":           0
        };
      }

      mapa[ccRaw]["Quantidade"]++;
      mapa[ccRaw]["Valor Vascon"]        += valores.valorVascon;
      mapa[ccRaw]["Desconto funcionários"] += valores.valorDesconto;
    });

    const total = pedidos.length || 1;
    Object.values(mapa).forEach(l => {
      l["Rateio NF"] = nf ? nf * (l["Quantidade"] / total) : 0;
    });

    return Object.values(mapa).sort((a, b) => String(a._ccRaw).localeCompare(String(b._ccRaw)));
  }

  function montarColaborador(pedidos, valores) {
    const mapa = {};

    pedidos.forEach(p => {
      const nome  = getNomeColaborador(p);
      const ccRaw = getCentroCustoRaw(p);
      const colab = getColabData(p);
      const ccLabel = ccComNome(ccRaw, colab);
      const key   = `${nome}|${ccRaw}`;

      if (!mapa[key]) {
        mapa[key] = {
          "Colaborador":              nome,
          "Centro de custo":          ccLabel,
          "Quantidade":               0,
          "Valor unitário desconto":  valores.valorDesconto,
          "Total desconto em folha":  0,
          "Valor Vascon estimado":    0
        };
      }

      mapa[key]["Quantidade"]++;
      mapa[key]["Total desconto em folha"] += valores.valorDesconto;
      mapa[key]["Valor Vascon estimado"]   += valores.valorVascon;
    });

    return Object.values(mapa).sort((a, b) =>
      String(a["Centro de custo"]).localeCompare(String(b["Centro de custo"])) ||
      String(a.Colaborador).localeCompare(String(b.Colaborador))
    );
  }

  // ── Rateio Vascon — igual ao modelo impresso ──────────────────────────────
  function montarRateioVascon(pedidos, valores, nf) {
    // Monta por CC com conta contábil
    const mapaCC = {};

    pedidos.forEach(p => {
      const ccRaw = getCentroCustoRaw(p);
      const colab = getColabData(p);
      const info  = CC_MAP[ccRaw];
      const conta = contaContabil(ccRaw, colab);
      const descr = info?.descricao || (colab?.Departamento || "").toUpperCase() || ccRaw;

      if (!mapaCC[ccRaw]) {
        mapaCC[ccRaw] = { conta, cc: ccRaw, descricao: descr, qtde: 0 };
      }
      mapaCC[ccRaw].qtde++;
    });

    const total = pedidos.length || 1;

    // Inclui todos os CCs do mapa (mesmo os com 0) para mostrar rateio completo
    const todos = Object.keys(CC_MAP).reduce((acc, cc) => {
      if (!acc[cc]) {
        acc[cc] = {
          conta:    CC_MAP[cc].conta,
          cc,
          descricao: CC_MAP[cc].descricao,
          qtde:      0
        };
      }
      return acc;
    }, { ...mapaCC });

    return Object.values(todos).sort((a, b) => String(a.cc).localeCompare(String(b.cc))).map(r => ({
      conta:    r.conta,
      cc:       r.cc,
      descricao: r.descricao,
      qtde:     r.qtde,
      soma:     r.qtde * valores.valorVascon,
      pct:      total > 0 ? (r.qtde / total) : 0
    }));
  }

  function atualizarCards(pedidos, valores, filtros) {
    const total    = pedidos.length;
    const custo    = total * valores.valorVascon;
    const desconto = total * valores.valorDesconto;
    const cont = { principal:0, light:0, carne:0, massa:0, lanche:0 };
    pedidos.forEach(p => { const op = getOpcao(p); if (cont[op] !== undefined) cont[op]++; });

    $("relTotalGeral").textContent    = total;
    $("relCustoVascon").textContent   = moeda(custo);
    $("relDescontoFolha").textContent = moeda(desconto);
    $("relDiferencaNF").textContent   = filtros.nf ? moeda(filtros.nf - custo) : "—";

    $("rel-principal").textContent = cont.principal;
    $("rel-light").textContent     = cont.light;
    $("rel-carne").textContent     = cont.carne;
    $("rel-massa").textContent     = cont.massa;
    $("rel-lanche").textContent    = cont.lanche;
  }

  function linhasDoTipo(resultado) {
    if (resultado.filtros.tipo === "centro-custo") return resultado.centroCusto;
    if (resultado.filtros.tipo === "colaborador")  return resultado.colaborador;
    return resultado.resumoDia;
  }

  function colunasDoTipo(tipo, linhas) {
    // Remove chaves internas (_ccRaw, _conta)
    if (linhas && linhas.length) return Object.keys(linhas[0]).filter(k => !k.startsWith("_"));
    return COLUNAS_TIPO[tipo] || ["Sem dados"];
  }

  function valorTela(valor) {
    if (typeof valor === "number") {
      // Monetário ou percentual
      if (valor > 0 && valor < 1) return (valor * 100).toFixed(2).replace(".", ",") + "%";
      return String(valor).replace(".", ",");
    }
    return valor ?? "";
  }

  function renderTabela(resultado) {
    const linhas  = linhasDoTipo(resultado);
    const colunas = colunasDoTipo(resultado.filtros.tipo, linhas);
    const titulo  = TIPOS[resultado.filtros.tipo] || "Relatório";

    $("relTituloTabela").textContent = titulo;
    $("relHead").innerHTML = `<tr>${colunas.map(c => `<th>${c}</th>`).join("")}</tr>`;

    if (!linhas.length) {
      $("relTableFinal").innerHTML = `<tr><td colspan="${colunas.length}" style="text-align:center;padding:2rem;color:rgba(143,170,210,0.4)">Sem dados no período.</td></tr>`;
      return;
    }

    $("relTableFinal").innerHTML = linhas.map(linha =>
      `<tr>${colunas.map(c => `<td>${valorTela(linha[c])}</td>`).join("")}</tr>`
    ).join("");
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
        resumoDia:   montarResumoDia(filtros, pedidos),
        centroCusto: montarCentroCusto(pedidos, valores, filtros.nf),
        colaborador: montarColaborador(pedidos, valores),
        rateio:      montarRateioVascon(pedidos, valores, filtros.nf)
      };

      state.resultado = resultado;
      atualizarCards(pedidos, valores, filtros);
      renderTabela(resultado);
    } catch (erro) {
      console.error("Erro ao gerar relatório final:", erro);
      if (typeof window.toast === "function") toast("Erro ao gerar relatório: " + (erro.message || erro), "error");
      else alert("Erro ao gerar relatório: " + (erro.message || erro));
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
    const b = { style: "thin", color: { argb: color } };
    return { top: b, left: b, bottom: b, right: b };
  }

  // ── Exportação Excel com 2 abas ────────────────────────────────────────────
  async function exportarRelatorioExcelFinal() {
    if (!state.resultado) await gerarRelatorioFinal(semanaAtualFallback());
    const resultado = state.resultado;
    if (!resultado) return;

    await carregarExcelJS();

    const wb     = new ExcelJS.Workbook();
    wb.creator   = "Homy Refeitório";

    // ── ABA 1: Relatório principal ────────────────────────────────────────
    const titulo  = TIPOS[resultado.filtros.tipo] || "Relatório";
    const linhas  = linhasDoTipo(resultado);
    const colunas = colunasDoTipo(resultado.filtros.tipo, linhas);
    const totalCols = Math.max(colunas.length, 6);

    const ws1 = wb.addWorksheet(titulo.substring(0, 31));

    // Cabeçalho
    const addHeader = (ws, text, row, cols, bgArgb, fontArgb) => {
      ws.mergeCells(row, 1, row, cols);
      const cell = ws.getCell(row, 1);
      cell.value     = text;
      cell.font      = { bold: true, size: row === 1 ? 14 : 11, color: { argb: fontArgb } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: bgArgb } };
    };

    addHeader(ws1, `RELATÓRIO REFEITÓRIO HOMY — ${titulo.toUpperCase()}`, 1, totalCols, "FF0B1F3C", "FFFFFFFF");
    addHeader(ws1, `Período: ${dataBR(resultado.filtros.dataInicio)} a ${dataBR(resultado.filtros.dataFim)}`, 2, totalCols, "FFC0281C", "FFFFFFFF");
    addHeader(ws1, `Gerado em: ${new Date().toLocaleDateString("pt-BR")}`, 3, totalCols, "FFE8F1FF", "FF0B1F3C");
    ws1.addRow([]);

    // Resumo numérico
    ws1.addRow(["Total de refeições", resultado.pedidos.length]);
    ws1.addRow(["Valor unitário Vascon", resultado.valores.valorVascon]);
    ws1.addRow(["Valor unitário desconto funcionário", resultado.valores.valorDesconto]);
    if (resultado.filtros.nf) ws1.addRow(["NF Vascon recebida", resultado.filtros.nf]);
    ws1.addRow([]);

    // Header da tabela
    const headerRow = ws1.addRow(colunas);
    headerRow.eachCell(cell => {
      cell.font      = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1F3C" } };
      cell.alignment = { horizontal: "center", wrapText: true };
      cell.border    = borda();
    });
    ws1.getRow(headerRow.number).height = 30;

    const moneyCols = ["Valor Vascon","Desconto funcionários","Rateio NF",
                       "Valor unitário desconto","Total desconto em folha","Valor Vascon estimado"];

    if (linhas.length) {
      linhas.forEach((linha, idx) => {
        const row = ws1.addRow(colunas.map(c => linha[c]));
        row.eachCell((cell, colNumber) => {
          const colName = colunas[colNumber - 1];
          cell.fill   = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? "FFF3F6FB" : "FFFFFFFF" } };
          cell.border = borda("FFD9E2F3");
          if (moneyCols.includes(colName)) {
            cell.numFmt = '"R$" #,##0.00';
          } else if (colName === "Quantidade" || colName === "Total") {
            cell.alignment = { horizontal: "center" };
          }
        });
      });

      // Linha de total
      if (resultado.filtros.tipo !== "resumo-dia") {
        ws1.addRow([]);
        const totRow = ws1.addRow(["TOTAL", resultado.pedidos.length,
          resultado.pedidos.length * resultado.valores.valorVascon,
          resultado.pedidos.length * resultado.valores.valorDesconto,
          resultado.filtros.nf || 0
        ].slice(0, colunas.length));
        totRow.eachCell(cell => {
          cell.font   = { bold: true, color: { argb: "FFFFFFFF" } };
          cell.fill   = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC0281C" } };
          cell.border = borda();
          cell.numFmt = '"R$" #,##0.00';
        });
      }
    } else {
      const row = ws1.addRow(["Sem dados no período."]);
      row.getCell(1).font = { italic: true, color: { argb: "FF64748B" } };
    }

    ws1.columns.forEach(col => {
      let max = 12;
      col.eachCell({ includeEmpty: true }, cell => {
        max = Math.max(max, String(cell.value || "").length + 2);
      });
      col.width = Math.min(max, 42);
    });

    // ── ABA 2: Rateio Vascon ──────────────────────────────────────────────
    const ws2 = wb.addWorksheet("Rateio Vascon");

    // Título
    ws2.mergeCells(1, 1, 1, 6);
    const t1 = ws2.getCell(1, 1);
    t1.value     = "RATEIO VASCON";
    t1.font      = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
    t1.alignment = { horizontal: "center", vertical: "middle" };
    t1.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1F3C" } };
    ws2.getRow(1).height = 28;

    // Competência
    ws2.mergeCells(2, 1, 2, 6);
    const t2 = ws2.getCell(2, 1);
    t2.value     = `COMPETÊNCIA ${dataBR(resultado.filtros.dataInicio)} A ${dataBR(resultado.filtros.dataFim)}`;
    t2.font      = { bold: true, color: { argb: "FFFFFFFF" } };
    t2.alignment = { horizontal: "center" };
    t2.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC0281C" } };

    ws2.addRow([]);

    // Valor unitário
    ws2.getCell(4, 2).value  = "Valor Unitário";
    ws2.getCell(4, 3).value  = resultado.valores.valorVascon;
    ws2.getCell(4, 3).numFmt = '"R$" #,##0.00';
    ws2.getCell(4, 3).font   = { bold: true };

    ws2.addRow([]);

    // Header rateio
    const rateioHeader = ["CONTA", "C. DE CUSTO", "DESCRIÇÃO CENTRO CUSTO", "QTDE", "SOMA", "%"];
    const hr = ws2.addRow(rateioHeader);
    hr.eachCell(cell => {
      cell.font      = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1F3C" } };
      cell.alignment = { horizontal: "center", wrapText: true };
      cell.border    = borda();
    });
    ws2.getRow(hr.number).height = 24;

    let totalQtde = 0, totalSoma = 0;

    (resultado.rateio || []).forEach((r, idx) => {
      const row = ws2.addRow([
        r.conta,
        r.cc,
        r.descricao,
        r.qtde,
        r.soma,
        r.pct
      ]);

      totalQtde += r.qtde;
      totalSoma += r.soma;

      row.eachCell((cell, ci) => {
        cell.fill   = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? "FFF3F6FB" : "FFFFFFFF" } };
        cell.border = borda("FFD9E2F3");
        if (ci === 5) cell.numFmt = '"R$" #,##0.00';
        if (ci === 6) cell.numFmt = "0.00%";
        if (ci === 4) cell.alignment = { horizontal: "center" };
      });
    });

    // Linha total geral
    ws2.addRow([]);
    const totRateio = ws2.addRow(["", "", "Total geral", totalQtde, totalSoma, 1]);
    totRateio.eachCell((cell, ci) => {
      cell.font   = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill   = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1F3C" } };
      cell.border = borda();
      if (ci === 5) cell.numFmt = '"R$" #,##0.00';
      if (ci === 6) cell.numFmt = "0.00%";
    });

    // Largura das colunas da aba 2
    ws2.getColumn(1).width = 14;
    ws2.getColumn(2).width = 14;
    ws2.getColumn(3).width = 44;
    ws2.getColumn(4).width = 10;
    ws2.getColumn(5).width = 18;
    ws2.getColumn(6).width = 10;

    // ── Download ──────────────────────────────────────────────────────────
    const buffer = await wb.xlsx.writeBuffer();
    const blob   = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const nome   = `relatorio-refeitorio-${slug(titulo)}-${resultado.filtros.dataInicio}-a-${resultado.filtros.dataFim}.xlsx`;
    const link   = document.createElement("a");
    link.href     = URL.createObjectURL(blob);
    link.download = nome;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    if (typeof window.toast === "function") toast("Relatório exportado com sucesso.", "success");
  }

  // ── Interface ─────────────────────────────────────────────────────────────
  function criarTelaRelatorios() {
    const mod = $("mod-relatorios");
    if (!mod) return;

    mod.dataset.advanced = "final-2026-06-10";

    mod.innerHTML = `
      <div class="section-header">
        <div>
          <div class="section-title">📈 Relatórios Gerenciais</div>
          <div style="font-size:0.78rem;color:rgba(143,170,210,0.58);margin-top:0.25rem">
            Quantidade, centro de custo, colaborador, valores e rateio Vascon.
          </div>
        </div>
        <button class="btn-secondary" type="button" id="btnExportarRelatorioFinal">📥 Exportar Excel</button>
      </div>

      <div class="form-grid" style="margin-bottom:1rem;grid-template-columns:minmax(160px,1fr) minmax(160px,1fr) minmax(200px,1fr) minmax(220px,1fr) minmax(200px,1fr);align-items:end">
        <div class="form-group">
          <label class="form-label">Data inicial</label>
          <input class="form-input" type="date" id="relDataInicio">
        </div>
        <div class="form-group">
          <label class="form-label">Data final</label>
          <input class="form-input" type="date" id="relDataFim">
        </div>
        <div class="form-group">
          <label class="form-label">NF Vascon recebida (R$)</label>
          <input class="form-input" type="number" step="0.01" id="relValorNF" placeholder="Ex: 28487.68">
        </div>
        <div class="form-group">
          <label class="form-label">Tipo de relatório</label>
          <select class="form-select" id="tipoRelatorio">
            <option value="resumo-dia">Resumo por dia</option>
            <option value="centro-custo">Por centro de custo</option>
            <option value="colaborador">Por colaborador</option>
          </select>
        </div>
        <div class="form-group" style="display:flex;align-items:flex-end">
          <button class="btn-primary" type="button" id="btnGerarRelatorioFinal" style="width:100%">🔍 Gerar relatório</button>
        </div>
      </div>

      <div class="stats-grid" style="margin-bottom:1rem">
        <div class="stat-card"><div class="stat-icon">🍽️</div><div class="stat-value" id="relTotalGeral">—</div><div class="stat-label">Total refeições</div></div>
        <div class="stat-card"><div class="stat-icon">💰</div><div class="stat-value" id="relCustoVascon">—</div><div class="stat-label">Custo Vascon</div></div>
        <div class="stat-card"><div class="stat-icon">💳</div><div class="stat-value" id="relDescontoFolha">—</div><div class="stat-label">Desconto funcionários</div></div>
        <div class="stat-card"><div class="stat-icon">📄</div><div class="stat-value" id="relDiferencaNF">—</div><div class="stat-label">Diferença NF</div></div>
        <div class="stat-card"><div class="stat-icon">🥗</div><div class="stat-value" id="rel-principal">0</div><div class="stat-label">Principal</div></div>
        <div class="stat-card"><div class="stat-icon">🥦</div><div class="stat-value" id="rel-light">0</div><div class="stat-label">Light</div></div>
        <div class="stat-card"><div class="stat-icon">🥩</div><div class="stat-value" id="rel-carne">0</div><div class="stat-label">Carne</div></div>
        <div class="stat-card"><div class="stat-icon">🍝</div><div class="stat-value" id="rel-massa">0</div><div class="stat-label">Massa</div></div>
        <div class="stat-card"><div class="stat-icon">🍔</div><div class="stat-value" id="rel-lanche">0</div><div class="stat-label">Lanche</div></div>
      </div>

      <div style="margin-top:1rem">
        <div class="section-title" id="relTituloTabela" style="font-size:1rem;margin-bottom:0.8rem">Resumo por dia</div>
        <div class="alert alert-info" style="margin-bottom:0.8rem;font-size:0.78rem">
          💡 O Excel exportado contém 2 abas: <b>Relatório</b> e <b>Rateio Vascon</b> (no padrão para envio ao fiscal junto com a NF).
        </div>
        <div class="table-wrap">
          <table class="table">
            <thead id="relHead"></thead>
            <tbody id="relTableFinal"></tbody>
          </table>
        </div>
      </div>
    `;

    const hoje      = new Date();
    const primeiroDia = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
    $("relDataInicio").value = toInputDate(primeiroDia);
    $("relDataFim").value    = toInputDate(hoje);

    $("btnGerarRelatorioFinal").onclick   = () => gerarRelatorioFinal(semanaAtualFallback());
    $("btnExportarRelatorioFinal").onclick = exportarRelatorioExcelFinal;
    $("tipoRelatorio").onchange           = () => gerarRelatorioFinal(semanaAtualFallback());
  }

  function getFiltros() {
    return {
      dataInicio: $("relDataInicio")?.value || toInputDate(new Date()),
      dataFim:    $("relDataFim")?.value    || toInputDate(new Date()),
      nf:         valorNumero($("relValorNF")?.value || 0),
      tipo:       $("tipoRelatorio")?.value || "resumo-dia"
    };
  }

  // ── Exposição pública ─────────────────────────────────────────────────────
  window.loadRelatorios = async function (semanaId) {
    criarTelaRelatorios();
    await gerarRelatorioFinal(semanaId || semanaAtualFallback());
  };

  window.loadRelatoriosAvancados  = window.loadRelatorios;
  window.exportarRelatorioCSV     = exportarRelatorioExcelFinal;
  window.exportarCSV              = exportarRelatorioExcelFinal;
  window.exportarExcel            = exportarRelatorioExcelFinal;

  window.AdminRelatorios = {
    carregar:     window.loadRelatorios,
    gerar:        gerarRelatorioFinal,
    exportarExcel: exportarRelatorioExcelFinal
  };

  // Intercepta cliques em "exportar" quando o módulo relatórios está ativo
  document.addEventListener("click", event => {
    const btn   = event.target.closest("button, a");
    if (!btn) return;
    const texto = normalizar(btn.innerText || "");
    const ativo = $("mod-relatorios")?.classList.contains("active") ||
      normalizar(document.querySelector(".nav-item.active")?.innerText || "").includes("relatorios");
    if (ativo && texto.includes("exportar")) {
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      exportarRelatorioExcelFinal();
      return false;
    }
  }, true);

  document.addEventListener("click", event => {
    const item  = event.target.closest("[data-module='relatorios'], .nav-item");
    if (!item) return;
    const texto = normalizar(`${item.innerText || ""} ${item.dataset?.module || ""}`);
    if (texto.includes("relatorios")) setTimeout(() => window.loadRelatorios(semanaAtualFallback()), 120);
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
