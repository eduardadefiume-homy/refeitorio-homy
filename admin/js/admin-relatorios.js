// admin-relatorios.js — Relatórios gerenciais Homy + Rateio Vascon padrão
// Padrão: cabeçalho Homy, período solicitado, data de geração, Valor Vascon/desconto e fórmulas no Excel.

const AdminRelatorios = window.AdminRelatorios = {
  _pedidos: [],
  _periodo: { ini: "", fim: "" },
  _valorRef: { vascon: 0, desconto: 0, titulo: "", inicio: "", fim: "" },
  _colsValores: null,

  _CC_MAPA: {
    "110101":"DIRETORIA PRESIDENCIAL","110201":"DIRETORIA ADMINISTRATIVA",
    "110202":"DIRETORIA DE PRODUTOS","120101":"ADM GERAL","120102":"CUSTOS",
    "120103":"LEGALIZAÇÃO","120201":"CONTABILIDADE","120202":"FISCAL",
    "120301":"FINANCEIRO","120401":"RECURSOS HUMANOS","120402":"DEPARTAMENTO PESSOAL",
    "120501":"TI","120601":"RECEPÇÃO","120602":"PORTARIA",
    "120603":"ASSEIO E CONSERVAÇÃO","120604":"JARDINAGEM","150101":"SUPRIMENTOS",
    "160101":"CONTROLADORIA E COMPLIANCE","160102":"ADM CONTRATOS","170101":"SGI",
    "180101":"P&D","190101":"PATIO EXTERNO","220101":"ADM VENDAS",
    "220201":"COML INTERNO - SUPORTE","220202":"COML INTERNO - ATIVO",
    "220301":"COML EXTERNO - CLT","220302":"COML EXTERNO - REPRESENTANTE",
    "230101":"SUPORTE TECNICO INDUSTRIAL","230102":"SUPORTE TECNICO OBRAS/INFRA",
    "240101":"MARKETING","250101":"FATURAMENTO","250102":"LOGISTICA",
    "250103":"EXPEDIÇÃO","320101":"PRODUÇÃO","320201":"ENVASE MANUAL",
    "320202":"ENVASE AUTOMATICO","320301":"LABORATORIO E CONTROLE QUALIDADE",
    "360101":"APOIO A PRODUÇÃO","360102":"PCP","360201":"MANUTENÇÃO",
    "360301":"ALMOXARIFADO DE INSUMOS"
  },

  async load(semanaId) {
    try {
      await SP.init();
      const datas = SP.getWeekDates(semanaId);
      const ini = datas[0].toISOString().slice(0, 10);
      const fim = datas[4].toISOString().slice(0, 10);
      this._setValAny(ini, "relDataIni");
      this._setValAny(fim, "relDataFim");
      this._bindControles();
      await this._buscar(ini, fim);
    } catch (e) {
      console.error("[Relatórios] load", e);
      AdminUtils.toast("Erro ao carregar Relatórios: " + (e.message || e), "error");
    }
  },

  _getEl(...ids) {
    for (const id of ids) {
      const el = document.getElementById(id);
      if (el) return el;
    }
    return null;
  },

  _setValAny(valor, ...ids) {
    const el = this._getEl(...ids);
    if (el) el.value = valor ?? "";
  },

  _valAny(...ids) {
    const el = this._getEl(...ids);
    return (el?.value || "").trim();
  },

  _setTxtAny(valor, ...ids) {
    const el = this._getEl(...ids);
    if (el) el.textContent = valor ?? "—";
  },

  _norm(v) {
    return String(v || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
  },

  _normCampo(v) {
    return this._norm(String(v || "")
      .replace(/_x00e1_/gi, "a")
      .replace(/_x00e3_/gi, "a")
      .replace(/_x00e7_/gi, "c"));
  },

  _pick(obj, ...keys) {
    for (const k of keys.filter(Boolean)) {
      const v = SP.pick ? SP.pick(obj, k) : obj?.[k];
      if (v !== undefined && v !== null && String(v) !== "") return v;
    }
    return "";
  },

  _campoCombina(tipo, key) {
    const k = this._normCampo(key);
    if (!k || k.startsWith("@odata")) return false;
    if (tipo === "titulo") return k === "title" || k === "titulo";
    if (tipo === "inicio") return k.includes("data") && k.includes("inicio");
    if (tipo === "fim") return k.includes("data") && k.includes("fim");
    if (tipo === "vascon") return k.includes("vascon");
    if (tipo === "desconto") return k.includes("desconto") && (k.includes("funcion") || k.includes("func"));
    if (tipo === "obs") return k.includes("observ") || k === "obs";
    if (tipo === "ativo") return k === "ativo" || k.endsWith("ativo");
    return false;
  },

  _acharCampoNoItem(item, tipo) {
    if (!item || typeof item !== "object") return null;
    return Object.keys(item).find(k => this._campoCombina(tipo, k)) || null;
  },

  _calibrarColunasPorItens(lista) {
    if (!this._colsValores) this._colsValores = {};
    const tipos = ["titulo", "inicio", "fim", "vascon", "desconto", "obs", "ativo"];
    for (const item of lista || []) {
      for (const tipo of tipos) {
        const k = this._acharCampoNoItem(item, tipo);
        if (k) this._colsValores[tipo] = k;
      }
    }
    return this._colsValores;
  },

  _valorCampo(item, tipo, ...fallbackKeys) {
    const c = this._colsValores || {};
    const direto = this._pick(item, c[tipo], ...fallbackKeys);
    if (direto !== "") return direto;
    const dinamico = this._acharCampoNoItem(item, tipo);
    if (dinamico) return item[dinamico];
    return "";
  },

  _dateISO(v) {
    if (!v) return "";
    const s = String(v);
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const d = new Date(s);
    return isNaN(d) ? "" : d.toISOString().slice(0, 10);
  },

  _brDate(iso) {
    const d = this._dateISO(iso);
    if (!d) return "—";
    const [a, m, dia] = d.split("-");
    return `${dia}/${m}/${a}`;
  },

  _money(v) {
    const n = Number(v || 0);
    return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  },

  _num(v) {
    if (v === null || v === undefined || String(v).trim() === "") return 0;
    let s = String(v).replace(/R\$/gi, "").replace(/\s/g, "").trim();
    if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
    s = s.replace(/[^0-9.-]/g, "");
    const n = Number(s);
    return Number.isFinite(n) ? n : 0;
  },

  _moeda(v) {
    if (v === null || v === undefined || String(v).trim() === "") return null;
    return this._num(v);
  },

  _periodosSobrepoem(iniA, fimA, iniB, fimB) {
    if (!iniA || !fimA || !iniB || !fimB) return false;
    return iniA <= fimB && fimA >= iniB;
  },

  async _resolverColunasValores() {
    if (this._colsValores) return this._colsValores;
    this._colsValores = {
      titulo: "Title",
      inicio: "Data_Inicio",
      fim: "Data_Fim",
      vascon: "Valor_Vascon",
      desconto: "Valor_Desconto_Funcionário",
      obs: "Observacao",
      ativo: "Ativo"
    };

    try {
      const siteId = await SP.getSiteId();
      const listId = await SP.getListId("Valores de Refeição");
      const data = await SP.graph("GET", `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`);
      const cols = data?.value || [];
      const byIncludes = (...terms) => {
        const nt = terms.map(t => this._normCampo(t));
        const found = cols.find(c => {
          const n = this._normCampo(`${c.name} ${c.displayName}`);
          return nt.every(t => n.includes(t));
        });
        return found?.name || null;
      };
      const byExact = (...cands) => {
        const cn = cands.map(c => this._normCampo(c));
        const found = cols.find(c => cn.includes(this._normCampo(c.name)) || cn.includes(this._normCampo(c.displayName)));
        return found?.name || null;
      };
      this._colsValores = {
        titulo: byExact("Title", "Título", "Titulo") || this._colsValores.titulo,
        inicio: byExact("Data_Inicio", "Data Inicio", "Data início") || byIncludes("data", "inicio") || this._colsValores.inicio,
        fim: byExact("Data_Fim", "Data Fim") || byIncludes("data", "fim") || this._colsValores.fim,
        vascon: byExact("Valor_Vascon", "Valor Vascon") || byIncludes("vascon") || this._colsValores.vascon,
        desconto: byExact("Valor_Desconto_Funcionário", "Valor_Desconto_Funcionario", "Valor_Desconto_Funcion_x00e1_rio", "Desconto Funcionário") || byIncludes("desconto", "funcion") || this._colsValores.desconto,
        obs: byExact("Observacao", "Observação") || byIncludes("observ") || this._colsValores.obs,
        ativo: byExact("Ativo") || this._colsValores.ativo
      };
    } catch (e) {
      console.warn("[Relatórios] Não foi possível resolver colunas de valores dinamicamente.", e);
    }
    return this._colsValores;
  },

  async _carregarValorReferencia(ini, fim) {
    try {
      await this._resolverColunasValores();
      let valores = await SP.getValoresRefeicao(false).catch(() => []);
      if ((!valores || !valores.length) && window.AdminValores?._lista?.length) valores = window.AdminValores._lista;
      this._calibrarColunasPorItens(valores);

      const norm = valores.map(v => ({
        raw: v,
        id: v.id || "",
        titulo: this._valorCampo(v, "titulo", "Title", "Titulo", "Título") || "Valor refeição",
        inicio: this._dateISO(this._valorCampo(v, "inicio", "Data_Inicio", "DataInicio")),
        fim: this._dateISO(this._valorCampo(v, "fim", "Data_Fim", "DataFim")),
        vascon: this._num(this._valorCampo(v, "vascon", "Valor_Vascon", "ValorVascon")),
        desconto: this._num(this._valorCampo(v, "desconto", "Valor_Desconto_Funcionário", "Valor_Desconto_Funcionario", "Valor_Desconto_Funcion_x00e1_rio", "Desconto")),
        ativo: SP.isTrue ? SP.isTrue(this._valorCampo(v, "ativo", "Ativo")) : !!this._valorCampo(v, "ativo", "Ativo")
      }));

      const escolhido =
        norm.find(v => v.ativo && this._periodosSobrepoem(ini, fim, v.inicio, v.fim)) ||
        norm.find(v => this._periodosSobrepoem(ini, fim, v.inicio, v.fim)) ||
        norm.find(v => v.ativo) ||
        norm[0] ||
        { vascon: 0, desconto: 0, titulo: "Sem valor cadastrado", inicio: "", fim: "" };

      this._valorRef = escolhido;
      this._setTxtAny(this._money(escolhido.vascon), "relValorVascon", "relVasconUnit", "relValorUnitarioVascon");
      this._setTxtAny(this._money(escolhido.desconto), "relValorDesconto", "relDescontoUnit", "relValorUnitarioDesconto");
    } catch (e) {
      console.warn("[Relatórios] Não foi possível buscar valor Vascon:", e);
      this._valorRef = { vascon: 0, desconto: 0, titulo: "Sem valor cadastrado", inicio: "", fim: "" };
    }
  },

  _dataPedido(p) {
    // Relatório deve considerar a data REAL da refeição.
    // Data_Hora é apenas a data em que o item foi criado/alterado no SharePoint.
    // Isso evita que pedidos gerados/alterados hoje entrem como refeição de hoje sem Dia válido.
    const semana = this._pick(p, "Semana_id", "Semana");
    const dia = this._pick(p, "Dia");
    if (semana && dia && SP.getDataRefBySemanaDia) return SP.getDataRefBySemanaDia(semana, dia);
    if (semana && !dia) return "";
    return this._dateISO(this._pick(p, "Data_Referencia", "Data", "Data_Hora"));
  },

  _pedidoValidoRelatorio(p) {
    const nome = String(this._pick(p, "Colaborador_nome", "Colaborador", "Nome") || "").trim();
    const colabId = String(this._pick(p, "Colaborador_id", "colaborador_id") || "").trim();
    const dia = String(this._pick(p, "Dia", "dia") || "").trim();
    const opcao = String(this._pick(p, "Opcao", "Opção", "opcao") || "").trim();
    const nomeNorm = this._norm(nome);

    // Linhas fantasmas no SharePoint costumam ter apenas Semana_id/Opcao/Confirmado/Data_Hora,
    // sem colaborador e sem dia. Elas não representam refeição real.
    if (!dia || !opcao) return false;
    if (!nome && !colabId) return false;
    if (nomeNorm === "pedido" && !colabId) return false;

    return true;
  },

  async _buscar(ini, fim) {
    this._periodo = { ini, fim };
    const wrap = document.getElementById("relConteudo");
    if (wrap) wrap.innerHTML = `<div class="alert alert-info">Carregando...</div>`;

    try {
      await SP.init();
      await this._carregarValorReferencia(ini, fim);
      const todos = await SP.getItems("Pedidos");
      this._pedidos = todos.filter(p => {
        if (!this._pedidoValidoRelatorio(p)) return false;
        const d = this._dataPedido(p);
        return d && d >= ini && d <= fim;
      });
      this._renderCards();
      this._renderTipo();
    } catch (e) {
      console.error("[Relatórios] buscar", e);
      if (wrap) wrap.innerHTML = `<div class="alert" style="background:rgba(220,50,50,.1);color:#ff8080">Erro: ${AdminUtils.esc(e.message || e)}</div>`;
    }
  },

  _isConf(p) {
    const s = this._norm(this._pick(p, "Status") || "");
    return ["confirmado", "extra", "aprovado", "travado"].includes(s) ||
           (SP.isTrue && SP.isTrue(this._pick(p, "Confirmado")));
  },

  _opcao(p) {
    return this._norm(this._pick(p, "Opcao", "Opção") || "");
  },

  _countOp(lista, op) {
    return lista.filter(p => this._opcao(p) === op).length;
  },

  _conf() {
    return this._pedidos.filter(p => this._pedidoValidoRelatorio(p) && this._isConf(p));
  },

  _nfInformada() {
    return this._moeda(this._valAny(
      "relNfVascon", "relNFVascon", "relNfVasconRecebida", "relTotalNFVascon",
      "relTotalNfVascon", "relTotalNF", "relTotalNf", "nfTotalRelatorio", "nfTotalDigitado"
    ));
  },

  _formatarCC(valor) {
    if (!valor) return "Sem setor";
    const v = String(valor).trim();
    if (!v || v === "—") return "Sem setor";
    const codigo = (v.match(/\b\d{6}\b/) || [""])[0];
    if (codigo) {
      const nomeMapa = this._CC_MAPA[codigo];
      if (nomeMapa) return `${codigo} - ${nomeMapa}`;
      const partes = v.split(" - ").map(x => x.trim()).filter(Boolean);
      const nome = partes.find(x => !/^\d{6}$/.test(x));
      return nome ? `${codigo} - ${nome}` : codigo;
    }
    return v;
  },

  _codigoCC(valor) {
    const v = String(valor || "");
    const codigo = (v.match(/\b\d{6}\b/) || [""])[0];
    return codigo || v || "Sem CC";
  },

  _nomeCC(valor) {
    const codigo = this._codigoCC(valor);
    if (/^\d{6}$/.test(codigo)) return this._CC_MAPA[codigo] || "";
    return String(valor || "Sem CC");
  },

  _centroCustoPedido(p) {
    const raw = this._pick(p, "Centro_Custo", "Setor", "Departamento");
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo") || "");
    const nome = this._norm(this._pick(p, "Colaborador_nome", "Title", "Nome") || "");
    const isExtra = origem.includes("extra") || origem.includes("guarda") || origem.includes("visitante") || origem.includes("motorista") || nome.includes("guarda") || nome.includes("refeicao extra");
    if ((!raw || String(raw).trim() === "—" || this._norm(raw).includes("sem")) && isExtra) return "120602 - PORTARIA";
    if (origem.includes("guarda") || nome.includes("guarda")) return "120602 - PORTARIA";
    return raw || "Sem CC";
  },
  _agruparPorCC(conf) {
    const mapa = {};
    conf.forEach(p => {
      const cc = this._formatarCC(this._centroCustoPedido(p));
      if (!mapa[cc]) mapa[cc] = [];
      mapa[cc].push(p);
    });
    return Object.entries(mapa)
      .sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0]))
      .map(([cc, lista]) => ({ cc, lista }));
  },

  _renderCards() {
    const conf = this._conf();
    const total = conf.length;
    const custoVascon = total * (this._valorRef.vascon || 0);
    const descontoFunc = total * (this._valorRef.desconto || 0);
    const nf = this._nfInformada();
    const dif = nf === null ? null : nf - custoVascon;

    this._setTxtAny(this._countOp(conf, "principal"), "rel-principal", "relPrincipal");
    this._setTxtAny(this._countOp(conf, "light"), "rel-light", "relLight");
    this._setTxtAny(this._countOp(conf, "carne"), "rel-carne", "relCarne");
    this._setTxtAny(this._countOp(conf, "massa"), "rel-massa", "relMassa");
    this._setTxtAny(this._countOp(conf, "lanche"), "rel-lanche", "relLanche");
    this._setTxtAny(total, "rel-total", "relTotal", "relTotalRefeicoes");
    this._setTxtAny(this._money(custoVascon), "rel-custo-vascon", "relCustoVascon", "relTotalVascon");
    this._setTxtAny(this._money(descontoFunc), "rel-desconto", "relDescontoFuncionarios", "relDescontoFuncionario");
    this._setTxtAny(dif === null ? "—" : this._money(dif), "rel-diferenca-nf", "relDiferencaNF", "relDifNF");
  },

  _renderTipo() {
    const tipo = this._valAny("relTipo") || "dia";
    const wrap = document.getElementById("relConteudo");
    if (!wrap) return;
    const conf = this._conf();
    if (tipo === "dia") this._renderPorDia(conf, wrap);
    else if (tipo === "cc") this._renderPorCC(conf, wrap);
    else this._renderPorCCFunc(conf, wrap);
  },

  _linhaResumoValor() {
    const v = this._valorRef || {};
    return `<div class="alert alert-info" style="margin-bottom:.8rem;line-height:1.6">
      Valor Vascon aplicado: <b>${this._money(v.vascon || 0)}</b> ·
      Desconto funcionário: <b>${this._money(v.desconto || 0)}</b>
      ${v.inicio && v.fim ? ` · Vigência ${this._brDate(v.inicio)} a ${this._brDate(v.fim)}` : ""}.
    </div>`;
  },

  _renderPorDia(conf, wrap) {
    const mapa = {};
    conf.forEach(p => {
      const d = this._dataPedido(p) || "—";
      if (!mapa[d]) mapa[d] = [];
      mapa[d].push(p);
    });
    const sorted = Object.entries(mapa).sort(([a], [b]) => a.localeCompare(b));
    const unit = this._valorRef.vascon || 0;
    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">📅 Resumo por dia — período ${this._brDate(this._periodo.ini)} a ${this._brDate(this._periodo.fim)}</div>
      ${this._linhaResumoValor()}
      <div class="table-wrap"><table class="table">
        <thead><tr><th>Data</th><th>Principal</th><th>Light</th><th>Carne</th><th>Massa</th><th>Lanche</th><th>Total</th><th>Total Vascon</th></tr></thead>
        <tbody>
          ${sorted.length ? sorted.map(([data, lista]) => `<tr>
            <td>${this._brDate(data)}</td><td>${this._countOp(lista, "principal")}</td><td>${this._countOp(lista, "light")}</td>
            <td>${this._countOp(lista, "carne")}</td><td>${this._countOp(lista, "massa")}</td><td>${this._countOp(lista, "lanche")}</td>
            <td><strong>${lista.length}</strong></td><td>${this._money(lista.length * unit)}</td>
          </tr>`).join("") : `<tr><td colspan="8" class="empty-cell">Nenhum pedido no período.</td></tr>`}
        </tbody></table></div>`;
  },

  _renderPorCC(conf, wrap) {
    const sorted = this._agruparPorCC(conf);
    const unit = this._valorRef.vascon || 0;
    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">🏢 Por centro de custo — período ${this._brDate(this._periodo.ini)} a ${this._brDate(this._periodo.fim)}</div>
      ${this._linhaResumoValor()}
      <div class="table-wrap"><table class="table">
        <thead><tr><th>Centro de Custo</th><th>Principal</th><th>Light</th><th>Carne</th><th>Massa</th><th>Lanche</th><th>Total</th><th>Total Vascon</th></tr></thead>
        <tbody>
          ${sorted.length ? sorted.map(({ cc, lista }) => `<tr>
            <td>${AdminUtils.esc(cc)}</td><td>${this._countOp(lista, "principal")}</td><td>${this._countOp(lista, "light")}</td>
            <td>${this._countOp(lista, "carne")}</td><td>${this._countOp(lista, "massa")}</td><td>${this._countOp(lista, "lanche")}</td>
            <td><strong>${lista.length}</strong></td><td>${this._money(lista.length * unit)}</td>
          </tr>`).join("") : `<tr><td colspan="8" class="empty-cell">Nenhum pedido no período.</td></tr>`}
        </tbody></table></div>`;
  },

  _renderPorCCFunc(conf, wrap) {
    const mapa = {};
    conf.forEach(p => {
      const cc = this._formatarCC(this._centroCustoPedido(p));
      const nome = this._pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "Desconhecido";
      const key = `${cc}||${nome}`;
      if (!mapa[key]) mapa[key] = { cc, nome, lista: [] };
      mapa[key].lista.push(p);
    });
    const sorted = Object.values(mapa).sort((a, b) => a.cc.localeCompare(b.cc) || a.nome.localeCompare(b.nome));
    const unitDesc = this._valorRef.desconto || 0;
    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">👤 Por CC e funcionário — período ${this._brDate(this._periodo.ini)} a ${this._brDate(this._periodo.fim)}</div>
      ${this._linhaResumoValor()}
      <div class="table-wrap"><table class="table">
        <thead><tr><th>Centro de Custo</th><th>Colaborador</th><th>Total refeições</th><th>Desconto total</th><th>Datas</th></tr></thead>
        <tbody>${sorted.length ? sorted.map(({ cc, nome, lista }) => `<tr>
          <td>${AdminUtils.esc(cc)}</td><td>${AdminUtils.esc(nome)}</td><td>${lista.length}</td><td>${this._money(lista.length * unitDesc)}</td>
          <td style="font-size:.78rem;color:rgba(143,170,210,.6)">${[...new Set(lista.map(p => this._dataPedido(p)).filter(Boolean))].sort().map(d => this._brDate(d)).join(", ")}</td>
        </tr>`).join("") : `<tr><td colspan="5" class="empty-cell">Nenhum pedido no período.</td></tr>`}</tbody>
      </table></div>`;
  },

  _linhasRelatorio(tipo, conf) {
    const unitV = this._valorRef.vascon || 0;
    const unitD = this._valorRef.desconto || 0;

    if (tipo === "dia") {
      const mapa = {};
      conf.forEach(p => {
        const d = this._dataPedido(p) || "—";
        if (!mapa[d]) mapa[d] = [];
        mapa[d].push(p);
      });
      return Object.entries(mapa).sort(([a], [b]) => a.localeCompare(b)).map(([data, lista]) => ({
        Dia: new Date(data + "T00:00:00").toLocaleDateString("pt-BR", { weekday: "long" }),
        Data: data,
        Principal: this._countOp(lista, "principal"),
        Light: this._countOp(lista, "light"),
        Carne: this._countOp(lista, "carne"),
        Massa: this._countOp(lista, "massa"),
        Lanche: this._countOp(lista, "lanche"),
        Total_Refeicoes: lista.length,
        Valor_Unitario_Vascon: unitV,
        Total_Vascon: lista.length * unitV,
        Valor_Desconto_Funcionario: unitD,
        Desconto_Total: lista.length * unitD
      }));
    }

    if (tipo === "cc") {
      return this._agruparPorCC(conf).map(({ cc, lista }) => ({
        Centro_Custo: cc,
        Principal: this._countOp(lista, "principal"),
        Light: this._countOp(lista, "light"),
        Carne: this._countOp(lista, "carne"),
        Massa: this._countOp(lista, "massa"),
        Lanche: this._countOp(lista, "lanche"),
        Total_Refeicoes: lista.length,
        Valor_Unitario_Vascon: unitV,
        Total_Vascon: lista.length * unitV
      }));
    }

    const mapa = {};
    conf.forEach(p => {
      const cc = this._formatarCC(this._centroCustoPedido(p));
      const nome = this._pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "Desconhecido";
      const key = `${cc}||${nome}`;
      if (!mapa[key]) mapa[key] = { cc, nome, lista: [] };
      mapa[key].lista.push(p);
    });
    return Object.values(mapa).sort((a, b) => a.cc.localeCompare(b.cc) || a.nome.localeCompare(b.nome)).map(({ cc, nome, lista }) => ({
      Centro_Custo: cc,
      Colaborador: nome,
      Total_Refeicoes: lista.length,
      Valor_Desconto_Funcionario: unitD,
      Desconto_Total: lista.length * unitD,
      Valor_Unitario_Vascon: unitV,
      Total_Vascon: lista.length * unitV
    }));
  },

  async _ensureExcelJS() {
    if (window.ExcelJS) return true;
    await new Promise((resolve, reject) => {
      const ja = document.querySelector('script[data-exceljs="1"]');
      if (ja) {
        ja.addEventListener("load", resolve, { once: true });
        ja.addEventListener("error", reject, { once: true });
        return;
      }
      const s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js";
      s.dataset.exceljs = "1";
      s.onload = resolve;
      s.onerror = () => reject(new Error("Não foi possível carregar ExcelJS."));
      document.head.appendChild(s);
    });
    return !!window.ExcelJS;
  },

  _downloadBuffer(buffer, filename) {
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  },

  _setCellStyle(cell, fill, fontColor = "FFFFFFFF", bold = true) {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } };
    cell.font = { bold, color: { argb: fontColor } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
  },

  _border(cell) {
    cell.border = {
      top: { style: "thin", color: { argb: "FFB8C7DD" } },
      left: { style: "thin", color: { argb: "FFB8C7DD" } },
      bottom: { style: "thin", color: { argb: "FFB8C7DD" } },
      right: { style: "thin", color: { argb: "FFB8C7DD" } }
    };
  },

  _aplicarBordas(ws, fromRow = 1, toRow = ws.rowCount, fromCol = 1, toCol = ws.columnCount) {
    for (let r = fromRow; r <= toRow; r++) {
      for (let c = fromCol; c <= toCol; c++) this._border(ws.getCell(r, c));
    }
  },

  _tituloTipo(tipo) {
    if (tipo === "cc") return "POR CENTRO DE CUSTO";
    if (tipo === "ccfunc") return "POR CC E FUNCIONÁRIO";
    return "RESUMO POR DIA";
  },

  _nomeAbaTipo(tipo) {
    if (tipo === "cc") return "Por centro de custo";
    if (tipo === "ccfunc") return "Por CC funcionário";
    return "Resumo por dia";
  },

  _addRelatorioSheet(wb, tipo, linhas, conf) {
    const ws = wb.addWorksheet(this._nomeAbaTipo(tipo), { properties: { tabColor: { argb: "FFC0281C" } } });
    const headers = Object.keys(linhas[0] || {});
    const lastCol = Math.max(headers.length, 8);
    const title = `RELATÓRIO REFEITÓRIO HOMY — ${this._tituloTipo(tipo)}`;

    ws.mergeCells(1, 1, 1, lastCol);
    ws.getCell(1, 1).value = title;
    this._setCellStyle(ws.getCell(1, 1), "FF09213F");
    ws.getRow(1).height = 24;

    ws.mergeCells(2, 1, 2, lastCol);
    ws.getCell(2, 1).value = `Período: ${this._brDate(this._periodo.ini)} a ${this._brDate(this._periodo.fim)}`;
    this._setCellStyle(ws.getCell(2, 1), "FFC0281C");

    ws.mergeCells(3, 1, 3, lastCol);
    ws.getCell(3, 1).value = `Gerado em: ${new Date().toLocaleDateString("pt-BR")}`;
    ws.getCell(3, 1).alignment = { horizontal: "center" };
    ws.getCell(3, 1).font = { bold: true, color: { argb: "FF09213F" } };
    ws.getCell(3, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF1FF" } };

    ws.getCell("A5").value = "Total de refeições";
    ws.getCell("B5").value = conf.length;
    ws.getCell("A6").value = "Valor unitário Vascon";
    ws.getCell("B6").value = this._valorRef.vascon || 0;
    ws.getCell("A7").value = "Valor unitário desconto funcionário";
    ws.getCell("B7").value = this._valorRef.desconto || 0;
    ws.getCell("A8").value = "Total Vascon";
    ws.getCell("B8").value = { formula: "B5*B6", result: conf.length * (this._valorRef.vascon || 0) };
    ws.getCell("D5").value = "Valor cadastrado";
    ws.getCell("E5").value = this._valorRef.titulo || "";
    ws.getCell("D6").value = "Vigência";
    ws.getCell("E6").value = `${this._brDate(this._valorRef.inicio)} a ${this._brDate(this._valorRef.fim)}`;

    ["A5", "A6", "A7", "A8", "D5", "D6"].forEach(addr => ws.getCell(addr).font = { bold: true });
    ["B6", "B7", "B8"].forEach(addr => ws.getCell(addr).numFmt = '"R$" #,##0.00');

    const headerRow = 10;
    ws.getRow(headerRow).values = headers;
    ws.getRow(headerRow).height = 22;
    headers.forEach((h, idx) => {
      const cell = ws.getCell(headerRow, idx + 1);
      this._setCellStyle(cell, "FF09213F");
    });

    linhas.forEach(obj => {
      const row = ws.addRow(headers.map(h => obj[h]));
      const r = row.number;
      const totalCol = headers.indexOf("Total_Refeicoes") + 1;
      const unitVCol = headers.indexOf("Valor_Unitario_Vascon") + 1;
      const totalVCol = headers.indexOf("Total_Vascon") + 1;
      const unitDCol = headers.indexOf("Valor_Desconto_Funcionario") + 1;
      const totalDCol = headers.indexOf("Desconto_Total") + 1;
      if (totalCol && unitVCol && totalVCol) {
        row.getCell(totalVCol).value = { formula: `${ws.getColumn(totalCol).letter}${r}*${ws.getColumn(unitVCol).letter}${r}`, result: obj.Total_Vascon || 0 };
      }
      if (totalCol && unitDCol && totalDCol) {
        row.getCell(totalDCol).value = { formula: `${ws.getColumn(totalCol).letter}${r}*${ws.getColumn(unitDCol).letter}${r}`, result: obj.Desconto_Total || 0 };
      }
    });

    const firstData = headerRow + 1;
    const lastData = headerRow + linhas.length;
    const totalRow = lastData + 1;
    ws.getCell(totalRow, 1).value = "TOTAL";
    headers.forEach((h, idx) => {
      const col = idx + 1;
      const letter = ws.getColumn(col).letter;
      if (["Principal", "Light", "Carne", "Massa", "Lanche", "Total_Refeicoes", "Total_Vascon", "Desconto_Total"].includes(h)) {
        ws.getCell(totalRow, col).value = { formula: `SUM(${letter}${firstData}:${letter}${lastData})`, result: linhas.reduce((acc, obj) => acc + Number(obj[h] || 0), 0) };
      }
    });

    ws.autoFilter = { from: { row: headerRow, column: 1 }, to: { row: headerRow, column: headers.length } };
    ws.views = [{ state: "frozen", ySplit: headerRow }];
    ws.getRow(totalRow).font = { bold: true };
    ws.getRow(totalRow).eachCell(cell => cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF1FF" } });

    for (let r = firstData; r <= totalRow; r++) {
      if ((r - firstData) % 2 === 0 && r !== totalRow) {
        ws.getRow(r).eachCell(cell => cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4F8FF" } });
      }
    }

    headers.forEach((h, idx) => {
      const col = idx + 1;
      if (h.includes("Valor") || h.includes("Total_Vascon") || h.includes("Desconto")) ws.getColumn(col).numFmt = '"R$" #,##0.00';
      if (["Principal", "Light", "Carne", "Massa", "Lanche", "Total_Refeicoes"].includes(h)) ws.getColumn(col).numFmt = '0';
    });

    this._aplicarBordas(ws, 1, ws.rowCount, 1, headers.length);
    ws.columns.forEach((col, idx) => {
      let max = idx === 0 ? 16 : 12;
      col.eachCell({ includeEmpty: true }, cell => {
        const v = cell.value?.formula ? cell.value.result : cell.value;
        max = Math.max(max, String(v ?? "").length + 2);
      });
      col.width = Math.min(Math.max(max, 12), idx === 0 ? 32 : 46);
    });
    return ws;
  },

  _addRateioSheet(wb, conf) {
    const ws = wb.addWorksheet("Rateio Vascon", { properties: { tabColor: { argb: "FF09213F" } } });
    const gruposReais = new Map();
    this._agruparPorCC(conf).forEach(({ cc, lista }) => gruposReais.set(this._codigoCC(cc), lista.length));
    const codigos = Object.keys(this._CC_MAPA).sort();
    const unit = this._valorRef.vascon || 0;

    ws.mergeCells("A1:F1");
    ws.getCell("A1").value = "RATEIO VASCON";
    this._setCellStyle(ws.getCell("A1"), "FF09213F");
    ws.getRow(1).height = 24;

    ws.mergeCells("A2:F2");
    ws.getCell("A2").value = `COMPETÊNCIA ${this._brDate(this._periodo.ini)} A ${this._brDate(this._periodo.fim)}`;
    this._setCellStyle(ws.getCell("A2"), "FFC0281C");

    ws.getCell("B4").value = "Valor Unitário";
    ws.getCell("C4").value = unit;
    ws.getCell("C4").numFmt = '"R$" #,##0.00';
    ws.getCell("B4").font = { bold: true };

    const headerRow = 6;
    const headers = ["CONTA", "C. DE CUSTO", "DESCRIÇÃO CENTRO CUSTO", "QTDE", "SOMA", "%"];
    ws.getRow(headerRow).values = headers;
    headers.forEach((_, idx) => this._setCellStyle(ws.getCell(headerRow, idx + 1), "FF09213F"));

    codigos.forEach(codigo => {
      const qtd = gruposReais.get(codigo) || 0;
      ws.addRow(["51101015", codigo, this._CC_MAPA[codigo], qtd, null, null]);
    });

    const firstData = headerRow + 1;
    const lastData = headerRow + codigos.length;
    const totalRow = lastData + 1;
    ws.addRow(["TOTAL", "", "", null, null, null]);

    for (let r = firstData; r <= lastData; r++) {
      ws.getCell(`E${r}`).value = { formula: `D${r}*$C$4`, result: Number(ws.getCell(`D${r}`).value || 0) * unit };
      ws.getCell(`F${r}`).value = { formula: `IF($D$${totalRow}=0,0,D${r}/$D$${totalRow})`, result: conf.length ? Number(ws.getCell(`D${r}`).value || 0) / conf.length : 0 };
    }

    ws.getCell(`D${totalRow}`).value = { formula: `SUM(D${firstData}:D${lastData})`, result: conf.length };
    ws.getCell(`E${totalRow}`).value = { formula: `SUM(E${firstData}:E${lastData})`, result: conf.length * unit };
    ws.getCell(`F${totalRow}`).value = { formula: `SUM(F${firstData}:F${lastData})`, result: conf.length ? 1 : 0 };

    ws.getColumn(5).numFmt = '"R$" #,##0.00';
    ws.getColumn(6).numFmt = '0.00%';
    ws.getColumn(4).numFmt = '0';
    ws.getRow(totalRow).font = { bold: true };
    ws.getRow(totalRow).eachCell(cell => cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF1FF" } });

    for (let r = firstData; r < totalRow; r++) {
      if ((r - firstData) % 2 === 0) ws.getRow(r).eachCell(cell => cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4F8FF" } });
    }

    this._aplicarBordas(ws, 1, totalRow, 1, 6);
    ws.views = [{ state: "frozen", ySplit: headerRow }];
    ws.autoFilter = { from: { row: headerRow, column: 1 }, to: { row: headerRow, column: 6 } };
    ws.columns = [
      { width: 13 }, { width: 14 }, { width: 36 }, { width: 10 }, { width: 16 }, { width: 12 }
    ];
    return ws;
  },

  async _exportar() {
    try {
      const ini = this._valAny("relDataIni") || this._periodo.ini;
      const fim = this._valAny("relDataFim") || this._periodo.fim;
      if (!ini || !fim) { AdminUtils.toast("Informe o período.", "error"); return; }
      if (ini > fim) { AdminUtils.toast("Data início maior que fim.", "error"); return; }

      if (ini !== this._periodo.ini || fim !== this._periodo.fim) await this._buscar(ini, fim);
      else await this._carregarValorReferencia(this._periodo.ini, this._periodo.fim);
      if ((!this._valorRef || !this._valorRef.vascon) && window.AdminValores?._valorParaPeriodo) {
        try {
          if (!window.AdminValores._lista?.length) await window.AdminValores._carregar?.();
          const v = window.AdminValores._valorParaPeriodo(this._periodo.ini, this._periodo.fim);
          if (v?.vascon) this._valorRef = { ...this._valorRef, ...v };
        } catch (_) {}
      }

      const conf = this._conf();
      const tipo = this._valAny("relTipo") || "dia";
      const linhas = this._linhasRelatorio(tipo, conf);
      if (!linhas.length) { AdminUtils.toast("Nenhum dado para exportar.", "info"); return; }

      await this._ensureExcelJS();
      if (!window.ExcelJS) throw new Error("ExcelJS não carregou.");

      const wb = new ExcelJS.Workbook();
      wb.creator = "Refeitório Homy";
      wb.created = new Date();
      wb.modified = new Date();
      wb.calcProperties.fullCalcOnLoad = true;

      this._addRelatorioSheet(wb, tipo, linhas, conf);
      this._addRateioSheet(wb, conf);

      const buffer = await wb.xlsx.writeBuffer();
      this._downloadBuffer(buffer, `relatorio-refeitorio-${tipo}-${this._periodo.ini}-a-${this._periodo.fim}.xlsx`);
      AdminUtils.toast("Excel exportado no padrão Homy com fórmulas.", "success");
    } catch (e) {
      console.error("[Relatórios] exportar", e);
      AdminUtils.toast("Erro ao exportar Relatórios: " + (e.message || e), "error");
    }
  },

  _bindControles() {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundRel) {
        el.dataset.boundRel = "1";
        el.addEventListener(ev, fn);
      }
    };

    bind("btnBuscarRelatorio", "click", async () => {
      const ini = this._valAny("relDataIni");
      const fim = this._valAny("relDataFim");
      if (!ini || !fim) { AdminUtils.toast("Informe o período.", "error"); return; }
      if (ini > fim) { AdminUtils.toast("Data início maior que fim.", "error"); return; }
      await this._buscar(ini, fim);
    });

    bind("relTipo", "change", () => this._renderTipo());
    bind("btnExportarRelatorio", "click", () => this._exportar());

    [
      "relNfVascon", "relNFVascon", "relNfVasconRecebida", "relTotalNFVascon",
      "relTotalNfVascon", "relTotalNF", "relTotalNf", "nfTotalRelatorio", "nfTotalDigitado"
    ].forEach(id => bind(id, "input", () => this._renderCards()));
  }
};
