// admin-relatorios.js — Relatórios gerenciais + Rateio Vascon automático
// Versão limpa: valor Vascon detectado da lista, exportação Excel formatada com fórmulas

const AdminRelatorios = window.AdminRelatorios = {
  _pedidos: [],
  _periodo: { ini: "", fim: "" },
  _valorRef: { vascon: 0, desconto: 0, titulo: "", inicio: "", fim: "" },
  _colsValores: null,

  async load(semanaId) {
    try {
      await SP.init();
      const datas = SP.getWeekDates(semanaId);
      const ini = datas[0].toISOString().slice(0, 10);
      const fim = datas[4].toISOString().slice(0, 10);

      this._setValAny(ini, "relDataIni");
      this._setValAny(fim, "relDataFim");
      this._bindControles();

      await this._carregarValorReferencia(ini, fim);
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

  _pick(obj, ...keys) {
    for (const k of keys.filter(Boolean)) {
      const v = SP.pick ? SP.pick(obj, k) : obj?.[k];
      if (v !== undefined && v !== null && String(v) !== "") return v;
    }
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

  _moeda(v) {
    if (AdminUtils.moeda) return AdminUtils.moeda(v);
    if (v === null || v === undefined || String(v).trim() === "") return null;
    let s = String(v).replace(/R\$/gi, "").replace(/\s/g, "").trim();
    if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
    s = s.replace(/[^0-9.-]/g, "");
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  },

  _num(v) {
    const n = this._moeda(v);
    return n === null ? 0 : n;
  },

  _periodosSobrepoem(iniA, fimA, iniB, fimB) {
    if (!iniA || !fimA || !iniB || !fimB) return false;
    return iniA <= fimB && fimA >= iniB;
  },

  async _resolverColunasValores() {
    if (this._colsValores) return this._colsValores;

    const fallback = {
      titulo: "Title",
      inicio: "Data_Inicio",
      fim: "Data_Fim",
      vascon: "Valor_Vascon",
      desconto: "Valor_Desconto_Funcionário",
      obs: "Observacao",
      ativo: "Ativo"
    };

    try {
      const spCols = typeof SP._resolveColunasValores === "function"
        ? await SP._resolveColunasValores().catch(() => ({}))
        : {};

      const siteId = await SP.getSiteId();
      const listId = await SP.getListId("Valores de Refeição");
      const data = await SP.graph("GET", `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`);
      const cols = data?.value || [];

      const byExact = (...cands) => {
        const normCands = cands.filter(Boolean).map(x => this._norm(x));
        const found = cols.find(c => normCands.includes(this._norm(c.name)) || normCands.includes(this._norm(c.displayName)));
        return found?.name || null;
      };
      const byIncludes = (...terms) => {
        const nt = terms.map(t => this._norm(t)).filter(Boolean);
        const found = cols.find(c => {
          const n = this._norm(`${c.name} ${c.displayName}`);
          return nt.every(t => n.includes(t));
        });
        return found?.name || null;
      };

      // Prioriza o nome interno real vindo de /columns; o SharePoint pode codificar
      // nomes com acento, e é esse nome que precisa ir em PATCH/leituras.
      this._colsValores = {
        titulo: byExact("Title", "Título", "Titulo") || spCols.titulo || fallback.titulo,
        inicio: byExact("Data_Inicio", "Data Inicio", "Data início") || byIncludes("data", "inicio") || spCols.inicio || fallback.inicio,
        fim: byExact("Data_Fim", "Data Fim") || byIncludes("data", "fim") || spCols.fim || fallback.fim,
        vascon: byExact("Valor_Vascon", "Valor Vascon", "Vascon") || byIncludes("vascon") || spCols.vascon || fallback.vascon,
        desconto: byExact("Valor_Desconto_Funcionário", "Valor_Desconto_Funcionario", "Desconto Funcionário", "Desconto Funcionario") || byIncludes("desconto", "funcionario") || spCols.desconto || fallback.desconto,
        obs: byExact("Observacao", "Observação") || byIncludes("observ") || spCols.obs || fallback.obs,
        ativo: byExact("Ativo") || spCols.ativo || fallback.ativo
      };
      return this._colsValores;
    } catch (e) {
      console.warn("[Relatórios] Falha ao resolver colunas de valores. Usando fallback.", e);
      this._colsValores = fallback;
      return this._colsValores;
    }
  },

  async _carregarValorReferencia(ini, fim) {
    try {
      await this._resolverColunasValores();
      const c = this._colsValores || {};
      const valores = await SP.getValoresRefeicao(false);
      const norm = valores.map(v => ({
        raw: v,
        id: v.id || "",
        titulo: this._pick(v, c.titulo, "Title", "Titulo", "Título") || "Valor refeição",
        inicio: this._dateISO(this._pick(v, c.inicio, "Data_Inicio", "DataInicio")),
        fim: this._dateISO(this._pick(v, c.fim, "Data_Fim", "DataFim")),
        vascon: this._num(this._pick(v, c.vascon, "Valor_Vascon", "ValorVascon")),
        desconto: this._num(this._pick(v, c.desconto, "Valor_Desconto_Funcionário", "Valor_Desconto_Funcionario", "Valor_Desconto", "Desconto")),
        ativo: SP.isTrue ? SP.isTrue(this._pick(v, c.ativo, "Ativo")) : !!this._pick(v, c.ativo, "Ativo")
      }));

      const escolhido =
        norm.find(v => v.ativo && this._periodosSobrepoem(ini, fim, v.inicio, v.fim)) ||
        norm.find(v => this._periodosSobrepoem(ini, fim, v.inicio, v.fim)) ||
        norm.find(v => v.ativo) ||
        norm[0] ||
        { vascon: 0, desconto: 0, titulo: "Sem valor cadastrado", inicio: "", fim: "" };

      this._valorRef = escolhido;
      this._setTxtAny(this._money(escolhido.vascon), "relValorVascon", "relVasconUnit", "relValorUnitarioVascon");
    } catch (e) {
      console.warn("[Relatórios] Não foi possível buscar valor Vascon:", e);
      this._valorRef = { vascon: 0, desconto: 0, titulo: "Sem valor cadastrado", inicio: "", fim: "" };
    }
  },

  _dataPedido(p) {
    const direta = this._dateISO(this._pick(p, "Data_Hora", "Data", "Data_Referencia"));
    if (direta) return direta;
    const semana = this._pick(p, "Semana_id", "Semana");
    const dia = this._pick(p, "Dia");
    if (semana && dia && SP.getDataRefBySemanaDia) return SP.getDataRefBySemanaDia(semana, dia);
    return "";
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

  _nfInformada() {
    return this._moeda(this._valAny(
      "relNfVascon",
      "relNFVascon",
      "relNfVasconRecebida",
      "relTotalNFVascon",
      "relTotalNfVascon",
      "relTotalNF",
      "relTotalNf",
      "nfTotalRelatorio",
      "nfTotalDigitado"
    ));
  },

  _conf() {
    return this._pedidos.filter(p => this._isConf(p));
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
    else if (tipo === "ccfunc") this._renderPorCCFunc(conf, wrap);
  },

  _linhaResumoValor() {
    const v = this._valorRef || {};
    return `<div class="alert alert-info" style="margin-bottom:.8rem;line-height:1.6">
      Valor Vascon aplicado no cálculo: <b>${this._money(v.vascon || 0)}</b>
      ${v.titulo ? ` — ${AdminUtils.esc(v.titulo)}` : ""}
      ${v.inicio && v.fim ? ` (${this._brDate(v.inicio)} a ${this._brDate(v.fim)})` : ""}.
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
      <div class="table-wrap">
        <table class="table">
          <thead><tr><th>Data</th><th>Principal</th><th>Light</th><th>Carne</th><th>Massa</th><th>Lanche</th><th>Total</th><th>Total Vascon</th></tr></thead>
          <tbody>
            ${sorted.length ? sorted.map(([data, lista]) => `<tr>
              <td>${this._brDate(data)}</td>
              <td>${this._countOp(lista, "principal")}</td>
              <td>${this._countOp(lista, "light")}</td>
              <td>${this._countOp(lista, "carne")}</td>
              <td>${this._countOp(lista, "massa")}</td>
              <td>${this._countOp(lista, "lanche")}</td>
              <td><strong>${lista.length}</strong></td>
              <td>${this._money(lista.length * unit)}</td>
            </tr>`).join("") + `<tr style="border-top:2px solid rgba(255,255,255,.15)">
              <td><strong>Total</strong></td>
              <td>${this._countOp(conf, "principal")}</td>
              <td>${this._countOp(conf, "light")}</td>
              <td>${this._countOp(conf, "carne")}</td>
              <td>${this._countOp(conf, "massa")}</td>
              <td>${this._countOp(conf, "lanche")}</td>
              <td><strong>${conf.length}</strong></td>
              <td><strong>${this._money(conf.length * unit)}</strong></td>
            </tr>` : `<tr><td colspan="8" class="empty-cell">Nenhum pedido no período.</td></tr>`}
          </tbody>
        </table>
      </div>`;
  },

  _agruparPorCC(conf) {
    const mapa = {};
    conf.forEach(p => {
      const cc = this._pick(p, "Centro_Custo", "Setor", "Departamento") || "Sem CC";
      if (!mapa[cc]) mapa[cc] = [];
      mapa[cc].push(p);
    });
    return Object.entries(mapa)
      .sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0]))
      .map(([cc, lista]) => ({ cc, lista }));
  },

  _renderPorCC(conf, wrap) {
    const sorted = this._agruparPorCC(conf);
    const unit = this._valorRef.vascon || 0;
    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">🏢 Por centro de custo — período ${this._brDate(this._periodo.ini)} a ${this._brDate(this._periodo.fim)}</div>
      ${this._linhaResumoValor()}
      <div class="table-wrap">
        <table class="table">
          <thead><tr><th>Centro de Custo</th><th>Principal</th><th>Light</th><th>Carne</th><th>Massa</th><th>Lanche</th><th>Total</th><th>Total Vascon</th></tr></thead>
          <tbody>
            ${sorted.length ? sorted.map(({ cc, lista }) => `<tr>
              <td>${AdminUtils.esc(cc)}</td>
              <td>${this._countOp(lista, "principal")}</td>
              <td>${this._countOp(lista, "light")}</td>
              <td>${this._countOp(lista, "carne")}</td>
              <td>${this._countOp(lista, "massa")}</td>
              <td>${this._countOp(lista, "lanche")}</td>
              <td><strong>${lista.length}</strong></td>
              <td>${this._money(lista.length * unit)}</td>
            </tr>`).join("") : `<tr><td colspan="8" class="empty-cell">Nenhum pedido no período.</td></tr>`}
          </tbody>
        </table>
      </div>`;
  },

  _renderPorCCFunc(conf, wrap) {
    const mapa = {};
    conf.forEach(p => {
      const cc = this._pick(p, "Centro_Custo", "Setor", "Departamento") || "Sem CC";
      const nome = this._pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "Desconhecido";
      const key = `${cc}||${nome}`;
      if (!mapa[key]) mapa[key] = { cc, nome, lista: [] };
      mapa[key].lista.push(p);
    });

    const sorted = Object.values(mapa).sort((a, b) => a.cc.localeCompare(b.cc) || a.nome.localeCompare(b.nome));
    const unitDesc = this._valorRef.desconto || 0;
    let ccAtual = null;
    const linhas = sorted.map(({ cc, nome, lista }) => {
      let cabecalho = "";
      if (cc !== ccAtual) {
        ccAtual = cc;
        const totalCC = sorted.filter(x => x.cc === cc).reduce((s, x) => s + x.lista.length, 0);
        cabecalho = `<tr style="background:rgba(255,255,255,.06)"><td colspan="5" style="font-weight:700;color:#fff">🏢 ${AdminUtils.esc(cc)} — ${totalCC} refeições</td></tr>`;
      }
      return cabecalho + `<tr>
        <td style="padding-left:1.5rem">${AdminUtils.esc(nome)}</td>
        <td>${AdminUtils.esc(cc)}</td>
        <td>${lista.length}</td>
        <td>${this._money(lista.length * unitDesc)}</td>
        <td style="font-size:.78rem;color:rgba(143,170,210,.6)">${[...new Set(lista.map(p => this._dataPedido(p)).filter(Boolean))].sort().map(d => this._brDate(d)).join(", ")}</td>
      </tr>`;
    }).join("");

    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">👤 Por CC e funcionário — período ${this._brDate(this._periodo.ini)} a ${this._brDate(this._periodo.fim)}</div>
      <div class="alert alert-info" style="margin-bottom:.8rem">Desconto funcionário aplicado: <b>${this._money(unitDesc)}</b> por refeição.</div>
      <div class="table-wrap"><table class="table">
        <thead><tr><th>Colaborador</th><th>Centro de Custo</th><th>Total refeições</th><th>Desconto total</th><th>Datas</th></tr></thead>
        <tbody>${linhas || `<tr><td colspan="5" class="empty-cell">Nenhum pedido no período.</td></tr>`}</tbody>
      </table></div>`;
  },

  _linhasRelatorio(tipo, conf) {
    if (tipo === "dia") {
      const mapa = {};
      conf.forEach(p => {
        const d = this._dataPedido(p) || "—";
        if (!mapa[d]) mapa[d] = [];
        mapa[d].push(p);
      });
      return Object.entries(mapa).sort(([a], [b]) => a.localeCompare(b)).map(([data, lista]) => ({
        Data: data,
        Principal: this._countOp(lista, "principal"),
        Light: this._countOp(lista, "light"),
        Carne: this._countOp(lista, "carne"),
        Massa: this._countOp(lista, "massa"),
        Lanche: this._countOp(lista, "lanche"),
        Total_Refeicoes: lista.length,
        Valor_Unitario_Vascon: this._valorRef.vascon || 0,
        Total_Vascon: lista.length * (this._valorRef.vascon || 0)
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
        Valor_Unitario_Vascon: this._valorRef.vascon || 0,
        Total_Vascon: lista.length * (this._valorRef.vascon || 0)
      }));
    }

    const mapa = {};
    conf.forEach(p => {
      const cc = this._pick(p, "Centro_Custo", "Setor", "Departamento") || "Sem CC";
      const nome = this._pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "Desconhecido";
      const key = `${cc}||${nome}`;
      if (!mapa[key]) mapa[key] = { cc, nome, lista: [] };
      mapa[key].lista.push(p);
    });

    return Object.values(mapa).sort((a, b) => a.cc.localeCompare(b.cc) || a.nome.localeCompare(b.nome)).map(({ cc, nome, lista }) => ({
      Centro_Custo: cc,
      Colaborador: nome,
      Total_Refeicoes: lista.length,
      Valor_Desconto_Funcionario: this._valorRef.desconto || 0,
      Desconto_Total: lista.length * (this._valorRef.desconto || 0),
      Periodo_Ini: this._periodo.ini,
      Periodo_Fim: this._periodo.fim
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

  _styleWorksheet(ws, options = {}) {
    const headerFill = "FF09213F";
    const redFill = "FFC0281C";
    const lightFill = "FFEAF1FF";
    const borderColor = "FFB8C7DD";

    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws.autoFilter = { from: "A1", to: ws.getRow(1).getCell(ws.columnCount).address };

    ws.getRow(1).height = 24;
    ws.getRow(1).eachCell(cell => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: headerFill } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.border = { bottom: { style: "thin", color: { argb: redFill } } };
    });

    ws.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      row.eachCell(cell => {
        cell.border = {
          top: { style: "thin", color: { argb: borderColor } },
          left: { style: "thin", color: { argb: borderColor } },
          bottom: { style: "thin", color: { argb: borderColor } },
          right: { style: "thin", color: { argb: borderColor } }
        };
        cell.alignment = { vertical: "middle" };
      });
      const first = String(row.getCell(1).value || "").toUpperCase();
      if (first === "TOTAL") {
        row.font = { bold: true };
        row.eachCell(cell => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: lightFill } };
        });
      }
    });

    (options.currencyCols || []).forEach(col => ws.getColumn(col).numFmt = '"R$" #,##0.00');
    (options.percentCols || []).forEach(col => ws.getColumn(col).numFmt = '0.00%');
    (options.integerCols || []).forEach(col => ws.getColumn(col).numFmt = '0');

    ws.columns.forEach((col, idx) => {
      let max = 12;
      col.eachCell({ includeEmpty: true }, cell => {
        const v = cell.value?.formula ? cell.value.result : cell.value;
        max = Math.max(max, String(v ?? "").length + 2);
      });
      col.width = Math.min(Math.max(max, idx === 0 ? 18 : 12), 42);
    });
  },

  _addRelatorioSheet(wb, tipo, linhas) {
    const ws = wb.addWorksheet("Relatorio", { properties: { tabColor: { argb: "FFC0281C" } } });
    const headers = Object.keys(linhas[0] || {});
    ws.addRow(headers);

    linhas.forEach((obj, idx) => {
      const row = ws.addRow(headers.map(h => obj[h]));
      const r = row.number;
      const totalCol = headers.indexOf("Total_Refeicoes") + 1;
      const unitVasconCol = headers.indexOf("Valor_Unitario_Vascon") + 1;
      const totalVasconCol = headers.indexOf("Total_Vascon") + 1;
      const descUnitCol = headers.indexOf("Valor_Desconto_Funcionario") + 1;
      const descTotalCol = headers.indexOf("Desconto_Total") + 1;

      if (totalCol > 0 && unitVasconCol > 0 && totalVasconCol > 0) {
        row.getCell(totalVasconCol).value = {
          formula: `${ws.getColumn(totalCol).letter}${r}*${ws.getColumn(unitVasconCol).letter}${r}`,
          result: obj.Total_Vascon
        };
      }
      if (totalCol > 0 && descUnitCol > 0 && descTotalCol > 0) {
        row.getCell(descTotalCol).value = {
          formula: `${ws.getColumn(totalCol).letter}${r}*${ws.getColumn(descUnitCol).letter}${r}`,
          result: obj.Desconto_Total
        };
      }
    });

    const totalRow = ws.addRow([]);
    totalRow.getCell(1).value = "TOTAL";
    headers.forEach((h, i) => {
      const col = i + 1;
      const letter = ws.getColumn(col).letter;
      if (["Principal", "Light", "Carne", "Massa", "Lanche", "Total_Refeicoes", "Total_Vascon", "Desconto_Total"].includes(h)) {
        const sum = linhas.reduce((acc, obj) => acc + Number(obj[h] || 0), 0);
        totalRow.getCell(col).value = { formula: `SUM(${letter}2:${letter}${linhas.length + 1})`, result: sum };
      }
    });

    const currencyCols = [];
    ["Valor_Unitario_Vascon", "Total_Vascon", "Valor_Desconto_Funcionario", "Desconto_Total"].forEach(h => {
      const i = headers.indexOf(h) + 1;
      if (i > 0) currencyCols.push(i);
    });
    const integerCols = [];
    ["Principal", "Light", "Carne", "Massa", "Lanche", "Total_Refeicoes"].forEach(h => {
      const i = headers.indexOf(h) + 1;
      if (i > 0) integerCols.push(i);
    });

    this._styleWorksheet(ws, { currencyCols, integerCols });
    return ws;
  },

  _addRateioSheet(wb, conf) {
    const ws = wb.addWorksheet("Rateio Vascon", { properties: { tabColor: { argb: "FF09213F" } } });
    const grupos = this._agruparPorCC(conf);
    const nf = this._nfInformada();
    const nfValor = nf === null ? 0 : nf;
    const unit = this._valorRef.vascon || 0;

    ws.addRow([
      "Centro_Custo",
      "Total_Refeicoes",
      "Valor_Unitario_Vascon",
      "Total_Rateado_Sistema",
      "Percentual_Rateio",
      "NF_Informada_Total",
      "Rateio_NF",
      "Diferenca_NF_vs_Sistema"
    ]);

    if (!grupos.length) grupos.push({ cc: "Sem dados", lista: [] });
    grupos.forEach(({ cc, lista }) => ws.addRow([cc, lista.length, unit, null, null, null, null, null]));

    const firstData = 2;
    const lastData = grupos.length + 1;
    const totalRow = lastData + 1;
    ws.addRow(["TOTAL", null, unit, null, null, nfValor, null, null]);

    for (let r = firstData; r <= lastData; r++) {
      const totalSistema = Number(ws.getCell(`B${r}`).value || 0) * unit;
      const percentual = grupos.reduce((s, g) => s + g.lista.length, 0) ? Number(ws.getCell(`B${r}`).value || 0) / grupos.reduce((s, g) => s + g.lista.length, 0) : 0;
      const rateioNF = nfValor * percentual;
      ws.getCell(`D${r}`).value = { formula: `B${r}*C${r}`, result: totalSistema };
      ws.getCell(`E${r}`).value = { formula: `IF($B$${totalRow}=0,0,B${r}/$B$${totalRow})`, result: percentual };
      ws.getCell(`F${r}`).value = { formula: `$F$${totalRow}`, result: nfValor };
      ws.getCell(`G${r}`).value = { formula: `$F$${totalRow}*E${r}`, result: rateioNF };
      ws.getCell(`H${r}`).value = { formula: `G${r}-D${r}`, result: rateioNF - totalSistema };
    }

    ws.getCell(`B${totalRow}`).value = { formula: `SUM(B${firstData}:B${lastData})`, result: grupos.reduce((s, g) => s + g.lista.length, 0) };
    ws.getCell(`D${totalRow}`).value = { formula: `SUM(D${firstData}:D${lastData})`, result: grupos.reduce((s, g) => s + g.lista.length, 0) * unit };
    ws.getCell(`E${totalRow}`).value = { formula: `SUM(E${firstData}:E${lastData})`, result: 1 };
    ws.getCell(`G${totalRow}`).value = { formula: `SUM(G${firstData}:G${lastData})`, result: nfValor };
    ws.getCell(`H${totalRow}`).value = { formula: `SUM(H${firstData}:H${lastData})`, result: nfValor - (grupos.reduce((s, g) => s + g.lista.length, 0) * unit) };

    this._styleWorksheet(ws, { currencyCols: [3, 4, 6, 7, 8], percentCols: [5], integerCols: [2] });
    return ws;
  },

  _addParametrosSheet(wb, conf) {
    const ws = wb.addWorksheet("Parametros", { properties: { tabColor: { argb: "FF40D090" } } });
    const nf = this._nfInformada();
    const total = conf.length;
    const unit = this._valorRef.vascon || 0;
    const esperado = total * unit;

    ws.addRow(["Parametro", "Valor"]);
    ws.addRow(["Período inicial", this._periodo.ini]);
    ws.addRow(["Período final", this._periodo.fim]);
    ws.addRow(["Valor cadastrado", this._valorRef.titulo || ""]);
    ws.addRow(["Vigência valor", `${this._valorRef.inicio || ""} a ${this._valorRef.fim || ""}`]);
    ws.addRow(["Valor unitário Vascon", unit]);
    ws.addRow(["Valor desconto funcionário", this._valorRef.desconto || 0]);
    ws.addRow(["Total refeições confirmadas", total]);
    ws.addRow(["Total sistema Vascon", { formula: "B6*B8", result: esperado }]);
    ws.addRow(["NF Vascon informada", nf === null ? 0 : nf]);
    ws.addRow(["Diferença NF x sistema", { formula: "B10-B9", result: nf === null ? -esperado : nf - esperado }]);
    ws.addRow([]);
    ws.addRow(["Memória de cálculo", ""]);
    ws.addRow(["Total_Rateado_Sistema", "Total_Refeicoes * Valor_Unitario_Vascon"]);
    ws.addRow(["Percentual_Rateio", "Total_Refeicoes do CC / Total_Refeicoes geral"]);
    ws.addRow(["Rateio_NF", "NF_Informada_Total * Percentual_Rateio"]);
    ws.addRow(["Diferenca_NF_vs_Sistema", "Rateio_NF - Total_Rateado_Sistema"]);

    this._styleWorksheet(ws, { currencyCols: [2], integerCols: [] });
    ws.getColumn(1).width = 34;
    ws.getColumn(2).width = 46;
    [6, 7, 9, 10, 11].forEach(r => ws.getCell(`B${r}`).numFmt = '"R$" #,##0.00');
    ws.getRow(13).font = { bold: true };
    return ws;
  },

  async _exportar() {
    try {
      const ini = this._valAny("relDataIni") || this._periodo.ini;
      const fim = this._valAny("relDataFim") || this._periodo.fim;
      if (!ini || !fim) { AdminUtils.toast("Informe o período.", "error"); return; }

      if (ini !== this._periodo.ini || fim !== this._periodo.fim) await this._buscar(ini, fim);
      else await this._carregarValorReferencia(this._periodo.ini, this._periodo.fim);

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

      this._addRelatorioSheet(wb, tipo, linhas);
      this._addRateioSheet(wb, conf);
      this._addParametrosSheet(wb, conf);

      const buffer = await wb.xlsx.writeBuffer();
      this._downloadBuffer(buffer, `relatorio-vascon-${tipo}-${this._periodo.ini}-${this._periodo.fim}.xlsx`);
      AdminUtils.toast("Excel exportado com formatação e fórmulas.", "success");
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
