// admin-relatorios.js — Relatórios gerenciais + rateio Vascon automático
// Correções: load garantido, valor Vascon detectado da lista Valores de Refeição e Excel com aba Rateio Vascon

const AdminRelatorios = window.AdminRelatorios = {
  _pedidos: [],
  _periodo: { ini: "", fim: "" },
  _valorRef: { vascon: 0, desconto: 0, titulo: "", inicio: "", fim: "" },

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

  _pick(obj, ...keys) {
    for (const k of keys) {
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

  async _carregarValorReferencia(ini, fim) {
    try {
      const valores = await SP.getValoresRefeicao(false);
      const norm = valores.map(v => ({
        raw: v,
        id: v.id || "",
        titulo: this._pick(v, "Title", "Titulo", "Título") || "Valor refeição",
        inicio: this._dateISO(this._pick(v, "Data_Inicio", "DataInicio")),
        fim: this._dateISO(this._pick(v, "Data_Fim", "DataFim")),
        vascon: this._num(this._pick(v, "Valor_Vascon", "ValorVascon")),
        desconto: this._num(this._pick(v, "Valor_Desconto_Funcionário", "Valor_Desconto_Funcionario", "Valor_Desconto", "Desconto")),
        ativo: SP.isTrue ? SP.isTrue(this._pick(v, "Ativo")) : !!this._pick(v, "Ativo")
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
      if (wrap) {
        wrap.innerHTML = `<div class="alert" style="background:rgba(220,50,50,.1);color:#ff8080">Erro: ${AdminUtils.esc(e.message || e)}</div>`;
      }
    }
  },

  _isConf(p) {
    const s = AdminUtils.norm(this._pick(p, "Status") || "");
    return ["confirmado", "extra", "aprovado", "travado"].includes(s) ||
           (SP.isTrue && SP.isTrue(this._pick(p, "Confirmado")));
  },

  _opcao(p) {
    return AdminUtils.norm(this._pick(p, "Opcao", "Opção") || "");
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
      "nfTotalRelatorio"
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
          <thead><tr><th>Data</th><th>Principal</th><th>Light</th><th>Carne</th><th>Massa</th><th>Lanche</th><th>Total</th><th>Valor Vascon</th></tr></thead>
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

  _renderPorCC(conf, wrap) {
    const sorted = this._agruparPorCC(conf);
    const unit = this._valorRef.vascon || 0;

    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">🏢 Por centro de custo — período ${this._brDate(this._periodo.ini)} a ${this._brDate(this._periodo.fim)}</div>
      ${this._linhaResumoValor()}
      <div class="table-wrap">
        <table class="table">
          <thead><tr><th>Centro de Custo</th><th>Principal</th><th>Light</th><th>Carne</th><th>Massa</th><th>Lanche</th><th>Total</th><th>Rateio Vascon</th></tr></thead>
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
        cabecalho = `<tr style="background:rgba(255,255,255,.06)">
          <td colspan="5" style="font-weight:700;color:#fff">🏢 ${AdminUtils.esc(cc)} — ${totalCC} refeições</td>
        </tr>`;
      }
      return cabecalho + `<tr>
        <td style="padding-left:1.5rem">${AdminUtils.esc(nome)}</td>
        <td>${AdminUtils.esc(cc)}</td>
        <td>${lista.length}</td>
        <td>${this._money(lista.length * unitDesc)}</td>
        <td style="font-size:.78rem;color:rgba(143,170,210,.6)">
          ${[...new Set(lista.map(p => this._dataPedido(p)).filter(Boolean))].sort().map(d => this._brDate(d)).join(", ")}
        </td>
      </tr>`;
    }).join("");

    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">👤 Por CC e funcionário — período ${this._brDate(this._periodo.ini)} a ${this._brDate(this._periodo.fim)}</div>
      <div class="alert alert-info" style="margin-bottom:.8rem">Desconto funcionário aplicado: <b>${this._money(unitDesc)}</b> por refeição.</div>
      <div class="table-wrap">
        <table class="table">
          <thead><tr><th>Colaborador</th><th>Centro de Custo</th><th>Total refeições</th><th>Desconto total</th><th>Datas</th></tr></thead>
          <tbody>${linhas || `<tr><td colspan="5" class="empty-cell">Nenhum pedido no período.</td></tr>`}</tbody>
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

  _criarAbaRateio(wb, conf) {
    const grupos = this._agruparPorCC(conf);
    const nf = this._nfInformada();
    const nfValor = nf === null ? 0 : nf;
    const unit = this._valorRef.vascon || 0;

    const aoa = [[
      "Centro_Custo",
      "Total_Refeicoes",
      "Valor_Unitario_Vascon",
      "Total_Rateado_Sistema",
      "Percentual_Rateio",
      "NF_Informada_Total",
      "Rateio_NF",
      "Diferenca_NF_vs_Sistema"
    ]];

    grupos.forEach(({ cc, lista }) => {
      aoa.push([cc, lista.length, unit, null, null, nfValor, null, null]);
    });

    if (!grupos.length) aoa.push(["Sem dados", 0, unit, null, null, nfValor, null, null]);

    const totalRow = aoa.length + 1;
    aoa.push(["TOTAL", null, unit, null, null, nfValor, null, null]);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const lastDataRow = aoa.length - 1;

    for (let r = 2; r <= lastDataRow; r++) {
      ws[`D${r}`] = { t: "n", f: `B${r}*C${r}` };
      ws[`E${r}`] = { t: "n", f: `IF(SUM($B$2:$B$${lastDataRow})=0,0,B${r}/SUM($B$2:$B$${lastDataRow}))` };
      ws[`G${r}`] = { t: "n", f: `F${r}*E${r}` };
      ws[`H${r}`] = { t: "n", f: `G${r}-D${r}` };
    }

    ws[`B${totalRow}`] = { t: "n", f: `SUM(B2:B${lastDataRow})` };
    ws[`D${totalRow}`] = { t: "n", f: `SUM(D2:D${lastDataRow})` };
    ws[`E${totalRow}`] = { t: "n", f: `SUM(E2:E${lastDataRow})` };
    ws[`G${totalRow}`] = { t: "n", f: `SUM(G2:G${lastDataRow})` };
    ws[`H${totalRow}`] = { t: "n", f: `SUM(H2:H${lastDataRow})` };

    ws["!cols"] = [
      { wch: 34 }, { wch: 16 }, { wch: 22 }, { wch: 24 },
      { wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 24 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Rateio Vascon");
  },

  _criarAbaParametros(wb, conf) {
    const nf = this._nfInformada();
    const total = conf.length;
    const unit = this._valorRef.vascon || 0;
    const esperado = total * unit;

    const ws = XLSX.utils.aoa_to_sheet([
      ["Parametro", "Valor"],
      ["Periodo inicial", this._periodo.ini],
      ["Periodo final", this._periodo.fim],
      ["Valor cadastrado", this._valorRef.titulo || ""],
      ["Vigência valor", `${this._valorRef.inicio || ""} a ${this._valorRef.fim || ""}`],
      ["Valor unitário Vascon", unit],
      ["Valor desconto funcionário", this._valorRef.desconto || 0],
      ["Total refeições confirmadas", total],
      ["Total sistema Vascon", esperado],
      ["NF Vascon informada", nf === null ? "" : nf],
      ["Diferença NF x sistema", nf === null ? "" : nf - esperado]
    ]);
    ws["!cols"] = [{ wch: 30 }, { wch: 35 }];
    XLSX.utils.book_append_sheet(wb, ws, "Parametros");
  },

  async _exportar() {
    if (typeof XLSX === "undefined") {
      AdminUtils.toast("Biblioteca XLSX não carregou.", "error");
      return;
    }

    try {
      const ini = this._valAny("relDataIni") || this._periodo.ini;
      const fim = this._valAny("relDataFim") || this._periodo.fim;
      if (ini && fim && (ini !== this._periodo.ini || fim !== this._periodo.fim)) {
        await this._buscar(ini, fim);
      } else {
        await this._carregarValorReferencia(this._periodo.ini, this._periodo.fim);
      }

      const tipo = this._valAny("relTipo") || "dia";
      const conf = this._conf();
      const linhas = this._linhasRelatorio(tipo, conf);

      if (!linhas.length) {
        AdminUtils.toast("Nenhum dado para exportar.", "info");
        return;
      }

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(linhas);
      ws["!cols"] = Object.keys(linhas[0] || {}).map(k => ({ wch: Math.max(14, String(k).length + 4) }));
      XLSX.utils.book_append_sheet(wb, ws, "Relatorio");

      this._criarAbaRateio(wb, conf);
      this._criarAbaParametros(wb, conf);

      XLSX.writeFile(wb, `relatorio-vascon-${tipo}-${this._periodo.ini}-${this._periodo.fim}.xlsx`);
      AdminUtils.toast("Excel exportado com aba Rateio Vascon.", "success");
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
      "relTotalNfVascon", "relTotalNF", "relTotalNf", "nfTotalRelatorio"
    ].forEach(id => bind(id, "input", () => this._renderCards()));
  }
};
