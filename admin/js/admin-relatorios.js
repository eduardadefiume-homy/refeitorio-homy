// ============================================================
// admin-relatorios.js — Relatórios oficiais do Admin Homy
// v: base-centralizada-v10-11-20260630
//
// Regra:
// - Se houver Fechamento Diario Refeitorio no período, o relatório usa o fechamento.
// - Dias ainda não fechados usam Pedidos calculados pela camada central de regras.
// - Pedidos brutos continuam auditáveis, mas não derrubam relatório oficial.
// ============================================================
const AdminRelatorios = window.AdminRelatorios = {
  _dados: null,
  _periodo: { ini: "", fim: "" },

  async load(semanaId) {
    const datas = SP.getWeekDates(semanaId);
    const ini = datas[0].toISOString().slice(0, 10);
    const fim = datas[4].toISOString().slice(0, 10);
    AdminUtils.setVal("relDataIni", ini);
    AdminUtils.setVal("relDataFim", fim);
    this._bindControles();
    await this._buscar(ini, fim);
  },

  async _buscar(ini, fim) {
    this._periodo = { ini, fim };
    const wrap = document.getElementById("relConteudo");
    if (wrap) wrap.innerHTML = `<div class="alert alert-info">Carregando relatório oficial...</div>`;

    try {
      await SP.init();
      this._dados = await SP.getRelatorioOficialPeriodo(ini, fim, { force: true });
      this._renderCards();
      this._renderTipo();
    } catch (e) {
      console.error("[Relatórios]", e);
      if (wrap) wrap.innerHTML = `<div class="alert" style="background:rgba(220,50,50,.1);color:#ff8080">Erro: ${AdminUtils.esc(e.message || e)}</div>`;
    }
  },

  _totais() {
    return this._dados?.totais || { principal: 0, light: 0, carne: 0, massa: 0, lanche: 0, total: 0 };
  },

  _itens() {
    return this._dados?.itens || [];
  },

  _countOp(lista, op) {
    return (lista || []).filter(p => AdminUtils.norm(SP.pick(p, "Opcao", "opcao")) === op).length;
  },

  _renderCards() {
    const t = this._totais();
    AdminUtils.setTxt("rel-principal", t.principal || 0);
    AdminUtils.setTxt("rel-light", t.light || 0);
    AdminUtils.setTxt("rel-carne", t.carne || 0);
    AdminUtils.setTxt("rel-massa", t.massa || 0);
    AdminUtils.setTxt("rel-lanche", t.lanche || 0);
    AdminUtils.setTxt("rel-total", t.total || 0);
  },

  _renderTipo() {
    const tipo = AdminUtils.getVal("relTipo") || "dia";
    const wrap = document.getElementById("relConteudo");
    if (!wrap) return;

    if (tipo === "dia") this._renderPorDia(wrap);
    else if (tipo === "cc") this._renderPorCC(wrap);
    else if (tipo === "ccfunc") this._renderPorCCFunc(wrap);
  },

  _fonteInfoHtml() {
    const d = this._dados || {};
    const fechamentos = d.fechamentosUsados?.length || 0;
    const classe = fechamentos ? "alert-success" : "alert-info";
    const msg = fechamentos
      ? `Relatório oficial usando ${fechamentos} fechamento(s) salvo(s). Dias sem fechamento usam Pedidos calculados.`
      : `Relatório calculado pelos Pedidos, pois não há fechamento oficial no período.`;
    return `<div class="alert ${classe}" style="margin-bottom:.8rem">${AdminUtils.esc(msg)}</div>`;
  },

  _renderPorDia(wrap) {
    const dias = this._dados?.dias || [];
    const t = this._totais();
    wrap.innerHTML = `
      ${this._fonteInfoHtml()}
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">📅 Por dia — período ${this._periodo.ini} a ${this._periodo.fim}</div>
      <div class="table-wrap">
        <table class="table">
          <thead><tr><th>Data</th><th>Fonte</th><th>Principal</th><th>Light</th><th>Carne</th><th>Massa</th><th>Lanche</th><th>Total</th></tr></thead>
          <tbody>
            ${dias.length ? dias.map(d => `<tr>
              <td>${AdminUtils.esc(d.data || "—")}</td>
              <td>${AdminUtils.esc(d.fonte || "—")}</td>
              <td>${Number(d.principal || 0)}</td>
              <td>${Number(d.light || 0)}</td>
              <td>${Number(d.carne || 0)}</td>
              <td>${Number(d.massa || 0)}</td>
              <td>${Number(d.lanche || 0)}</td>
              <td><strong>${Number(d.total || 0)}</strong></td>
            </tr>`).join("") + `<tr style="border-top:2px solid rgba(255,255,255,.15)">
              <td colspan="2"><strong>Total</strong></td>
              <td>${t.principal || 0}</td><td>${t.light || 0}</td><td>${t.carne || 0}</td><td>${t.massa || 0}</td><td>${t.lanche || 0}</td><td><strong>${t.total || 0}</strong></td>
            </tr>` : `<tr><td colspan="8" class="empty-cell">Nenhum dado no período.</td></tr>`}
          </tbody>
        </table>
      </div>`;
  },

  _renderPorCC(wrap) {
    const mapa = {};
    for (const p of this._itens()) {
      const cc = SP.pick(p, "Centro_Custo", "CentroCusto", "Departamento", "Setor") || "Sem CC";
      if (!mapa[cc]) mapa[cc] = [];
      mapa[cc].push(p);
    }
    const sorted = Object.entries(mapa).sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0]));
    wrap.innerHTML = `
      ${this._fonteInfoHtml()}
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">🏢 Por centro de custo — período ${this._periodo.ini} a ${this._periodo.fim}</div>
      <div class="table-wrap"><table class="table">
        <thead><tr><th>Centro de Custo</th><th>Principal</th><th>Light</th><th>Carne</th><th>Massa</th><th>Lanche</th><th>Total</th></tr></thead>
        <tbody>${sorted.length ? sorted.map(([cc, lista]) => `<tr>
          <td>${AdminUtils.esc(cc)}</td>
          <td>${this._countOp(lista, "principal")}</td><td>${this._countOp(lista, "light")}</td><td>${this._countOp(lista, "carne")}</td><td>${this._countOp(lista, "massa")}</td><td>${this._countOp(lista, "lanche")}</td><td><strong>${lista.length}</strong></td>
        </tr>`).join("") : `<tr><td colspan="7" class="empty-cell">Nenhum dado no período.</td></tr>`}</tbody>
      </table></div>`;
  },

  _renderPorCCFunc(wrap) {
    const mapa = {};
    for (const p of this._itens()) {
      const cc = SP.pick(p, "Centro_Custo", "CentroCusto", "Departamento", "Setor") || "Sem CC";
      const nome = SP.pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "Desconhecido";
      const data = SP._dataISOOperacional ? SP._dataISOOperacional(SP.pick(p, "Data_Operacao", "Data_Hora", "Data")) : String(SP.pick(p, "Data_Hora") || "").slice(0, 10);
      const key = `${cc}||${nome}`;
      if (!mapa[key]) mapa[key] = { cc, nome, lista: [], datas: new Set() };
      mapa[key].lista.push(p);
      if (data) mapa[key].datas.add(data);
    }

    const sorted = Object.values(mapa).sort((a, b) => a.cc.localeCompare(b.cc) || a.nome.localeCompare(b.nome, "pt-BR"));
    let ccAtual = null;
    const linhas = sorted.map(({ cc, nome, lista, datas }) => {
      let cabecalho = "";
      if (cc !== ccAtual) {
        ccAtual = cc;
        const totalCC = sorted.filter(x => x.cc === cc).reduce((s, x) => s + x.lista.length, 0);
        cabecalho = `<tr style="background:rgba(255,255,255,.06)"><td colspan="4" style="font-weight:700;color:#fff">🏢 ${AdminUtils.esc(cc)} — ${totalCC} refeições</td></tr>`;
      }
      return cabecalho + `<tr>
        <td style="padding-left:1.5rem">${AdminUtils.esc(nome)}</td>
        <td>${AdminUtils.esc(cc)}</td>
        <td>${lista.length}</td>
        <td style="font-size:.78rem;color:rgba(143,170,210,.6)">${[...datas].sort().join(", ")}</td>
      </tr>`;
    }).join("");

    wrap.innerHTML = `
      ${this._fonteInfoHtml()}
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">👤 Por CC e funcionário — período ${this._periodo.ini} a ${this._periodo.fim}</div>
      <div class="alert alert-info" style="margin-bottom:.8rem">Total oficial por colaborador no período. Quando o dia está fechado, a fonte é o snapshot do fechamento.</div>
      <div class="table-wrap"><table class="table"><thead><tr><th>Colaborador</th><th>Centro de Custo</th><th>Total refeições</th><th>Datas</th></tr></thead><tbody>${linhas || `<tr><td colspan="4" class="empty-cell">Nenhum dado no período.</td></tr>`}</tbody></table></div>`;
  },

  _exportar() {
    if (typeof XLSX === "undefined") { AdminUtils.toast("Biblioteca XLSX não carregou.", "error"); return; }
    const tipo = AdminUtils.getVal("relTipo") || "dia";
    const linhas = [];
    let nomeAba = "Relatorio";

    if (tipo === "dia") {
      nomeAba = "Por Dia";
      for (const d of this._dados?.dias || []) {
        linhas.push({ Data: d.data, Fonte: d.fonte, Principal: d.principal, Light: d.light, Carne: d.carne, Massa: d.massa, Lanche: d.lanche, Total: d.total });
      }
    } else if (tipo === "cc") {
      nomeAba = "Por CC";
      const mapa = {};
      for (const p of this._itens()) {
        const cc = SP.pick(p, "Centro_Custo", "CentroCusto", "Departamento", "Setor") || "Sem CC";
        if (!mapa[cc]) mapa[cc] = [];
        mapa[cc].push(p);
      }
      for (const [cc, lista] of Object.entries(mapa)) {
        linhas.push({ Centro_Custo: cc, Principal: this._countOp(lista, "principal"), Light: this._countOp(lista, "light"), Carne: this._countOp(lista, "carne"), Massa: this._countOp(lista, "massa"), Lanche: this._countOp(lista, "lanche"), Total: lista.length });
      }
    } else {
      nomeAba = "Por CC e Funcionario";
      const mapa = {};
      for (const p of this._itens()) {
        const cc = SP.pick(p, "Centro_Custo", "CentroCusto", "Departamento", "Setor") || "Sem CC";
        const nome = SP.pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "Desconhecido";
        const key = `${cc}||${nome}`;
        if (!mapa[key]) mapa[key] = { cc, nome, total: 0 };
        mapa[key].total++;
      }
      for (const r of Object.values(mapa)) linhas.push({ Centro_Custo: r.cc, Colaborador: r.nome, Total_Refeicoes: r.total, Periodo_Ini: this._periodo.ini, Periodo_Fim: this._periodo.fim });
    }

    if (!linhas.length) { AdminUtils.toast("Nenhum dado para exportar.", "info"); return; }
    const ws = XLSX.utils.json_to_sheet(linhas);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, nomeAba);
    XLSX.writeFile(wb, `relatorio-oficial-${tipo}-${this._periodo.ini}-${this._periodo.fim}.xlsx`);
    AdminUtils.toast("Excel exportado.", "success");
  },

  _bindControles() {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundRel) { el.dataset.boundRel = "1"; el.addEventListener(ev, fn); }
    };

    bind("btnBuscarRelatorio", "click", async () => {
      const ini = AdminUtils.getVal("relDataIni");
      const fim = AdminUtils.getVal("relDataFim");
      if (!ini || !fim) { AdminUtils.toast("Informe o período.", "error"); return; }
      if (ini > fim) { AdminUtils.toast("Data início maior que fim.", "error"); return; }
      await this._buscar(ini, fim);
    });
    bind("relTipo", "change", () => this._renderTipo());
    bind("btnExportarRelatorio", "click", () => this._exportar());
  }
};
