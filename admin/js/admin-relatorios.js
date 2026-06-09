// admin-relatorios.js — Relatórios inteligentes do Admin Homy
// Filtro por período, por dia, por CC, por CC+funcionário, exportação Excel

const AdminRelatorios = window.AdminRelatorios = {

  _pedidos: [],
  _periodo: { ini: "", fim: "" },

  async load(semanaId) {
    // Preenche período padrão com a semana atual
    const datas = SP.getWeekDates(semanaId);
    const ini   = datas[0].toISOString().slice(0, 10);
    const fim   = datas[4].toISOString().slice(0, 10);
    AdminUtils.setVal("relDataIni", ini);
    AdminUtils.setVal("relDataFim", fim);
    this._bindControles();
    await this._buscar(ini, fim);
  },

  async _buscar(ini, fim) {
    this._periodo = { ini, fim };
    const wrap = document.getElementById("relConteudo");
    if (wrap) wrap.innerHTML = `<div class="alert alert-info">Carregando...</div>`;

    try {
      await SP.init();
      const todos = await SP.getItems("Pedidos");

      this._pedidos = todos.filter(p => {
        // Filtra por Data_Hora ou tenta Semana_id como fallback
        const dh   = (SP.pick(p, "Data_Hora") || "").slice(0, 10);
        const ok   = dh && dh >= ini && dh <= fim;
        return ok;
      });

      this._renderCards();
      this._renderTipo();
    } catch (e) {
      const wrap = document.getElementById("relConteudo");
      if (wrap) wrap.innerHTML = `<div class="alert" style="background:rgba(220,50,50,.1);color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</div>`;
    }
  },

  _isConf(p) {
    const s = AdminUtils.norm(SP.pick(p, "Status") || "");
    return s === "confirmado" || s === "extra" || SP.isTrue(SP.pick(p, "Confirmado"));
  },

  _countOp(lista, op) {
    return lista.filter(p => AdminUtils.norm(SP.pick(p, "Opcao")) === op).length;
  },

  _renderCards() {
    const conf = this._pedidos.filter(p => this._isConf(p));
    AdminUtils.setTxt("rel-principal", this._countOp(conf, "principal"));
    AdminUtils.setTxt("rel-light",     this._countOp(conf, "light"));
    AdminUtils.setTxt("rel-carne",     this._countOp(conf, "carne"));
    AdminUtils.setTxt("rel-massa",     this._countOp(conf, "massa"));
    AdminUtils.setTxt("rel-lanche",    this._countOp(conf, "lanche"));
    AdminUtils.setTxt("rel-total",     conf.length);
  },

  _renderTipo() {
    const tipo = AdminUtils.getVal("relTipo") || "dia";
    const wrap = document.getElementById("relConteudo");
    if (!wrap) return;

    const conf = this._pedidos.filter(p => this._isConf(p));

    if (tipo === "dia")         this._renderPorDia(conf, wrap);
    else if (tipo === "cc")     this._renderPorCC(conf, wrap);
    else if (tipo === "ccfunc") this._renderPorCCFunc(conf, wrap);
  },

  // Mapa dia → data real baseado no período selecionado
  _getDiaParaData() {
    // Retorna mapa { "segunda": "2026-06-08", "terca": "2026-06-09", ... }
    // baseado no período buscado (usa a semana que contém a data de início)
    const ini = this._periodo.ini;
    if (!ini) return {};
    // Encontra a segunda-feira da semana do início
    const d = new Date(ini + "T12:00:00");
    const dow = d.getDay(); // 0=dom, 1=seg...
    const seg = new Date(d);
    seg.setDate(d.getDate() - (dow === 0 ? 6 : dow - 1));
    const DIAS = ["segunda","terca","quarta","quinta","sexta"];
    const mapa = {};
    DIAS.forEach((dia, i) => {
      const dt = new Date(seg);
      dt.setDate(seg.getDate() + i);
      mapa[dia] = dt.toISOString().slice(0, 10);
    });
    return mapa;
  },

  _renderPorDia(conf, wrap) {
    // Agrupa por campo Dia (nome do dia da semana) e mostra data real
    const diaParaData = this._getDiaParaData();
    const ORDEM_DIAS = ["segunda","terca","quarta","quinta","sexta"];

    // Também aceita pedidos agrupados por Data_Hora para compatibilidade
    const mapasDia   = {};  // { "segunda": [pedidos] }
    const mapaData   = {};  // { "2026-06-09": [pedidos] }

    conf.forEach(p => {
      const diaNorm = AdminUtils.norm(SP.pick(p,"Dia") || "");
      const dtPed   = (SP.pick(p,"Data_Hora") || "").slice(0,10);

      if (ORDEM_DIAS.includes(diaNorm)) {
        if (!mapasDia[diaNorm]) mapasDia[diaNorm] = [];
        mapasDia[diaNorm].push(p);
      } else if (dtPed) {
        if (!mapaData[dtPed]) mapaData[dtPed] = [];
        mapaData[dtPed].push(p);
      }
    });

    // Constrói linhas: primeiro pelos dias nomeados (com data real), depois por data bruta
    const linhas = [];

    ORDEM_DIAS.forEach(dia => {
      if (!mapasDia[dia]) return;
      const lista    = mapasDia[dia];
      const dataReal = diaParaData[dia] || dia;
      linhas.push({ label: dataReal, lista });
      // Remove do mapaData para não duplicar
      if (mapaData[dataReal]) delete mapaData[dataReal];
    });

    // Pedidos sem campo Dia nomeado — agrupa pela data do Data_Hora
    Object.entries(mapaData).sort(([a],[b])=>a.localeCompare(b)).forEach(([data,lista])=>{
      linhas.push({ label: data, lista });
    });

    linhas.sort((a,b)=>a.label.localeCompare(b.label));

    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">📅 Por dia — período ${this._periodo.ini} a ${this._periodo.fim}</div>
      <div class="table-wrap">
        <table class="table">
          <thead><tr><th>Data</th><th>Principal</th><th>Light</th><th>Carne</th><th>Massa</th><th>Lanche</th><th>Total</th></tr></thead>
          <tbody>
            ${linhas.length ? linhas.map(({label, lista}) => `<tr>
              <td>${label}</td>
              <td>${this._countOp(lista,"principal")}</td>
              <td>${this._countOp(lista,"light")}</td>
              <td>${this._countOp(lista,"carne")}</td>
              <td>${this._countOp(lista,"massa")}</td>
              <td>${this._countOp(lista,"lanche")}</td>
              <td><strong>${lista.length}</strong></td>
            </tr>`).join("") + `
            <tr style="border-top:2px solid rgba(255,255,255,.15)">
              <td><strong>Total semana</strong></td>
              <td>${this._countOp(conf,"principal")}</td>
              <td>${this._countOp(conf,"light")}</td>
              <td>${this._countOp(conf,"carne")}</td>
              <td>${this._countOp(conf,"massa")}</td>
              <td>${this._countOp(conf,"lanche")}</td>
              <td><strong>${conf.length}</strong></td>
            </tr>` : `<tr><td colspan="7" class="empty-cell">Nenhum pedido no período.</td></tr>`}
          </tbody>
        </table>
      </div>`;
  },

  _renderPorCC(conf, wrap) {
    const mapa = {};
    conf.forEach(p => {
      const cc = SP.pick(p, "Centro_Custo") || "Sem CC";
      if (!mapa[cc]) mapa[cc] = [];
      mapa[cc].push(p);
    });

    const sorted = Object.entries(mapa).sort((a, b) => b[1].length - a[1].length);

    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">🏢 Por centro de custo — período ${this._periodo.ini} a ${this._periodo.fim}</div>
      <div class="table-wrap">
        <table class="table">
          <thead><tr><th>Centro de Custo</th><th>Principal</th><th>Light</th><th>Carne</th><th>Massa</th><th>Lanche</th><th>Total</th></tr></thead>
          <tbody>
            ${sorted.length ? sorted.map(([cc, lista]) => `<tr>
              <td>${AdminUtils.esc(cc)}</td>
              <td>${this._countOp(lista, "principal")}</td>
              <td>${this._countOp(lista, "light")}</td>
              <td>${this._countOp(lista, "carne")}</td>
              <td>${this._countOp(lista, "massa")}</td>
              <td>${this._countOp(lista, "lanche")}</td>
              <td><strong>${lista.length}</strong></td>
            </tr>`).join("") : `<tr><td colspan="7" class="empty-cell">Nenhum pedido no período.</td></tr>`}
          </tbody>
        </table>
      </div>`;
  },

  _renderPorCCFunc(conf, wrap) {
    // Agrupa por CC e dentro por colaborador
    const mapa = {};
    conf.forEach(p => {
      const cc   = SP.pick(p, "Centro_Custo")      || "Sem CC";
      const nome = SP.pick(p, "Colaborador_nome", "Title") || "Desconhecido";
      const key  = `${cc}||${nome}`;
      if (!mapa[key]) mapa[key] = { cc, nome, lista: [] };
      mapa[key].lista.push(p);
    });

    const sorted = Object.values(mapa).sort((a, b) =>
      a.cc.localeCompare(b.cc) || b.lista.length - a.lista.length
    );

    let ccAtual = null;
    const linhas = sorted.map(({ cc, nome, lista }) => {
      let cabecalho = "";
      if (cc !== ccAtual) {
        ccAtual = cc;
        const totalCC = sorted.filter(x => x.cc === cc).reduce((s, x) => s + x.lista.length, 0);
        cabecalho = `<tr style="background:rgba(255,255,255,.06)">
          <td colspan="4" style="font-weight:700;color:#fff">🏢 ${AdminUtils.esc(cc)} — ${totalCC} refeições</td>
        </tr>`;
      }
      return cabecalho + `<tr>
        <td style="padding-left:1.5rem">${AdminUtils.esc(nome)}</td>
        <td>${AdminUtils.esc(cc)}</td>
        <td>${lista.length}</td>
        <td style="font-size:.78rem;color:rgba(143,170,210,.6)">
          ${[...new Set(lista.map(p => (SP.pick(p, "Data_Hora") || "").slice(0, 10)).filter(Boolean))].sort().join(", ")}
        </td>
      </tr>`;
    }).join("");

    wrap.innerHTML = `
      <div class="section-title" style="font-size:.95rem;margin-bottom:.7rem">👤 Por CC e funcionário — período ${this._periodo.ini} a ${this._periodo.fim}</div>
      <div class="alert alert-info" style="margin-bottom:.8rem">Total de refeições por colaborador no período — use para cálculo de desconto em folha.</div>
      <div class="table-wrap">
        <table class="table">
          <thead><tr><th>Colaborador</th><th>Centro de Custo</th><th>Total refeições</th><th>Datas</th></tr></thead>
          <tbody>${linhas || `<tr><td colspan="4" class="empty-cell">Nenhum pedido no período.</td></tr>`}</tbody>
        </table>
      </div>`;
  },

  // ── Exportação Excel ─────────────────────────────────────────
  _exportar() {
    if (typeof XLSX === "undefined") { AdminUtils.toast("Biblioteca XLSX não carregou.", "error"); return; }
    const tipo = AdminUtils.getVal("relTipo") || "dia";
    const conf = this._pedidos.filter(p => this._isConf(p));

    let linhas = [];
    let nomeAba = "Relatorio";

    if (tipo === "dia") {
      nomeAba = "Por Dia";
      const diaParaData = this._getDiaParaData();
      const ORDEM_DIAS  = ["segunda","terca","quarta","quinta","sexta"];
      const mapasDia = {};
      const mapaData = {};
      conf.forEach(p => {
        const diaNorm = AdminUtils.norm(SP.pick(p,"Dia")||"");
        const dtPed   = (SP.pick(p,"Data_Hora")||"").slice(0,10);
        if(ORDEM_DIAS.includes(diaNorm)){
          if(!mapasDia[diaNorm]) mapasDia[diaNorm]=[];
          mapasDia[diaNorm].push(p);
        } else if(dtPed){
          if(!mapaData[dtPed]) mapaData[dtPed]=[];
          mapaData[dtPed].push(p);
        }
      });
      const linhas2 = [];
      ORDEM_DIAS.forEach(dia=>{
        if(!mapasDia[dia]) return;
        const lista=mapasDia[dia];
        const data=diaParaData[dia]||dia;
        linhas2.push({label:data,lista});
        if(mapaData[data]) delete mapaData[data];
      });
      Object.entries(mapaData).sort(([a],[b])=>a.localeCompare(b)).forEach(([data,lista])=>linhas2.push({label:data,lista}));
      linhas2.sort((a,b)=>a.label.localeCompare(b.label));
      linhas = [
        { Data: `RELATÓRIO REFEITÓRIO HOMY  QUANTIDADE POR DIA` },
        { Data: `Período: ${this._periodo.ini} a ${this._periodo.fim}` },
        {},
        { Data: "Total de refeições", Principal: conf.length },
        {},
        ...linhas2.map(({label,lista})=>({
          Data:      label,
          Principal: this._countOp(lista,"principal"),
          Light:     this._countOp(lista,"light"),
          Carne:     this._countOp(lista,"carne"),
          Massa:     this._countOp(lista,"massa"),
          Lanche:    this._countOp(lista,"lanche"),
          Total:     lista.length
        }))
      ];
    } else if (tipo === "cc") {
      nomeAba = "Por CC";
      const mapa = {};
      conf.forEach(p => {
        const cc = SP.pick(p, "Centro_Custo") || "Sem CC";
        if (!mapa[cc]) mapa[cc] = [];
        mapa[cc].push(p);
      });
      linhas = Object.entries(mapa).sort((a, b) => b[1].length - a[1].length).map(([cc, lista]) => ({
        Centro_Custo: cc,
        Principal:    this._countOp(lista, "principal"),
        Light:        this._countOp(lista, "light"),
        Carne:        this._countOp(lista, "carne"),
        Massa:        this._countOp(lista, "massa"),
        Lanche:       this._countOp(lista, "lanche"),
        Total:        lista.length
      }));
    } else if (tipo === "ccfunc") {
      nomeAba = "Por CC e Funcionario";
      const mapa = {};
      conf.forEach(p => {
        const cc   = SP.pick(p, "Centro_Custo")             || "Sem CC";
        const nome = SP.pick(p, "Colaborador_nome", "Title") || "Desconhecido";
        const key  = `${cc}||${nome}`;
        if (!mapa[key]) mapa[key] = { cc, nome, lista: [] };
        mapa[key].lista.push(p);
      });
      linhas = Object.values(mapa)
        .sort((a, b) => a.cc.localeCompare(b.cc) || b.lista.length - a.lista.length)
        .map(({ cc, nome, lista }) => ({
          Centro_Custo:     cc,
          Colaborador:      nome,
          Total_Refeicoes:  lista.length,
          Periodo_Ini:      this._periodo.ini,
          Periodo_Fim:      this._periodo.fim
        }));
    }

    if (!linhas.length) { AdminUtils.toast("Nenhum dado para exportar.", "info"); return; }

    const ws = XLSX.utils.json_to_sheet(linhas);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, nomeAba);
    XLSX.writeFile(wb, `relatorio-${tipo}-${this._periodo.ini}-${this._periodo.fim}.xlsx`);
    AdminUtils.toast("Excel exportado.", "success");
  },

  _bindControles() {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundRel) { el.dataset.boundRel = "1"; el.addEventListener(ev, fn); }
    };

    bind("btnBuscarRelatorio",    "click",  async () => {
      const ini = AdminUtils.getVal("relDataIni");
      const fim = AdminUtils.getVal("relDataFim");
      if (!ini || !fim) { AdminUtils.toast("Informe o período.", "error"); return; }
      if (ini > fim)    { AdminUtils.toast("Data início maior que fim.", "error"); return; }
      await this._buscar(ini, fim);
    });

    bind("relTipo", "change", () => this._renderTipo());
    bind("btnExportarRelatorio", "click", () => this._exportar());
  }
};
