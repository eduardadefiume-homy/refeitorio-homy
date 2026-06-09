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
    if (wrap) wrap.innerHTML = `<div class="alert alert-info">⏳ Carregando pedidos...</div>`;

    try {
      await SP.init();
      const todos = await SP.getItems("Pedidos");

      if (!todos.length) {
        if (wrap) wrap.innerHTML = `<div class="alert alert-warning">⚠️ Nenhum pedido encontrado na lista Pedidos.</div>`;
        this._pedidos = [];
        this._renderCards();
        return;
      }

      // Calcula semanas no período
      const semanasNoPeriodo = new Set();
      try {
        const d = new Date(ini + "T12:00:00");
        const fimDate = new Date(fim + "T12:00:00");
        while (d <= fimDate) {
          semanasNoPeriodo.add(SP.getSemanaId(d));
          d.setDate(d.getDate() + 7); // avança semana a semana (mais rápido)
        }
      } catch(e) {
        console.warn("[relatorios] getSemanaId falhou:", e);
      }

      this._pedidos = todos.filter(p => {
        const dh  = (SP.pick(p, "Data_Hora") || "").slice(0, 10);
        // 1. Data_Hora dentro do período
        if (dh && dh >= ini && dh <= fim) return true;
        // 2. Semana_id que cobre o período
        const sid = SP.pick(p, "Semana_id") || "";
        if (sid && semanasNoPeriodo.has(sid)) return true;
        // 3. Fallback: aceita qualquer pedido cujo Dia+Semana_id seja da semana no período
        // (pedidos automáticos sem Data_Hora correta)
        if (sid) {
          // Verifica se semana_id começa com o ano do período
          const anoIni = ini.slice(0, 4);
          if (sid.startsWith(anoIni)) return true;
        }
        return false;
      });

      console.log(`[relatorios] ${this._pedidos.length} pedidos carregados para ${ini} → ${fim}`);
      this._renderCards();
      this._renderTipo();
    } catch (e) {
      console.error("[relatorios] _buscar:", e);
      if (wrap) wrap.innerHTML = `<div class="alert" style="background:rgba(220,50,50,.1);color:#ff8080">
        ❌ Erro ao carregar: ${AdminUtils.esc(e.message)}<br>
        <small style="opacity:.7">Verifique o console (F12) para detalhes.</small>
      </div>`;
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

    if (!conf.length) {
      wrap.innerHTML = `<div class="alert alert-warning">
        ⚠️ Nenhum pedido confirmado encontrado para o período <strong>${this._periodo.ini}</strong> a <strong>${this._periodo.fim}</strong>.<br>
        <small style="opacity:.7">Total bruto carregado: ${this._pedidos.length} pedidos (incluindo não confirmados).</small>
      </div>`;
      return;
    }

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
  // ── Exportação Excel com padrão visual Homy ──────────────────
  async _exportar() {
    const tipo = AdminUtils.getVal("relTipo") || "dia";
    const conf = this._pedidos.filter(p => this._isConf(p));
    if (!conf.length) { AdminUtils.toast("Nenhum dado para exportar.", "info"); return; }

    if (typeof ExcelJS === "undefined") {
      AdminUtils.toast("Biblioteca ExcelJS não carregou. Verifique conexão.", "error");
      return;
    }

    const wb  = new ExcelJS.Workbook();
    wb.creator  = "Refeitório Homy Química";
    wb.created  = new Date();

    // ── Cores padrão Homy ──────────────────────────────────────
    const COR_AZUL   = "FF0A1E3D";   // azul escuro (header)
    const COR_VERM   = "FFC0281C";   // vermelho Homy (sub-header)
    const COR_BEGE   = "FFF5F5F5";   // linhas pares
    const COR_BRANCO = "FFFFFFFF";

    const estTitulo = {
      font:      { bold: true, color: { argb: COR_BRANCO }, size: 13, name: "Calibri" },
      fill:      { type: "pattern", pattern: "solid", fgColor: { argb: COR_AZUL } },
      alignment: { horizontal: "center", vertical: "middle" }
    };
    const estPeriodo = {
      font:      { bold: true, color: { argb: COR_BRANCO }, size: 11, name: "Calibri" },
      fill:      { type: "pattern", pattern: "solid", fgColor: { argb: COR_VERM } },
      alignment: { horizontal: "center", vertical: "middle" }
    };
    const estCabecalho = {
      font:      { bold: true, color: { argb: COR_BRANCO }, size: 10, name: "Calibri" },
      fill:      { type: "pattern", pattern: "solid", fgColor: { argb: COR_AZUL } },
      alignment: { horizontal: "center", vertical: "middle" },
      border: {
        bottom: { style: "thin", color: { argb: COR_VERM } }
      }
    };
    const estDado = (par) => ({
      font:      { size: 10, name: "Calibri" },
      fill:      { type: "pattern", pattern: "solid", fgColor: { argb: par ? COR_BEGE : COR_BRANCO } },
      alignment: { horizontal: "center", vertical: "middle" },
      border: {
        bottom: { style: "hair", color: { argb: "FFE0E0E0" } }
      }
    });
    const estDadoEsq = (par) => ({
      ...estDado(par),
      alignment: { horizontal: "left", vertical: "middle" }
    });
    const estTotal = {
      font:      { bold: true, size: 10, name: "Calibri", color: { argb: COR_BRANCO } },
      fill:      { type: "pattern", pattern: "solid", fgColor: { argb: COR_AZUL } },
      alignment: { horizontal: "center", vertical: "middle" }
    };

    function aplicarEstilo(cell, est) {
      if (!est) return;
      if (est.font)      cell.font      = est.font;
      if (est.fill)      cell.fill      = est.fill;
      if (est.alignment) cell.alignment = est.alignment;
      if (est.border)    cell.border    = est.border;
    }

    function mesclar(ws, inicio, fim) {
      ws.mergeCells(`${inicio}:${fim}`);
    }

    function fmtData(v) {
      if (!v) return v;
      if (/^\d{4}-\d{2}-\d{2}$/.test(v)) {
        const [y,m,d] = v.split("-");
        return `${d}/${m}/${y}`;
      }
      return v;
    }

    // ── Por Dia ────────────────────────────────────────────────
    if (tipo === "dia") {
      const ws = wb.addWorksheet("Por Dia");
      ws.columns = [
        { width: 18 }, { width: 13 }, { width: 10 },
        { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }
      ];

      // Linha 1 — Título
      ws.addRow(["RELATÓRIO REFEITÓRIO HOMY  QUANTIDADE POR DIA"]);
      mesclar(ws, "A1", "G1");
      ws.getRow(1).height = 24;
      aplicarEstilo(ws.getCell("A1"), estTitulo);

      // Linha 2 — Período
      ws.addRow([`Período: ${fmtData(this._periodo.ini)} a ${fmtData(this._periodo.fim)}`]);
      mesclar(ws, "A2", "G2");
      ws.getRow(2).height = 20;
      aplicarEstilo(ws.getCell("A2"), estPeriodo);

      // Linha 3 — vazia
      ws.addRow([]);

      // Linha 4 — Total
      ws.addRow(["Total de refeições", conf.length]);
      ws.getCell("A4").font = { bold: true, size: 10 };
      ws.getCell("B4").font = { bold: true, size: 10 };

      // Linha 5 — vazia
      ws.addRow([]);

      // Linha 6 — Cabeçalhos
      const cabRow = ws.addRow(["Data", "Principal", "Light", "Carne", "Massa", "Lanche", "Total"]);
      ws.getRow(6).height = 20;
      cabRow.eachCell(cell => aplicarEstilo(cell, estCabecalho));

      // Dados por dia
      const diaParaData = this._getDiaParaData();
      const ORDEM_DIAS  = ["segunda","terca","quarta","quinta","sexta"];
      const mapasDia = {}; const mapaData = {};
      conf.forEach(p => {
        const dn = AdminUtils.norm(SP.pick(p,"Dia")||"");
        const dt = (SP.pick(p,"Data_Hora")||"").slice(0,10);
        if (ORDEM_DIAS.includes(dn)) { if(!mapasDia[dn]) mapasDia[dn]=[]; mapasDia[dn].push(p); }
        else if (dt) { if(!mapaData[dt]) mapaData[dt]=[]; mapaData[dt].push(p); }
      });
      const linhas2 = [];
      ORDEM_DIAS.forEach(dia => {
        if (!mapasDia[dia]) return;
        const data = diaParaData[dia] || dia;
        linhas2.push({ label: data, lista: mapasDia[dia] });
        if (mapaData[data]) delete mapaData[data];
      });
      Object.entries(mapaData).sort(([a],[b])=>a.localeCompare(b)).forEach(([data,lista])=>linhas2.push({label:data,lista}));
      linhas2.sort((a,b)=>a.label.localeCompare(b.label));

      linhas2.forEach(({label, lista}, i) => {
        const par = i % 2 === 0;
        const r = ws.addRow([
          fmtData(label),
          this._countOp(lista,"principal"),
          this._countOp(lista,"light"),
          this._countOp(lista,"carne"),
          this._countOp(lista,"massa"),
          this._countOp(lista,"lanche"),
          lista.length
        ]);
        r.getCell(1).alignment = { horizontal: "left" };
        r.eachCell((cell, col) => aplicarEstilo(cell, col === 1 ? estDadoEsq(par) : estDado(par)));
      });

      // Linha total final
      const totalRow = ws.addRow([
        "TOTAL",
        this._countOp(conf,"principal"),
        this._countOp(conf,"light"),
        this._countOp(conf,"carne"),
        this._countOp(conf,"massa"),
        this._countOp(conf,"lanche"),
        conf.length
      ]);
      totalRow.eachCell(cell => aplicarEstilo(cell, estTotal));

      // Filtro automático na linha 6
      ws.autoFilter = { from: "A6", to: "G6" };

    // ── Por CC ─────────────────────────────────────────────────
    } else if (tipo === "cc") {
      const ws = wb.addWorksheet("Por CC");
      ws.columns = [{width:42},{width:13},{width:10},{width:10},{width:10},{width:10},{width:10}];
      ws.addRow(["RELATÓRIO REFEITÓRIO HOMY  QUANTIDADE POR CENTRO DE CUSTO"]);
      mesclar(ws, "A1", "G1"); ws.getRow(1).height = 24;
      aplicarEstilo(ws.getCell("A1"), estTitulo);
      ws.addRow([`Período: ${fmtData(this._periodo.ini)} a ${fmtData(this._periodo.fim)}`]);
      mesclar(ws, "A2", "G2"); ws.getRow(2).height = 20;
      aplicarEstilo(ws.getCell("A2"), estPeriodo);
      ws.addRow([]);
      const cab = ws.addRow(["Centro de Custo","Principal","Light","Carne","Massa","Lanche","Total"]);
      ws.getRow(4).height = 20;
      cab.eachCell(cell => aplicarEstilo(cell, estCabecalho));
      const mapa = {};
      conf.forEach(p => { const cc = SP.pick(p,"Centro_Custo")||"Sem CC"; if(!mapa[cc])mapa[cc]=[]; mapa[cc].push(p); });
      Object.entries(mapa).sort((a,b)=>b[1].length-a[1].length).forEach(([cc,lista],i) => {
        const par = i%2===0;
        const r = ws.addRow([cc,this._countOp(lista,"principal"),this._countOp(lista,"light"),this._countOp(lista,"carne"),this._countOp(lista,"massa"),this._countOp(lista,"lanche"),lista.length]);
        r.getCell(1).alignment={horizontal:"left"};
        r.eachCell((cell,col)=>aplicarEstilo(cell,col===1?estDadoEsq(par):estDado(par)));
      });
      const tr = ws.addRow(["TOTAL",this._countOp(conf,"principal"),this._countOp(conf,"light"),this._countOp(conf,"carne"),this._countOp(conf,"massa"),this._countOp(conf,"lanche"),conf.length]);
      tr.eachCell(cell=>aplicarEstilo(cell,estTotal));
      ws.autoFilter = { from:"A4", to:"G4" };

    // ── Por CC + Funcionário ───────────────────────────────────
    } else if (tipo === "ccfunc") {
      const ws = wb.addWorksheet("Por CC e Funcionario");
      ws.columns = [{width:42},{width:32},{width:16},{width:14},{width:14}];
      ws.addRow(["RELATÓRIO REFEITÓRIO HOMY  POR CC E FUNCIONÁRIO"]);
      mesclar(ws,"A1","E1"); ws.getRow(1).height=24;
      aplicarEstilo(ws.getCell("A1"),estTitulo);
      ws.addRow([`Período: ${fmtData(this._periodo.ini)} a ${fmtData(this._periodo.fim)}`]);
      mesclar(ws,"A2","E2"); ws.getRow(2).height=20;
      aplicarEstilo(ws.getCell("A2"),estPeriodo);
      ws.addRow([]);
      const cab=ws.addRow(["Centro de Custo","Colaborador","Total Refeições","Período Início","Período Fim"]);
      ws.getRow(4).height=20;
      cab.eachCell(cell=>aplicarEstilo(cell,estCabecalho));
      const mapa={};
      conf.forEach(p=>{const cc=SP.pick(p,"Centro_Custo")||"Sem CC";const nome=SP.pick(p,"Colaborador_nome","Title")||"Desconhecido";const key=`${cc}||${nome}`;if(!mapa[key])mapa[key]={cc,nome,lista:[]};mapa[key].lista.push(p);});
      Object.values(mapa).sort((a,b)=>a.cc.localeCompare(b.cc)||b.lista.length-a.lista.length).forEach(({cc,nome,lista},i)=>{
        const par=i%2===0;
        const r=ws.addRow([cc,nome,lista.length,fmtData(this._periodo.ini),fmtData(this._periodo.fim)]);
        r.getCell(1).alignment={horizontal:"left"}; r.getCell(2).alignment={horizontal:"left"};
        r.eachCell((cell,col)=>aplicarEstilo(cell,col<=2?estDadoEsq(par):estDado(par)));
      });
      ws.autoFilter={from:"A4",to:"E4"};
    }

    // ── Download ───────────────────────────────────────────────
    try {
      const buffer = await wb.xlsx.writeBuffer();
      const blob   = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url    = URL.createObjectURL(blob);
      const a      = document.createElement("a");
      a.href       = url;
      a.download   = `relatorio-${tipo}-${this._periodo.ini}-${this._periodo.fim}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
      AdminUtils.toast("✅ Excel exportado com padrão Homy.", "success");
    } catch (e) {
      AdminUtils.toast("Erro ao gerar Excel: " + e.message, "error");
    }
  },
  },

  _bindFeito: false,

  _bindControles() {
    if (this._bindFeito) return;
    this._bindFeito = true;

    const btnBuscar = document.getElementById("btnBuscarRelatorio");
    const btnExport = document.getElementById("btnExportarRelatorio");
    const selTipo   = document.getElementById("relTipo");

    if (btnBuscar) {
      btnBuscar.addEventListener("click", async () => {
        const ini = (document.getElementById("relDataIni")?.value || "").trim();
        const fim = (document.getElementById("relDataFim")?.value || "").trim();
        if (!ini || !fim) { AdminUtils.toast("Informe o período.", "error"); return; }
        if (ini > fim)    { AdminUtils.toast("Data início maior que fim.", "error"); return; }
        await AdminRelatorios._buscar(ini, fim);
      });
    }

    if (selTipo) {
      selTipo.addEventListener("change", () => AdminRelatorios._renderTipo());
    }

    if (btnExport) {
      btnExport.addEventListener("click", () => AdminRelatorios._exportar());
    }
  }
};
