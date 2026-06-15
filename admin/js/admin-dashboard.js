// admin-dashboard.js — Dashboard do Admin Homy
// Correção: travamento cria principal confirmado e não conta como ausência

const AdminDashboard = window.AdminDashboard = {
  _prazoTimer: null,
  _cache: null,
  _loading: false,
  _syncTimer: null,
  _syncBound: false,

  async load(semanaId) {
    if (this._loading) return;
    this._loading = true;
    try {
      await SP.init();
      this._ensureStatsLayout();
      this._ensurePrazoPainel();
      this._bindSyncTempoReal();

      const base = typeof SP.getDashboardResumo === "function"
        ? await SP.getDashboardResumo(semanaId).catch(e => {
            console.warn("[Dashboard] getDashboardResumo falhou, usando cálculo local.", e);
            return {};
          })
        : {};

      const resumo = await this._montarResumoDashboard(semanaId, base || {});
      this._cache = resumo;

      this._renderCards(resumo);

      await this._safeRender("toggle-status", () => this._renderToggleStatus(resumo));
      await this._safeRender("extras-hoje", () => this._renderExtrasHoje(resumo, semanaId));
      await this._safeRender("semana-table", () => this._renderSemanaTable(resumo));
      await this._safeRender("ausencias", () => this._renderAusencias(resumo, semanaId));
      await this._safeRender("operacao-dia", () => this._renderOperacaoDia(resumo, semanaId));
      await this._safeRender("gerencial", () => this._renderGerencial(resumo));
      await this._safeRender("setores", () => this._renderSetores(resumo));
      await this._safeRender("proximos-dias", () => this._renderProximosDias(resumo));
      await this._safeRender("alertas", () => this._renderAlertas(resumo));
      await this._safeRender("prazo", () => this._carregarPrazo(semanaId, resumo));

      AdminUtils.setTxt("semanaLabel", AdminState.getSemanaLabel());
    } catch (e) {
      console.error("[Dashboard]", e);
      AdminUtils.toast("Erro ao carregar dashboard: " + (e.message || e), "error");
    } finally {
      this._loading = false;
    }
  },

  _bindSyncTempoReal() {
    if (this._syncBound) return;
    this._syncBound = true;

    const recarregar = () => {
      if (document.hidden) return;
      if (window.AdminState?.moduloAtivo && AdminState.moduloAtivo !== "dashboard") return;
      this.load(AdminState.getSemanaId()).catch?.(console.warn);
    };

    this._syncTimer = setInterval(recarregar, 7000);
    window.addEventListener("focus", recarregar);
    document.addEventListener("visibilitychange", () => { if (!document.hidden) recarregar(); });
    window.addEventListener("storage", e => {
      if (e.key === "homy_refeitorio_sync" || e.key === "homy_cozinha_checkin_sync") recarregar();
    });

    try {
      const canal = new BroadcastChannel("homy-refeitorio-sync");
      canal.onmessage = ev => {
        if (["checkin", "pedido", "extra", "dashboard"].includes(ev?.data?.tipo)) recarregar();
      };
      this._syncChannel = canal;
    } catch (_) {}
  },

  async _safeRender(nome, fn) {
    try { return await fn(); }
    catch (e) {
      console.error(`[Dashboard/${nome}]`, e);
      AdminUtils.toast(`Erro no bloco ${nome}: ` + (e.message || e), "error");
      return null;
    }
  },

  _num(v, fallback = 0) {
    if (v === null || v === undefined || v === "") return fallback;
    if (typeof v === "number") return Number.isFinite(v) ? v : fallback;
    let s = String(v).trim();
    if (!s) return fallback;
    if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
    s = s.replace(/[^0-9.-]/g, "");
    const n = Number(s);
    return Number.isFinite(n) ? n : fallback;
  },

  _fmtQtd(v) {
    const n = this._num(v, 0);
    return Number.isInteger(n) ? String(n) : n.toLocaleString("pt-BR", { maximumFractionDigits: 2 });
  },

  _norm(v) {
    return String(v || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim();
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

  _diaHoje(r) {
    return r?.diaHoje || (typeof AdminUtils.DIA_HOJE === "function" ? AdminUtils.DIA_HOJE() : this._diaDaSemana(new Date()));
  },

  _diaDaSemana(date) {
    const dias = ["domingo", "segunda", "terca", "quarta", "quinta", "sexta", "sabado"];
    const d = date instanceof Date ? date : new Date(date);
    const dia = dias[d.getDay()] || "segunda";
    return ["sabado", "domingo"].includes(dia) ? "segunda" : dia;
  },

  _inputDateLocal(date) {
    if (!(date instanceof Date) || isNaN(date)) return "";
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  },

  _inputTimeLocal(date) {
    if (!(date instanceof Date) || isNaN(date)) return "";
    const h = String(date.getHours()).padStart(2, "0");
    const m = String(date.getMinutes()).padStart(2, "0");
    return `${h}:${m}`;
  },

  _ensureStatsLayout() {
    const grid = document.querySelector("#mod-dashboard .stats-grid");
    if (!grid || grid.dataset.dashboardLayoutV2) return;

    grid.dataset.dashboardLayoutV2 = "1";
    grid.innerHTML = `
      <div class="stat-card"><div class="stat-icon">👥</div><div class="stat-value" id="stat-colab">—</div><div class="stat-label">Colaboradores ativos</div></div>
      <div class="stat-card"><div class="stat-icon">✅</div><div class="stat-value" id="stat-confirmados">—</div><div class="stat-label">Colaboradores confirmados</div></div>
      <div class="stat-card"><div class="stat-icon">⏳</div><div class="stat-value" id="stat-pendentes">—</div><div class="stat-label">Colaboradores pendentes</div></div>
      <div class="stat-card"><div class="stat-icon">📅</div><div class="stat-value" id="dashPedidosConfirmadosSemana">—</div><div class="stat-label">Pedidos confirmados da semana</div></div>
      <div class="stat-card"><div class="stat-icon">⌛</div><div class="stat-value" id="dashPedidosPendentesSemana">—</div><div class="stat-label">Pedidos pendentes da semana</div></div>
      <div class="stat-card"><div class="stat-icon">➕</div><div class="stat-value" id="dashExtrasAtivos">—</div><div class="stat-label">Extras ativos</div></div>
      <div class="stat-card"><div class="stat-icon">📦</div><div class="stat-value" id="dashExtrasSemana">—</div><div class="stat-label">Extras da semana</div></div>
      <div class="stat-card"><div class="stat-icon">🚫</div><div class="stat-value" id="dashAusenciasSemana">—</div><div class="stat-label">Ausências da semana</div></div>
      <div class="stat-card"><div class="stat-icon">🛑</div><div class="stat-value" id="dashAusenciasHoje">—</div><div class="stat-label">Ausências hoje</div></div>
      <div class="stat-card"><div class="stat-icon">🍽️</div><div class="stat-value" id="stat-checkin">—</div><div class="stat-label">Check-ins hoje</div></div>
      <div class="stat-card"><div class="stat-icon">🧾</div><div class="stat-value" id="dashExtrasDia">—</div><div class="stat-label">Extras do dia</div></div>
      <div class="stat-card"><div class="stat-icon">📊</div><div class="stat-value" id="dashTotalHoje">—</div><div class="stat-label">Total de pedidos hoje</div></div>`;
  },

  _ensurePrazoPainel() {
    if (document.getElementById("dashPainelPrazo")) return;
    const box = document.querySelector("#mod-dashboard .prazo-box");
    if (!box) return;
    const painel = document.createElement("div");
    painel.id = "dashPainelPrazo";
    painel.style.marginTop = ".8rem";
    box.appendChild(painel);
  },

  _getSemanaDatas(semanaId) {
    if (SP.getWeekDates) return SP.getWeekDates(semanaId).map(d => this._dateISO(d));
    const [year, week] = semanaId.split("-W").map(Number);
    const jan4 = new Date(year, 0, 4);
    const start = new Date(jan4);
    start.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1 + (week - 1) * 7);
    return Array.from({ length: 5 }, (_, i) => {
      const d = new Date(start);
      d.setDate(start.getDate() + i);
      return this._inputDateLocal(d);
    });
  },

  _getSemanaRange(semanaId) {
    const datas = this._getSemanaDatas(semanaId);
    return { ini: datas[0], fim: datas[4], datas };
  },

  _dataPorDia(semanaId, dia) {
    if (SP.getDataRefBySemanaDia) return SP.getDataRefBySemanaDia(semanaId, dia);
    const idx = { segunda: 0, terca: 1, terça: 1, quarta: 2, quinta: 3, sexta: 4 }[this._norm(dia)];
    return this._getSemanaDatas(semanaId)[idx ?? 0];
  },

  _isAtivoColaborador(c) {
    const ativo = this._pick(c, "Ativo", "ativo");
    if (ativo === "") return true;
    return SP.isTrue ? SP.isTrue(ativo) : !!ativo;
  },

  _colabKeyFromColaborador(c) {
    const id = this._pick(c, "id", "ID", "Colaborador_id", "ColaboradorId");
    if (id) return `id:${String(id)}`;
    const nome = this._pick(c, "Nome", "Title", "Colaborador_nome", "Colaborador");
    return `nome:${this._norm(nome)}`;
  },

  _colabNome(c) {
    return this._pick(c, "Nome", "Title", "Colaborador_nome", "Colaborador") || "Sem nome";
  },

  _colabSetor(c) {
    return this._pick(c, "Centro_Custo", "Setor", "Departamento") || "—";
  },

  _colabKeyFromPedido(p) {
    const id = this._pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId");
    if (id) return `id:${String(id)}`;
    const nome = this._pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title");
    return `nome:${this._norm(nome)}`;
  },

  _valorValido(v) {
    const s = String(v ?? "").trim();
    if (!s) return false;
    const n = this._norm(s);
    return !["-", "_", "—", "sem", "sem setor", "undefined", "null"].includes(n);
  },

  _pedidoTemConteudo(p) {
    // Pedido válido precisa ter identidade real e referência de dia/data.
    // Registros quebrados criados por versões antigas vinham como "Pedido",
    // sem colaborador real, sem dia e sem centro de custo. Eles não devem aparecer
    // nem contam no dashboard/por setor.
    const nome = this._pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title");
    const dia = this._pick(p, "Dia");
    const data = this._pick(p, "Data_Hora", "Data", "Data_Referencia");
    const colabId = this._pick(p, "Colaborador_id", "colaboradorId", "ColaboradorId");
    const nomeNorm = this._norm(nome);

    if (!this._valorValido(nome) || ["pedido", "pedidos", "sem nome", "undefined", "null"].includes(nomeNorm)) return false;
    if (!this._valorValido(dia) && !this._valorValido(data)) return false;
    if (nomeNorm === "pedido" && !this._valorValido(colabId)) return false;
    return true;
  },

  _isPedidoProd(p) {
    if (!this._pedidoTemConteudo(p)) return false;
    const s = this._norm(this._pick(p, "Status") || "");
    return ["confirmado", "extra", "aprovado", "travado"].includes(s) ||
           (SP.isTrue && SP.isTrue(this._pick(p, "Confirmado")));
  },

  _isPedidoAusencia(p) {
    const s = this._norm(this._pick(p, "Status") || "");
    return ["cancelado", "afastado", "ferias", "férias", "nao vai almocar", "não vai almoçar", "bloqueado", "atestado", "licenca", "licença"].includes(s);
  },

  _isExtraPedido(p) {
    if (SP.isExtraPedido) return SP.isExtraPedido(p);
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo") || "");
    const nome = this._norm(this._pick(p, "Colaborador_nome", "Title", "Nome") || "");
    return origem.includes("extra") || origem.includes("investigador") || origem.includes("guarda") || nome.includes("refeicao extra");
  },

  _dataPedido(p, semanaId) {
    // A data real da refeição é sempre Semana + Dia quando esses campos existem.
    // Não usar Data_Hora/Created antes do Dia, porque registros alterados hoje
    // podem ser de outro dia da semana e não devem aparecer como ausência/pedido de hoje.
    const semana = this._pick(p, "Semana_id", "Semana") || semanaId;
    const dia = this._pick(p, "Dia", "dia");
    if (semana && dia) return this._dataPorDia(semana, dia);

    return this._dateISO(this._pick(p, "Data_Refeicao", "Data_Referencia", "Data"));
  },

  _ausenciaAtiva(a) {
    const ativo = this._pick(a, "Ativo", "ativo");
    const status = this._norm(this._pick(a, "Status", "status"));
    if (["inativo", "cancelado", "false", "nao", "não", "0"].includes(status)) return false;
    if (ativo === "") return true;
    return SP.isTrue ? SP.isTrue(ativo) : !!ativo;
  },

  _ausenciaInicio(a) {
    return this._dateISO(this._pick(a, "Data_Inicio", "Inicio", "DataInicio", "Data"));
  },

  _ausenciaFim(a) {
    return this._dateISO(this._pick(a, "Data_Fim", "Fim", "DataFim", "Data")) || this._ausenciaInicio(a);
  },

  _ausenciaKey(a) {
    const id = this._pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId");
    if (id) return `id:${String(id)}`;
    const nome = this._pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title");
    return `nome:${this._norm(nome)}`;
  },

  _periodosSobrepoem(iniA, fimA, iniB, fimB) {
    if (!iniA || !fimA || !iniB || !fimB) return false;
    return iniA <= fimB && fimA >= iniB;
  },

  async _buscarPedidos(semanaId) {
    try {
      if (typeof SP.getPedidos === "function") return await SP.getPedidos(semanaId);
      const todos = await SP.getItems("Pedidos");
      return todos.filter(p => this._pick(p, "Semana_id", "Semana") === semanaId);
    } catch (e) {
      console.warn("[Dashboard] Não foi possível carregar pedidos.", e);
      return [];
    }
  },

  async _buscarColaboradores() {
    try {
      const lista = typeof SP.getTodosColaboradores === "function"
        ? await SP.getTodosColaboradores()
        : await SP.getColaboradores();
      return (lista || []).filter(c => this._isAtivoColaborador(c));
    } catch (e) {
      console.warn("[Dashboard] Não foi possível carregar colaboradores.", e);
      return [];
    }
  },

  async _buscarExtras(semanaId) {
    try {
      if (typeof SP.getExtras === "function") return await SP.getExtras(semanaId);
      const todos = await SP.getItems("Extras");
      return todos.filter(e => this._pick(e, "Semana_id", "Semana") === semanaId);
    } catch (e) {
      console.warn("[Dashboard] Não foi possível carregar extras.", e);
      return [];
    }
  },

  async _buscarAusencias() {
    try {
      if (typeof SP.getAusencias === "function") return await SP.getAusencias(false);
    } catch (e) {
      console.warn("[Dashboard] SP.getAusencias falhou. Tentando nomes alternativos.", e);
    }

    const nomes = [
      "Ausencias do Refeitorio",
      "Ausências do Refeitório",
      "Ausencias_Refeitorio",
      "Ausências_Refeitorio",
      "Ausencias",
      "Ausências"
    ];

    for (const nome of nomes) {
      try {
        return await SP.getItems(nome);
      } catch (_) {}
    }
    return [];
  },


  async _buscarCheckinsHoje(semanaId, dia, dataHoje) {
    const diaNorm = this._norm(dia);
    const hojeISO = dataHoje || this._dataPorDia(semanaId, dia);
    let items = [];

    try {
      if (typeof SP.getCheckIn === "function") {
        items = await SP.getCheckIn(semanaId, dia);
      }
    } catch (e) {
      console.warn("[Dashboard] SP.getCheckIn falhou. Tentando getItems(CheckIn).", e);
    }

    if (!Array.isArray(items) || !items.length) {
      try {
        items = await SP.getItems("CheckIn");
      } catch (e) {
        console.warn("[Dashboard] Não foi possível carregar CheckIn.", e);
        items = [];
      }
    }

    const validos = (items || []).filter(c => {
      const retirouValor = this._pick(c, "Retirou", "retirou", "Confirmado", "confirmado");
      const retirou = SP.isTrue ? SP.isTrue(retirouValor) : (retirouValor === true || ["true", "sim", "1", "yes"].includes(this._norm(retirouValor)));
      if (!retirou) return false;

      const semana = this._pick(c, "Semana_id", "Semana", "semana_id");
      const diaItem = this._pick(c, "Dia", "dia");
      const dataRetirada = this._dateISO(this._pick(c, "Data_Hora_Retirada", "Data_Hora", "Data", "Created", "Modified"));

      const semanaOk = !semana || String(semana) === String(semanaId);
      const diaOk = diaItem && this._norm(diaItem) === diaNorm;
      const dataOk = dataRetirada && dataRetirada === hojeISO;
      return semanaOk && (diaOk || dataOk);
    });

    const unico = new Map();
    validos.forEach(c => {
      const cid = this._pick(c, "Colaborador_id", "colaborador_id", "ColaboradorId");
      const nome = this._pick(c, "Colaborador_nome", "Nome", "Title");
      const key = cid ? `id:${cid}` : `nome:${this._norm(nome)}`;
      if (key !== "nome:" && !unico.has(key)) unico.set(key, c);
    });

    return Array.from(unico.values());
  },

  _isExtraAtivoDashboard(e) {
    const status = this._norm(this._pick(e, "Status", "status") || "");
    const ativo = this._pick(e, "Ativo", "ativo");

    if (["cancelado", "cancelada", "excluido", "excluida", "inativo", "inativa", "false", "nao", "não", "0"].includes(status)) return false;
    if (ativo !== "" && ativo !== undefined && ativo !== null) {
      if (SP.isTrue) return SP.isTrue(ativo);
      return !["false", "nao", "não", "0", "inativo", "inativa"].includes(this._norm(ativo));
    }
    return true;
  },

  _extraDiaBate(e, semanaId, dia) {
    const diaNorm = this._norm(dia);
    const diaExtra = this._norm(this._pick(e, "Dia", "dia"));
    if (diaExtra && diaExtra === diaNorm) return true;

    const dataExtra = this._dateISO(this._pick(e, "Data", "Data_Hora", "Data_Referencia"));
    if (!dataExtra) return false;
    return dataExtra === this._dataPorDia(semanaId, dia);
  },

  _nomeExtra(e) {
    const nome = this._pick(e, "Nome", "Title", "Colaborador_nome", "Colaborador");
    const tipo = this._norm(this._pick(e, "tipo", "Tipo", "Origem"));
    if (this._valorValido(nome)) return String(nome).trim();
    if (tipo.includes("guarda")) return "Guarda";
    if (tipo.includes("investigador")) return "Investigador";
    return "Refeição Extra";
  },

  _tipoExtra(e) {
    return this._pick(e, "tipo", "Tipo", "Origem") || "Extra";
  },

  _pratoLabelPorOpcao(opcao) {
    const op = this._norm(opcao);
    if (op === "principal") return "Prato Principal";
    if (op === "light") return "Opção Light";
    if (op === "carne") return "Opção Carne";
    if (op === "massa") return "Opção Massa";
    if (op === "lanche") return "Lanche";
    return "Cardápio do Dia";
  },

  _extraKey(e) {
    const id = this._pick(e, "id", "ID");
    if (id) return `extraid:${id}`;
    return [
      this._norm(this._pick(e, "Semana_id", "Semana")),
      this._norm(this._pick(e, "Dia", "dia")),
      this._norm(this._nomeExtra(e)),
      this._norm(this._tipoExtra(e)),
      this._norm(this._pick(e, "Opcao", "opcao") || "principal")
    ].join("|");
  },

  _pedidoEquivaleExtra(p, e) {
    const extraId = String(this._pick(e, "id", "ID") || "").trim();
    const obsPedido = String(this._pick(p, "Observacao", "Observação", "observacao") || "");
    if (extraId && obsPedido.includes(`ExtraID:${extraId}`)) return true;

    const nomePedido = this._norm(this._pick(p, "Colaborador_nome", "Nome", "Title", "Colaborador"));
    const nomeExtra = this._norm(this._nomeExtra(e));
    const tipoPedido = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
    const tipoExtra = this._norm(this._tipoExtra(e));
    const diaPedido = this._norm(this._pick(p, "Dia", "dia"));
    const diaExtra = this._norm(this._pick(e, "Dia", "dia"));

    if (diaExtra && diaPedido && diaPedido !== diaExtra) return false;
    if (nomePedido && nomeExtra && nomePedido === nomeExtra) return true;
    if (nomePedido && nomeExtra && nomePedido.includes(nomeExtra)) return true;
    if (tipoPedido && tipoExtra && tipoPedido.includes(tipoExtra) && nomePedido && nomePedido.includes(tipoExtra)) return true;
    return false;
  },

  _extraComoPedidoVirtual(e, semanaId, diaFallback) {
    const nome = this._nomeExtra(e);
    const tipo = this._tipoExtra(e);
    const dia = this._pick(e, "Dia", "dia") || diaFallback;
    const opcao = this._pick(e, "Opcao", "opcao") || "principal";
    const status = this._pick(e, "Status", "status") || "Confirmado";
    const centroCusto = this._centroCustoExtra(e);
    const idExtra = this._pick(e, "id", "ID") || this._extraKey(e);
    const prato = this._pick(e, "Nome_Prato", "Prato", "Descricao", "Descrição") || this._pratoLabelPorOpcao(opcao);

    return {
      id: `extra-${idExtra}`,
      ID: `extra-${idExtra}`,
      _virtualExtra: true,
      _extraSource: e,
      Semana_id: this._pick(e, "Semana_id", "Semana") || semanaId,
      Colaborador_id: `extra-${idExtra}`,
      Colaborador_nome: nome,
      Dia: dia,
      Opcao: opcao,
      Nome_Prato: prato,
      Confirmado: true,
      Status: status,
      Centro_Custo: centroCusto,
      Origem: tipo,
      Observacao: this._pick(e, "Observacao", "Observação", "observacao") || ""
    };
  },

  _mesclarPedidosEExtrasDoDia(semanaId, dia, pedidos, extras) {
    const diaNorm = this._norm(dia);
    const pedidosDia = (pedidos || [])
      .filter(p => this._pedidoTemConteudo(p))
      .filter(p => this._norm(this._pick(p, "Dia", "dia")) === diaNorm)
      .filter(p => this._isPedidoProd(p));

    const extrasDia = (extras || [])
      .filter(e => this._isExtraAtivoDashboard(e))
      .filter(e => this._extraDiaBate(e, semanaId, dia));

    const virtuais = [];
    extrasDia.forEach(e => {
      const jaExiste = pedidosDia.some(p => this._pedidoEquivaleExtra(p, e));
      if (!jaExiste) virtuais.push(this._extraComoPedidoVirtual(e, semanaId, dia));
    });

    return [...pedidosDia, ...virtuais].sort((a, b) => {
      const ea = a._virtualExtra ? 1 : 0;
      const eb = b._virtualExtra ? 1 : 0;
      return ea - eb || String(this._pick(a, "Colaborador_nome", "Nome", "Title")).localeCompare(String(this._pick(b, "Colaborador_nome", "Nome", "Title")), "pt-BR");
    });
  },

  _atualizarSetoresPorListaOperacao(lista, colaboradores = []) {
    const mapa = new Map();
    (lista || []).forEach(p => {
      const setor = p._virtualExtra ? this._centroCustoExtra(p._extraSource || p) : this._centroCustoPedidoResolvido(p, colaboradores);
      const cc = this._formatarCC(setor);
      if (cc === "Sem setor") return;
      mapa.set(cc, (mapa.get(cc) || 0) + 1);
    });
    this._renderSetores({ setoresHoje: Array.from(mapa.entries()) });
  },

  async _montarResumoDashboard(semanaId, base) {
    const { ini, fim, datas } = this._getSemanaRange(semanaId);
    const diaHoje = base.diaHoje || this._diaHoje(base);
    const dataHoje = this._dataPorDia(semanaId, diaHoje);

    const [pedidos, colaboradoresAtivos, extras, ausenciasRaw, checkinsHojeLista] = await Promise.all([
      this._buscarPedidos(semanaId),
      this._buscarColaboradores(),
      this._buscarExtras(semanaId),
      this._buscarAusencias(),
      this._buscarCheckinsHoje(semanaId, diaHoje, dataHoje)
    ]);

    const pedidosDaSemana = pedidos.filter(p => {
      if (!this._pedidoTemConteudo(p)) return false;
      const semana = this._pick(p, "Semana_id", "Semana");
      if (semana && semana === semanaId) return true;
      const d = this._dataPedido(p, semanaId);
      return d && d >= ini && d <= fim;
    });

    const pedidosProd = pedidosDaSemana.filter(p => this._isPedidoProd(p) && !this._isExtraPedido(p));
    const pedidosProdComExtras = pedidosDaSemana.filter(p => this._isPedidoProd(p));
    const pedidosHoje = pedidosDaSemana.filter(p => this._norm(this._pick(p, "Dia")) === this._norm(diaHoje) || this._dataPedido(p, semanaId) === dataHoje);
    const pedidosHojeProd = pedidosHoje.filter(p => this._isPedidoProd(p));

    const confirmadosKeys = new Set();
    pedidosProd.forEach(p => {
      const k = this._colabKeyFromPedido(p);
      if (!k.endsWith("nome:")) confirmadosKeys.add(k);
    });

    const ausenciasAtivas = (ausenciasRaw || []).filter(a => this._ausenciaAtiva(a));
    const ausenciasSemana = ausenciasAtivas.filter(a => this._periodosSobrepoem(this._ausenciaInicio(a), this._ausenciaFim(a), ini, fim));
    const ausenciasHojeLista = ausenciasAtivas.filter(a => this._periodosSobrepoem(this._ausenciaInicio(a), this._ausenciaFim(a), dataHoje, dataHoje));
    const ausentesSemanaKeys = new Set(ausenciasSemana.map(a => this._ausenciaKey(a)));

    const pedidosAusenciaHoje = pedidosHoje.filter(p => this._isPedidoAusencia(p));
    const pedidosAusenciaSemana = pedidosDaSemana.filter(p => this._isPedidoAusencia(p));

    const colaboradoresPendentesLista = colaboradoresAtivos
      .filter(c => {
        const k = this._colabKeyFromColaborador(c);
        return !confirmadosKeys.has(k) && !ausentesSemanaKeys.has(k);
      })
      .sort((a, b) => this._colabNome(a).localeCompare(this._colabNome(b), "pt-BR"))
      .map(c => ({
        id: this._pick(c, "id", "ID") || "",
        nome: this._colabNome(c),
        setor: this._colabSetor(c),
        key: this._colabKeyFromColaborador(c)
      }));

    const expectedSemana = colaboradoresAtivos.length * 5;
    const pedidosConfirmadosSemana = pedidosProd.length;
    const pedidosPendentesSemana = this._num(base.pendentesPedidosSemana ?? base.pedidosPendentesSemana ?? base.pendentesColaboradores, NaN);
    const pendentesSemanaFinal = Number.isFinite(pedidosPendentesSemana)
      ? pedidosPendentesSemana
      : Math.max(0, expectedSemana - pedidosConfirmadosSemana);

    const extrasAtivos = extras.filter(e => this._isExtraAtivoDashboard(e));

    const extrasHoje = extrasAtivos.filter(e => this._extraDiaBate(e, semanaId, diaHoje));

    const porDia = {};
    AdminUtils.DIAS.forEach(dia => {
      const listaDia = pedidosDaSemana.filter(p => this._norm(this._pick(p, "Dia")) === this._norm(dia) || this._dataPedido(p, semanaId) === this._dataPorDia(semanaId, dia));
      const prodDia = listaDia.filter(p => this._isPedidoProd(p));
      const totalDia = prodDia.length;
      porDia[dia] = {
        total: totalDia,
        principal: prodDia.filter(p => this._norm(this._pick(p, "Opcao")) === "principal").length,
        light: prodDia.filter(p => this._norm(this._pick(p, "Opcao")) === "light").length,
        carne: prodDia.filter(p => this._norm(this._pick(p, "Opcao")) === "carne").length,
        massa: prodDia.filter(p => this._norm(this._pick(p, "Opcao")) === "massa").length,
        lanche: prodDia.filter(p => this._norm(this._pick(p, "Opcao")) === "lanche").length,
        pendentes: Math.max(0, colaboradoresAtivos.length - new Set(prodDia.filter(p => !this._isExtraPedido(p)).map(p => this._colabKeyFromPedido(p))).size)
      };
    });

    const setoresHojeMap = new Map();
    pedidosHojeProd.forEach(p => {
      const setor = this._centroCustoPedidoResolvido(p, colaboradoresAtivos);
      if (this._formatarCC(setor) === "Sem setor") return;
      setoresHojeMap.set(setor, (setoresHojeMap.get(setor) || 0) + 1);
    });

    // Garante que extras lançados no módulo Extras apareçam no setor correto
    // mesmo quando ainda não existe pedido espelho válido em Pedidos.
    const extraJaExisteNoPedidoHoje = e => pedidosHojeProd.some(p => this._pedidoEquivaleExtra(p, e));

    const extrasHojeSemPedido = extrasHoje.filter(e => !extraJaExisteNoPedidoHoje(e));
    extrasHojeSemPedido.forEach(e => {
      const setor = this._centroCustoExtra(e);
      setoresHojeMap.set(setor, (setoresHojeMap.get(setor) || 0) + 1);
    });

    const ausenciasHojeComPedidos = [
      ...ausenciasHojeLista.map(a => ({
        nome: this._pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title") || "—",
        setor: this._pick(a, "Centro_Custo", "Setor", "Departamento") || "—",
        status: this._pick(a, "Motivo", "Status") || "Ausente",
        origem: "Ausências"
      })),
      ...pedidosAusenciaHoje.map(p => ({
        nome: this._pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "—",
        setor: this._pick(p, "Centro_Custo", "Setor", "Departamento") || "—",
        status: this._pick(p, "Status") || "Ausente",
        origem: "Pedidos"
      }))
    ];

    const uniqueAusHoje = new Map();
    ausenciasHojeComPedidos.forEach(a => {
      const key = `${this._norm(a.nome)}|${this._norm(a.status)}`;
      if (!uniqueAusHoje.has(key)) uniqueAusHoje.set(key, a);
    });

    return {
      ...base,
      semanaId,
      diaHoje,
      dataHoje,
      colaboradoresAtivos: colaboradoresAtivos.length,
      colaboradoresConfirmados: confirmadosKeys.size,
      colaboradoresPendentes: colaboradoresPendentesLista.length,
      pedidosConfirmadosColaboradores: confirmadosKeys.size,
      pendentesColaboradores: colaboradoresPendentesLista.length,
      pedidosConfirmadosSemana,
      pedidosPendentesSemana: pendentesSemanaFinal,
      totalPedidosSemana: pedidosProdComExtras.length,
      totalPedidosHoje: pedidosHojeProd.length + extrasHojeSemPedido.length,
      checkinsHoje: Array.isArray(checkinsHojeLista) ? checkinsHojeLista.length : this._num(base.checkinsHoje),
      _checkinsHojeLista: checkinsHojeLista || [],
      extrasAtivos: extrasAtivos.length || this._num(base.extrasAtivos),
      extrasSemana: extrasAtivos.length,
      extrasDia: extrasHoje.length,
      extrasConfirmados: extrasAtivos.filter(e => ["confirmado", "aprovado", "extra", ""].includes(this._norm(this._pick(e, "Status")))).length,
      extrasPendentes: extrasAtivos.filter(e => this._norm(this._pick(e, "Status")) === "pendente").length,
      ausenciasSemana: ausenciasSemana.length + pedidosAusenciaSemana.length,
      ausenciasHoje: uniqueAusHoje.size,
      principalHoje: pedidosHojeProd.filter(p => this._norm(this._pick(p, "Opcao")) === "principal").length,
      lightHoje: pedidosHojeProd.filter(p => this._norm(this._pick(p, "Opcao")) === "light").length,
      outrasHoje: pedidosHojeProd.filter(p => !["principal", "light"].includes(this._norm(this._pick(p, "Opcao")))).length,
      porDia: { ...(base.porDia || {}), ...porDia },
      setoresHoje: Array.from(setoresHojeMap.entries()),
      _pedidosSemana: pedidosDaSemana,
      _pedidosHoje: pedidosHoje,
      _colaboradoresPendentesLista: colaboradoresPendentesLista,
      _ausenciasHojeLista: Array.from(uniqueAusHoje.values()),
      _ausenciasSemanaLista: ausenciasSemana,
      _extrasSemanaLista: extrasAtivos,
      _extrasHojeLista: extrasHoje
    };
  },

  _renderCards(r) {
    AdminUtils.setTxt("stat-colab", this._num(r.colaboradoresAtivos));
    AdminUtils.setTxt("stat-confirmados", this._num(r.colaboradoresConfirmados));
    AdminUtils.setTxt("stat-pendentes", this._num(r.colaboradoresPendentes));
    AdminUtils.setTxt("dashPedidosConfirmadosSemana", this._num(r.pedidosConfirmadosSemana));
    AdminUtils.setTxt("dashPedidosPendentesSemana", this._num(r.pedidosPendentesSemana));
    AdminUtils.setTxt("dashExtrasAtivos", this._num(r.extrasAtivos));
    AdminUtils.setTxt("dashExtrasSemana", this._num(r.extrasSemana));
    AdminUtils.setTxt("dashAusenciasSemana", this._num(r.ausenciasSemana));
    AdminUtils.setTxt("dashAusenciasHoje", this._num(r.ausenciasHoje));
    AdminUtils.setTxt("stat-checkin", this._num(r.checkinsHoje));
    AdminUtils.setTxt("dashExtrasDia", this._num(r.extrasDia));
    AdminUtils.setTxt("dashTotalHoje", this._num(r.totalPedidosHoje));

    // Compatibilidade com IDs antigos fora do grid.
    AdminUtils.setTxt("dashTotalSemana", this._num(r.totalPedidosSemana));
    AdminUtils.setTxt("dashPrincipalHoje", this._num(r.principalHoje));
    AdminUtils.setTxt("dashLightHoje", this._num(r.lightHoje));
    AdminUtils.setTxt("dashOutrasHoje", this._num(r.outrasHoje));
    AdminUtils.setTxt("dashSetoresHoje", this._normalizarSetores(r.setoresHoje).length);
  },

  async _renderToggleStatus() {
    const liberado = await SP.isCardapioLiberado().catch(() => false);
    const cardapioV = typeof SP.getCardapioVisivel === "function"
      ? await SP.getCardapioVisivel().catch(() => false)
      : SP.isTrue(await SP.getConfig("cardapio_visivel").catch(() => null));

    const tMarcacao = document.getElementById("toggleMarcacao");
    const tCardapio = document.getElementById("toggleCardapio");
    if (tMarcacao) tMarcacao.checked = !!liberado;
    if (tCardapio) tCardapio.checked = !!cardapioV;

    if (tMarcacao && !tMarcacao.dataset.bound) {
      tMarcacao.dataset.bound = "1";
      tMarcacao.addEventListener("change", async function () {
        try {
          await SP.setMarcacaoLiberada(this.checked);
          AdminUtils.toast(this.checked ? "Marcação liberada." : "Marcação bloqueada.", "success");
          AdminUtils.setTxt("dashMarcacaoStatus", this.checked ? "Aberta" : "Fechada");
        } catch (e) {
          AdminUtils.toast("Erro ao alterar marcação: " + (e.message || e), "error");
          this.checked = !this.checked;
        }
      });
    }

    if (tCardapio && !tCardapio.dataset.bound) {
      tCardapio.dataset.bound = "1";
      tCardapio.addEventListener("change", async function () {
        try {
          await SP.setCardapioVisivel(this.checked);
          AdminUtils.toast(this.checked ? "Cardápio visível." : "Cardápio ocultado.", "success");
          AdminUtils.setTxt("dashCardapioStatus", this.checked ? "Liberado" : "Bloqueado");
        } catch (e) {
          AdminUtils.toast("Erro ao alterar cardápio: " + (e.message || e), "error");
          this.checked = !this.checked;
        }
      });
    }

    AdminUtils.setTxt("dashMarcacaoStatus", liberado ? "Aberta" : "Fechada");
    AdminUtils.setTxt("dashCardapioStatus", cardapioV ? "Liberado" : "Bloqueado");
  },

  async _carregarPrazo(semanaId, resumo) {
    const prazo = await SP.getPrazoMarcacao().catch(() => null);

    if (prazo) {
      const dt = new Date(prazo);
      AdminUtils.setVal("prazoData", this._inputDateLocal(dt));
      AdminUtils.setVal("prazoHora", this._inputTimeLocal(dt));
    }

    this._atualizarPainelPrazo(prazo, this._num(resumo.colaboradoresPendentes));

    const btnSalvar = document.getElementById("btnSalvarPrazo");
    if (btnSalvar && !btnSalvar.dataset.bound) {
      btnSalvar.dataset.bound = "1";
      btnSalvar.addEventListener("click", async () => {
        const data = AdminUtils.getVal("prazoData");
        const hora = AdminUtils.getVal("prazoHora") || "18:00";
        if (!data) { AdminUtils.toast("Informe a data limite.", "error"); return; }
        const valor = `${data}T${hora}:00`;
        try {
          await SP.setPrazoMarcacao(valor);
          AdminUtils.toast("✅ Prazo salvo.", "success");
          await AdminDashboard.load(AdminState.getSemanaId());
        } catch (e) {
          AdminUtils.toast("Erro ao salvar prazo: " + (e.message || e), "error");
        }
      });
    }
  },

  _atualizarPainelPrazo(prazoISO, pendentesColaboradores) {
    const el = document.getElementById("dashPainelPrazo");
    if (!el) return;

    const agora = new Date();
    const dt = prazoISO ? new Date(prazoISO) : null;
    const vencido = dt && !isNaN(dt) && agora > dt;
    const qtdPend = this._num(pendentesColaboradores);

    const fmtDt = dt && !isNaN(dt)
      ? dt.toLocaleDateString("pt-BR", { weekday: "long", day: "2-digit", month: "2-digit" }) +
        " às " + dt.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })
      : null;

    let alertClass = "alert-info";
    let icone = "⏱️";
    let statusTxt = "";

    if (!dt || isNaN(dt)) {
      statusTxt = "Nenhum prazo definido para esta semana.";
      alertClass = "alert-warning";
      icone = "⚠️";
    } else if (vencido) {
      alertClass = "alert-red";
      icone = "🔒";
      statusTxt = qtdPend > 0
        ? `Prazo encerrado em ${fmtDt}. ${qtdPend} colaborador(es) ainda não marcaram.`
        : `Prazo encerrado em ${fmtDt}. Todos os colaboradores marcaram ou estão ausentes.`;
    } else {
      const msRestante = dt - agora;
      const horas = Math.floor(msRestante / 3600000);
      const minutos = Math.floor((msRestante % 3600000) / 60000);
      const tempoTxt = horas > 0 ? `${horas}h ${minutos}min restantes` : `${minutos} min restantes`;
      statusTxt = `Prazo: ${fmtDt} (${tempoTxt}). ${qtdPend} colaborador(es) ainda não marcaram.`;
    }

    const btnHtml = qtdPend > 0
      ? `<button id="btnTravarPendentes" class="btn-primary" style="flex-shrink:0;font-size:.82rem;padding:.5rem 1rem;margin-top:.5rem">
           🔒 Travar ${qtdPend} colaborador(es) pendente(s) como Principal
         </button>`
      : "";

    el.innerHTML = `
      <div class="alert ${alertClass}" style="display:flex;flex-direction:column;gap:.5rem">
        <span>${icone} ${AdminUtils.esc(statusTxt)}</span>
        ${btnHtml}
      </div>`;

    const btnTravar = document.getElementById("btnTravarPendentes");
    if (btnTravar && !btnTravar.dataset.bound) {
      btnTravar.dataset.bound = "1";
      btnTravar.addEventListener("click", () => this._confirmarTravamento(AdminState.getSemanaId(), qtdPend));
    }
  },

  _pratoPrincipalPorDia(cardapio, dia) {
    const item = (cardapio || []).find(c =>
      this._norm(this._pick(c, "Dia")) === this._norm(dia) &&
      this._norm(this._pick(c, "Opcao", "Opção")) === "principal"
    );
    return this._pick(item, "Nome_Prato", "Descricao", "Descrição", "Title") || "Principal";
  },

  async _travarColaboradoresPendentesComoPrincipal(semanaId, pendentes) {
    const listaPendentes = Array.isArray(pendentes) ? pendentes : (this._cache?._colaboradoresPendentesLista || []);
    if (!listaPendentes.length) return { colaboradores: 0, pedidosCriados: 0 };

    const [pedidosSemana, cardapio] = await Promise.all([
      this._buscarPedidos(semanaId),
      typeof SP.getCardapio === "function" ? SP.getCardapio(semanaId).catch(() => []) : Promise.resolve([])
    ]);

    const existentes = new Set(pedidosSemana.map(p => `${this._colabKeyFromPedido(p)}|${this._norm(this._pick(p, "Dia"))}`));
    let pedidosCriados = 0;

    for (const colab of listaPendentes) {
      const colaboradorId = String(colab.id || String(colab.key || "").replace(/^id:/, "") || "").trim();
      const colaboradorNome = String(colab.nome || "").trim();
      const centroCusto = colab.setor || "";

      // Proteção contra pedido fantasma: nunca cria pedido sem colaborador real.
      if (!colaboradorId && !colaboradorNome) {
        console.warn("[Dashboard] Colaborador pendente ignorado sem id/nome:", colab);
        continue;
      }

      const key = colab.key || (colaboradorId ? `id:${colaboradorId}` : `nome:${this._norm(colaboradorNome)}`);
      const nomeFinal = colaboradorNome || `Colaborador ${colaboradorId}`;

      for (const dia of AdminUtils.DIAS) {
        if (!dia) continue;
        const chave = `${key}|${this._norm(dia)}`;
        if (existentes.has(chave)) continue;

        const nomePrato = this._pratoPrincipalPorDia(cardapio, dia);
        await SP.savePedido(semanaId, colaboradorId || this._norm(nomeFinal), nomeFinal, dia, "principal", nomePrato, {
          confirmado: true,
          status: "Confirmado",
          centroCusto,
          origem: "Travamento automático",
          observacao: "Marcado automaticamente como Principal — prazo encerrado",
          alteradoPor: SP.getUserName ? SP.getUserName() : "Admin"
        });
        existentes.add(chave);
        pedidosCriados++;
      }
    }

    return { colaboradores: listaPendentes.length, pedidosCriados };
  },

  async _confirmarTravamento(semanaId, qtd) {
    const pendentes = this._cache?._colaboradoresPendentesLista || [];
    const total = this._num(qtd ?? pendentes.length ?? this._cache?.colaboradoresPendentes);
    if (total <= 0) {
      AdminUtils.toast("✅ Todos os colaboradores já marcaram ou estão ausentes.", "success");
      return;
    }

    const confirmar = window.confirm(
      `Confirmar travamento?

` +
      `${total} colaborador(es) ainda não marcaram a refeição.
` +
      `Será criado pedido PRINCIPAL para esses colaboradores nos dias sem pedido desta semana.

` +
      `A observação "Marcado automaticamente — prazo encerrado" será registrada para auditoria.

` +
      `Esta ação não pode ser desfeita.`
    );
    if (!confirmar) return;

    const btn = document.getElementById("btnTravarPendentes");
    if (btn) { btn.disabled = true; btn.textContent = "⏳ Travando..."; }

    try {
      const resultado = await this._travarColaboradoresPendentesComoPrincipal(semanaId, pendentes);
      AdminUtils.toast(
        `✅ Travamento concluído: ${resultado.colaboradores} colaborador(es), ${resultado.pedidosCriados} pedido(s) criados.`,
        "success"
      );
      await SP.setMarcacaoLiberada(false).catch(() => null);
      await AdminDashboard.load(semanaId);
    } catch (e) {
      AdminUtils.toast("Erro ao travar: " + (e.message || e), "error");
      if (btn) { btn.disabled = false; btn.textContent = `🔒 Travar ${total} colaborador(es) pendente(s) como Principal`; }
    }
  },

  async _renderExtrasHoje(r) {
    AdminUtils.setTxt("dashExtrasSolicitados", this._num(r.extrasDia));
    AdminUtils.setTxt("dashExtrasConfirmados", this._num(r.extrasConfirmados));
    AdminUtils.setTxt("dashExtrasPendentes", this._num(r.extrasPendentes));

    const tbody = document.getElementById("dashExtrasTable");
    if (!tbody) return;
    const extras = r._extrasHojeLista || [];
    if (!extras.length) {
      tbody.innerHTML = `<tr><td colspan="4" class="empty-cell">Nenhum extra para hoje.</td></tr>`;
      return;
    }
    tbody.innerHTML = extras.slice(0, 12).map(e => `
      <tr>
        <td>${AdminUtils.esc(this._pick(e, "Nome", "Title") || "Extra")}</td>
        <td>${AdminUtils.esc(this._pick(e, "tipo", "Tipo") || "Extra")}</td>
        <td>${AdminUtils.esc(this._pick(e, "Opcao") || "—")}</td>
        <td>${AdminUtils.badge(this._pick(e, "Status") || "Confirmado")}</td>
      </tr>`).join("");
  },

  _renderSemanaTable(r) {
    const tbody = document.getElementById("dashSemanaTable");
    if (!tbody) return;
    tbody.innerHTML = AdminUtils.DIAS.map(dia => {
      const d = r.porDia?.[dia] || {};
      return `<tr>
        <td>${AdminUtils.DIA_LABEL[dia]}</td>
        <td>${this._num(d.total)}</td>
        <td>${this._num(d.principal)}</td>
        <td>${this._num(d.light)}</td>
        <td>${this._num(d.pendentes)}</td>
      </tr>`;
    }).join("");
  },

  async _renderAusencias(r) {
    AdminUtils.setTxt("dashNaoMarcaram", this._num(r.colaboradoresPendentes));
    AdminUtils.setTxt("dashCancelados", this._num(r.ausenciasHoje));
    AdminUtils.setTxt("dashTravados", this._num(r.travadosHoje ?? r.travadosSemana ?? 0));

    const tbody = document.getElementById("dashAusenciasTable");
    if (!tbody) return;

    const ausentes = r._ausenciasHojeLista || [];
    if (!ausentes.length) {
      tbody.innerHTML = `<tr><td colspan="3" class="empty-cell">Nenhuma ausência hoje.</td></tr>`;
      return;
    }

    tbody.innerHTML = ausentes.slice(0, 14).map(a => `
      <tr>
        <td>${AdminUtils.esc(a.nome || "—")}</td>
        <td>${AdminUtils.esc(this._formatarCC(a.setor || "—"))}</td>
        <td>${AdminUtils.badge(a.status || "Ausente")}</td>
      </tr>`).join("");
  },

  async _renderOperacaoDia(r, semanaId) {
    const diaHoje = r.diaHoje || this._diaHoje(r);
    const sel = document.getElementById("dashOperacaoDia");

    if (sel) {
      if (!sel.value) sel.value = diaHoje;
      if (!sel.dataset.bound) {
        sel.dataset.bound = "1";
        sel.addEventListener("change", () => this._carregarOperacaoTabela(AdminState.getSemanaId(), sel.value));
      }
    }

    const btn = document.getElementById("btnAtualizarDashboard");
    if (btn && !btn.dataset.bound) {
      btn.dataset.bound = "1";
      btn.addEventListener("click", () => AdminDashboard.load(AdminState.getSemanaId()));
    }

    await this._carregarOperacaoTabela(semanaId, sel?.value || diaHoje);
  },

  async _carregarOperacaoTabela(semanaId, dia) {
    const tbody = document.getElementById("dashOperacaoTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    const pedidos = this._cache?._pedidosSemana || await this._buscarPedidos(semanaId);
    const extras = this._cache?._extrasSemanaLista || await this._buscarExtras(semanaId);
    const colaboradores = await this._buscarColaboradores();
    const lista = this._mesclarPedidosEExtrasDoDia(semanaId, dia, pedidos, extras);

    // Quando o usuário muda o dia na Operação do Dia do dashboard,
    // o bloco "Por Setor" passa a refletir o mesmo dia selecionado.
    this._atualizarSetoresPorListaOperacao(lista, colaboradores);

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhum pedido para este dia.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.slice(0, 60).map(p => {
      const id = AdminUtils.esc(p.id || "");
      const nome = AdminUtils.esc(this._pick(p, "Colaborador_nome", "Title", "Nome") || "—");
      const setor = AdminUtils.esc(this._formatarCC(p._virtualExtra ? this._centroCustoExtra(p._extraSource || p) : this._centroCustoPedidoResolvido(p, colaboradores)));
      const opcao = AdminUtils.esc(this._pick(p, "Opcao") || "—");
      const prato = AdminUtils.esc(this._pick(p, "Nome_Prato", "Descricao") || "—");
      const status = this._pick(p, "Status") || "Pendente";
      const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
      const acoes = p._virtualExtra
        ? `<span class="badge badge-yellow" title="Registro vindo da lista Extras">EXTRA</span>`
        : `<div class="table-actions">
          <button class="btn-icon" title="Confirmar" onclick="AdminOperacao.alterarStatus('${id}','Confirmado')">✅</button>
          <button class="btn-icon danger" title="Cancelar" onclick="AdminOperacao.alterarStatus('${id}','Cancelado')">❌</button>
          <button class="btn-icon" title="Não vai almoçar" onclick="AdminOperacao.alterarStatus('${id}','Não vai almoçar')">🚫</button>
        </div>`;
      return `<tr>
        <td${origem.includes("extra") || origem.includes("guarda") || origem.includes("investigador") ? ' style="color:#ffd36d;font-weight:700"' : ""}>${nome}</td>
        <td>${setor}</td>
        <td><span class="badge badge-blue">${opcao}</span></td>
        <td>${prato}</td>
        <td>${AdminUtils.badge(status)}</td>
        <td>${acoes}</td>
      </tr>`;
    }).join("");
  },

  _renderGerencial(r) {
    const el = document.getElementById("dashGerencialList");
    if (!el) return;
    el.innerHTML = [
      ["Consumo da semana", `${this._num(r.totalPedidosSemana)} refeições`],
      ["Pedidos pendentes", `${this._num(r.pedidosPendentesSemana)} refeições`],
      ["Colaboradores pendentes", `${this._num(r.colaboradoresPendentes)} colaboradores`],
      ["Extras da semana", `${this._num(r.extrasSemana)} extras`],
      ["Ausências da semana", `${this._num(r.ausenciasSemana)} registros`]
    ].map(([a, b]) => `
      <div class="dashboard-list-item">
        <div><div class="dashboard-list-main">${AdminUtils.esc(a)}</div><div class="dashboard-list-sub">${AdminUtils.esc(b)}</div></div>
      </div>`).join("");
  },

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

  _formatarCC(valor) {
    if (!valor) return "Sem setor";
    const v = String(valor).trim();
    const nv = this._norm(v);
    if (!v || v === "—" || v === "-" || v === "_" || nv === "sem" || nv === "sem setor" || nv === "undefined" || nv === "null") return "Sem setor";

    // Padrão oficial: CÓDIGO - NOME DO SETOR.
    // Corrige entradas antigas como "ADM GERAL - 120101" para "120101 - ADM GERAL".
    const match = v.match(/(\d{6})/);
    const codigo = match ? match[1] : "";
    if (codigo) {
      const nomeMapa = this._CC_MAPA[codigo];
      if (nomeMapa) return `${codigo} - ${nomeMapa}`;

      const partes = v.split(" - ").map(x => x.trim()).filter(Boolean);
      const nome = partes.find(x => !/^\d{6}$/.test(x));
      return nome ? `${codigo} - ${nome}` : codigo;
    }

    return v;
  },

  _centroCustoPedido(p) {
    const raw = this._pick(p, "Centro_Custo", "CentroCusto", "Setor", "Departamento");
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo") || "");
    const nome = this._norm(this._pick(p, "Colaborador_nome", "Title", "Nome") || "");
    const obs = this._pick(p, "Observacao", "Observação") || "";
    const rawInvalido = !this._valorValido(raw) || this._norm(raw).includes("sem setor");
    const isExtra = this._isExtraPedido(p) || origem.includes("guarda") || origem.includes("extra") ||
                    origem.includes("visitante") || origem.includes("prestador") || origem.includes("motorista") ||
                    origem.includes("investigador") || nome.includes("guarda") || nome.includes("investigador") ||
                    nome.includes("refeicao extra");

    if (rawInvalido && obs) {
      const m = String(obs).match(/(\d{6})(?:\s*-\s*([A-Za-zÀ-ÿ\s/.-]+))?/);
      if (m) return this._formatarCC(m[0]);
    }

    if ((rawInvalido && isExtra) || origem.includes("guarda") || nome.includes("guarda")) return "120602 - PORTARIA";
    return rawInvalido ? "Sem setor" : this._formatarCC(raw);
  },

  _centroCustoPedidoResolvido(p, colaboradores = []) {
    const direto = this._centroCustoPedido(p);
    if (direto && this._formatarCC(direto) !== "Sem setor") return this._formatarCC(direto);

    const idPedido = String(this._pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    const nomePedido = this._norm(this._pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "");
    const colab = (colaboradores || []).find(c => {
      const idColab = String(this._pick(c, "id", "ID", "Colaborador_id", "ColaboradorId") || "").trim();
      const nomeColab = this._norm(this._pick(c, "Nome", "Title", "Colaborador_nome", "Colaborador") || "");
      return (idPedido && idColab && idPedido === idColab) || (nomePedido && nomeColab && nomePedido === nomeColab);
    });

    const setorColab = colab ? this._pick(colab, "Centro_Custo", "Setor", "Departamento") : "";
    return setorColab ? this._formatarCC(setorColab) : "Sem setor";
  },

  _centroCustoExtra(extra) {
    const direto = this._pick(extra, "Centro_Custo", "CentroCusto", "Setor", "Departamento");
    if (direto && this._formatarCC(direto) !== "Sem setor") return this._formatarCC(direto);

    const obs = this._pick(extra, "Observacao", "Observação") || "";
    const m = String(obs).match(/(\d{6})(?:\s*-\s*([A-Za-zÀ-ÿ\s/.-]+))?/);
    if (m) return this._formatarCC(m[0]);

    // Guarda, investigador e extras sem centro de custo caem na Portaria.
    return "120602 - PORTARIA";
  },

  _extrairSetorTotal(item) {
    if (Array.isArray(item)) return { setor: item[0] ?? "Sem setor", total: this._num(item[1]) };
    if (item && typeof item === "object") {
      const entries = Object.entries(item);
      if (entries.length === 1 && typeof entries[0][1] !== "object") return { setor: entries[0][0] || "Sem setor", total: this._num(entries[0][1]) };
      const setor = this._pick(item, "setor", "Setor", "centroCusto", "CentroCusto", "Centro_Custo", "departamento", "Departamento", "nome", "Nome", "label", "Label", "key", "Key");
      const total = this._pick(item, "total", "Total", "quantidade", "Quantidade", "qtd", "Qtd", "count", "Count", "valor", "Valor");
      return { setor: setor || "Sem setor", total: this._num(total) };
    }
    return { setor: "Sem setor", total: 0 };
  },

  _normalizarSetores(raw) {
    let base = [];
    if (Array.isArray(raw)) base = raw;
    else if (raw instanceof Map) base = Array.from(raw.entries());
    else if (raw && typeof raw === "object") base = Object.entries(raw);

    const acumulado = new Map();
    base.forEach(item => {
      const { setor, total } = this._extrairSetorTotal(item);
      const qtd = this._num(total);
      if (qtd <= 0) return;
      const s = this._formatarCC(String(setor || "Sem setor").trim() || "Sem setor");
      if (s === "Sem setor") return;
      acumulado.set(s, (acumulado.get(s) || 0) + qtd);
    });

    return Array.from(acumulado.entries()).map(([setor, total]) => ({ setor, total }))
      .sort((a, b) => b.total - a.total || String(a.setor).localeCompare(String(b.setor)));
  },

  _renderSetores(r) {
    const el = document.getElementById("dashSetoresList");
    if (!el) return;
    const setores = this._normalizarSetores(r.setoresHoje).slice(0, 8);
    el.innerHTML = setores.length
      ? setores.map(({ setor, total }) => `
          <div class="dashboard-list-item">
            <div>
              <div class="dashboard-list-main">${AdminUtils.esc(this._formatarCC(setor))}</div>
              <div class="dashboard-list-sub">${this._fmtQtd(total)} refeições hoje</div>
            </div>
            <span class="badge badge-blue">${this._fmtQtd(total)}</span>
          </div>`).join("")
      : `<div class="dashboard-list-item"><div class="dashboard-list-main">Sem dados hoje</div></div>`;
  },

  _renderProximosDias(r) {
    const el = document.getElementById("dashProximosDias");
    if (!el) return;
    el.innerHTML = AdminUtils.DIAS.map(dia => {
      const d = r.porDia?.[dia] || {};
      return `<div class="dashboard-list-item">
        <div>
          <div class="dashboard-list-main">${AdminUtils.esc(AdminUtils.DIA_LABEL[dia] || dia)}</div>
          <div class="dashboard-list-sub">${this._num(d.total)} conf. · ${this._num(d.pendentes)} pend.</div>
        </div>
      </div>`;
    }).join("");
  },

  _renderAlertas(r) {
    const el = document.getElementById("dashAlertas");
    if (!el) return;

    const pendentes = r._colaboradoresPendentesLista || [];
    const extrasPendentes = this._num(r.extrasPendentes);
    const itens = [];

    if (pendentes.length > 0) {
      const linhas = pendentes.map(c => `
        <tr>
          <td style="color:#ffd45a;font-weight:700">${AdminUtils.esc(c.nome)}</td>
          <td style="color:rgba(255,220,120,.82);text-align:right">${AdminUtils.esc(this._formatarCC(c.setor))}</td>
        </tr>`).join("");

      itens.push(`
        <div class="dashboard-alert-item" style="border:1px solid rgba(255,190,60,.32);background:rgba(255,180,0,.055);padding:.8rem;border-radius:12px">
          <div style="display:flex;align-items:center;justify-content:space-between;gap:.7rem;margin-bottom:.65rem;color:#ffd45a;font-weight:700">
            <span>⚠️ ${pendentes.length} colaborador(es) ainda não marcaram a refeição.</span>
            <span class="badge badge-yellow">Alerta</span>
          </div>
          <div class="table-wrap" style="max-height:360px;overflow-y:auto;border-color:rgba(255,190,60,.24);background:rgba(255,180,0,.035)">
            <table class="table">
              <thead>
                <tr>
                  <th style="color:#ffd45a">Nome</th>
                  <th style="color:#ffd45a;text-align:right">Setor</th>
                </tr>
              </thead>
              <tbody>${linhas}</tbody>
            </table>
          </div>
        </div>`);
    }

    if (extrasPendentes > 0) {
      itens.push(`<div class="dashboard-alert-item">⚠️ ${extrasPendentes} extras aguardando confirmação.</div>`);
    }

    if (!itens.length) {
      itens.push(`<div class="dashboard-alert-item ok">✅ Operação sem alertas críticos no momento.</div>`);
    }

    el.innerHTML = itens.join("");
    AdminUtils.setTxt("dashAlteracoesManuais", this._num(r.alteracoesManuais ?? 0));
  }
};
