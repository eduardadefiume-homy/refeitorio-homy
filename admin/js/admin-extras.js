// admin-extras.js — Extras / Visitantes do Admin Homy
// CC automático por tipo: Guarda/Investigador → Portaria, outros → ADM Geral
// Prestador → Luana escolhe CC da lista ou digita manualmente

const AdminExtras = window.AdminExtras = {

  _lista: [],
  _ccsDisponiveis: [],

  // CC automático por tipo de extra (formato NOME - CÓDIGO)
  CC_POR_TIPO: {
    "guarda":       "PORTARIA - 120602",
    "investigador": "PORTARIA - 120602",
    "visitante":    "ADM GERAL - 120101",
    "motorista":    "ADM GERAL - 120101",
    "marmita":      "ADM GERAL - 120101",
    "prestador":    "" // dinâmico — Luana escolhe
  },

  async load(semanaId) {
    await this._carregarCCs();
    this._bindFiltros(semanaId);
    this._bindBotoes(semanaId);
    await this._carregar(semanaId);
  },

  // Mapa oficial: código → nome (plano de contas Homy)
  CC_MAPA: {
    "110101": "DIRETORIA PRESIDENCIAL",
    "110201": "DIRETORIA ADMINISTRATIVA",
    "110202": "DIRETORIA DE PRODUTOS",
    "120101": "ADM GERAL",
    "120102": "CUSTOS",
    "120103": "LEGALIZAÇÃO",
    "120201": "CONTABILIDADE",
    "120202": "FISCAL",
    "120301": "FINANCEIRO",
    "120401": "RECURSOS HUMANOS",
    "120402": "DEPARTAMENTO PESSOAL",
    "120501": "TI",
    "120601": "RECEPÇÃO",
    "120602": "PORTARIA",
    "120603": "ASSEIO E CONSERVAÇÃO",
    "120604": "JARDINAGEM",
    "150101": "SUPRIMENTOS",
    "160101": "CONTROLADORIA E COMPLIANCE",
    "160102": "ADM CONTRATOS",
    "170101": "SGI",
    "180101": "P&D",
    "190101": "PATIO EXTERNO",
    "220101": "ADM VENDAS",
    "220201": "COML INTERNO - SUPORTE",
    "220202": "COML INTERNO - ATIVO",
    "220301": "COML EXTERNO - CLT",
    "220302": "COML EXTERNO - REPRESENTANTE",
    "230101": "SUPORTE TECNICO INDUSTRIAL",
    "230102": "SUPORTE TECNICO OBRAS/INFRA",
    "240101": "MARKETING",
    "250101": "FATURAMENTO",
    "250102": "LOGISTICA",
    "250103": "EXPEDIÇÃO",
    "320101": "PRODUÇÃO",
    "320201": "ENVASE MANUAL",
    "320202": "ENVASE AUTOMATICO",
    "320301": "LABORATORIO E CONTROLE QUALIDADE",
    "360101": "APOIO A PRODUÇÃO",
    "360102": "PCP",
    "360201": "MANUTENÇÃO",
    "360301": "ALMOXARIFADO DE INSUMOS"
  },

  // Normaliza um valor de CC para o formato "NOME - CÓDIGO"
  _normalizarCC(valor) {
    if (!valor) return null;
    const v = String(valor).trim();
    // Se já está no formato "NOME - CÓDIGO", retorna como está
    if (v.includes(" - ")) return v;
    // Se é só código numérico, busca o nome no mapa
    if (/^\d+$/.test(v)) {
      const nome = this.CC_MAPA[v];
      return nome ? `${nome} - ${v}` : null; // descarta códigos desconhecidos
    }
    // É um nome — busca o código correspondente
    const entrada = Object.entries(this.CC_MAPA).find(([, n]) => n === v.toUpperCase());
    return entrada ? `${entrada[1]} - ${entrada[0]}` : v;
  },

  // Carrega CCs distintos da lista de Colaboradores + fixos do mapa oficial
  async _carregarCCs() {
    try {
      await SP.init();
      const colabs = await SP.getTodosColaboradores().catch(() => []);

      // Normaliza CCs vindos do SharePoint
      const ccsDoSP = colabs
        .map(c => this._normalizarCC(SP.pick(c, "Centro_Custo") || ""))
        .filter(Boolean);

      // CCs fixos do mapa oficial no formato "NOME - CÓDIGO"
      const ccsFixos = Object.entries(this.CC_MAPA)
        .map(([cod, nome]) => `${nome} - ${cod}`);

      // Mescla sem duplicar, ordena por nome
      const todos = [...new Set([...ccsFixos, ...ccsDoSP])]
        .sort((a, b) => a.localeCompare(b, "pt-BR"));

      this._ccsDisponiveis = todos;
    } catch (e) {
      // Fallback: só os fixos
      this._ccsDisponiveis = Object.entries(this.CC_MAPA)
        .map(([cod, nome]) => `${nome} - ${cod}`)
        .sort((a, b) => a.localeCompare(b, "pt-BR"));
    }
  },

  async _carregar(semanaId) {
    const tbody = document.getElementById("extrasTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const todos = await SP.getExtras(semanaId);

      // Remove duplicatas de extra automático — mantém só 1 por dia
      const seenAuto = new Set();
      this._lista = todos.filter(e => {
        const norm = v => AdminUtils.norm(v);
        const nome  = norm(SP.pick(e, "Nome", "Title") || "");
        const tipo  = norm(SP.pick(e, "tipo", "Tipo")  || "");
        const dia   = norm(SP.pick(e, "Dia")           || "");
        const isAuto = nome.includes("refeicaoextra") || (nome.includes("extra") && tipo.includes("extra"));
        if (isAuto) {
          const k = `auto-${dia}`;
          if (seenAuto.has(k)) return false;
          seenAuto.add(k);
        }
        return true;
      });

      const diaOrd = { segunda: 1, terca: 2, quarta: 3, quinta: 4, sexta: 5 };
      this._lista.sort((a, b) => {
        const da = diaOrd[AdminUtils.norm(SP.pick(a, "Dia"))] || 9;
        const db = diaOrd[AdminUtils.norm(SP.pick(b, "Dia"))] || 9;
        if (da !== db) return da - db;
        return AdminUtils.norm(SP.pick(a, "tipo")).localeCompare(AdminUtils.norm(SP.pick(b, "tipo")));
      });

      this._render();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
    }
  },

  _render() {
    const tbody = document.getElementById("extrasTable");
    if (!tbody) return;

    const txt  = AdminUtils.norm(AdminUtils.getVal("fExtraTexto"));
    const dia  = AdminUtils.norm(AdminUtils.getVal("fExtraDia"));
    const tipo = AdminUtils.norm(AdminUtils.getVal("fExtraTipo"));

    const lista = this._lista.filter(e => {
      const all = AdminUtils.norm([
        SP.pick(e, "Nome", "Title"), SP.pick(e, "Observacao"),
        SP.pick(e, "tipo", "Tipo"), SP.pick(e, "Dia"), SP.pick(e, "Centro_Custo")
      ].join(" "));
      return (!txt || all.includes(txt)) &&
             (!dia  || AdminUtils.norm(SP.pick(e, "Dia"))            === dia)  &&
             (!tipo || AdminUtils.norm(SP.pick(e, "tipo", "Tipo")).includes(tipo));
    });

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Nenhum extra encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(e => {
      const id   = AdminUtils.esc(e.id || "");
      const nome = AdminUtils.esc(SP.pick(e, "Nome", "Title") || "Extra");
      const tipo = AdminUtils.esc(SP.pick(e, "tipo", "Tipo")  || "—");
      const dia  = AdminUtils.esc(SP.pick(e, "Dia")           || "—");
      const opc  = AdminUtils.esc(SP.pick(e, "Opcao")         || "principal");
      const cc   = AdminUtils.esc(SP.pick(e, "Centro_Custo")  || "—");
      const obs  = AdminUtils.esc(SP.pick(e, "Observacao")    || "—");
      return `<tr>
        <td>${nome}</td>
        <td><span class="badge badge-yellow">${tipo}</span></td>
        <td>${dia}</td>
        <td><span class="badge badge-blue">${opc}</span></td>
        <td><span class="badge badge-blue" style="background:rgba(64,208,144,.15);color:#40d090;border-color:rgba(64,208,144,.25)">${cc}</span></td>
        <td style="font-size:.78rem;color:rgba(143,170,210,.7)">${obs}</td>
        <td><div class="table-actions">
          <button class="btn-icon danger" title="Excluir" onclick="AdminExtras.excluir('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  // ── Modal ────────────────────────────────────────────────────
  abrirModal(predefinido = "") {
    const modal = document.getElementById("modalExtra");
    if (!modal) return;
    modal.dataset.predefinido = predefinido;

    // Reseta campos
    ["extraNome", "extraObs"].forEach(id => AdminUtils.setVal(id, ""));
    AdminUtils.setVal("extraOpcao", "principal");
    AdminUtils.setVal("extraTipo",  "visitante");
    AdminUtils.setVal("extraCC",    "");

    // Popula datalist de CCs
    this._popularDatalistCC();

    // Controla visibilidade do campo CC
    const ccGroup = document.getElementById("extraCCGroup");

    if (predefinido === "investigador") {
      AdminUtils.setVal("extraNome",  "Investigador");
      AdminUtils.setVal("extraTipo",  "investigador");
      AdminUtils.setVal("extraOpcao", "principal");
      AdminUtils.setVal("extraCC",    "Portaria");
      AdminUtils.setVal("extraObs",   "Investigador");
      if (ccGroup) ccGroup.style.display = "none"; // CC automático, não precisa mostrar
    } else if (predefinido === "guarda") {
      AdminUtils.setVal("extraNome",  "Guarda");
      AdminUtils.setVal("extraTipo",  "guarda");
      AdminUtils.setVal("extraOpcao", "principal");
      AdminUtils.setVal("extraCC",    "Portaria");
      AdminUtils.setVal("extraObs",   "Guarda");
      if (ccGroup) ccGroup.style.display = "none";
    } else {
      if (ccGroup) ccGroup.style.display = "";
      // CC padrão para tipo atual
      this._atualizarCCPorTipo();
    }

    AdminUtils.openModal("modalExtra");
  },

  _popularDatalistCC() {
    const dl = document.getElementById("extraCCList");
    if (!dl) return;
    dl.innerHTML = this._ccsDisponiveis
      .map(cc => `<option value="${AdminUtils.esc(cc)}">`)
      .join("");
  },

  _atualizarCCPorTipo() {
    const tipo = AdminUtils.getVal("extraTipo");
    const ccAutomatico = this.CC_POR_TIPO[tipo.toLowerCase()] ?? "";
    const ccGroup = document.getElementById("extraCCGroup");

    if (tipo === "prestador") {
      // Prestador: Luana escolhe CC — mostra campo
      if (ccGroup) ccGroup.style.display = "";
      AdminUtils.setVal("extraCC", "");
    } else if (ccAutomatico) {
      // CC automático: preenche e oculta campo (não precisa escolher)
      AdminUtils.setVal("extraCC", ccAutomatico);
      if (ccGroup) ccGroup.style.display = "none";
    } else {
      // Tipo desconhecido: mostra campo livre
      if (ccGroup) ccGroup.style.display = "";
      AdminUtils.setVal("extraCC", "ADM Geral");
    }
  },

  async salvar() {
    const modal       = document.getElementById("modalExtra");
    const predefinido = modal?.dataset.predefinido || "";
    const semanaId    = AdminState.getSemanaId();
    const dia         = AdminUtils.getVal("extraDia") || "segunda";

    let nome, tipo, opcao, obs, cc;

    if (predefinido === "investigador") {
      const pedidos = await SP.getPedidos(semanaId).catch(() => []);
      const qtd = pedidos.filter(p =>
        AdminUtils.norm(SP.pick(p, "Dia")) === AdminUtils.norm(dia) &&
        AdminUtils.norm(SP.pick(p, "Origem", "tipo")).includes("investigador")
      ).length;
      nome  = `Investigador ${qtd + 1}`;
      tipo  = "investigador";
      opcao = "principal";
      obs   = "Investigador";
      cc    = "Portaria";
    } else if (predefinido === "guarda") {
      nome  = "Guarda";
      tipo  = "guarda";
      opcao = "principal";
      obs   = "Guarda";
      cc    = "Portaria";
    } else {
      nome  = AdminUtils.getVal("extraNome");
      tipo  = AdminUtils.getVal("extraTipo")  || "visitante";
      opcao = AdminUtils.getVal("extraOpcao") || "principal";
      obs   = AdminUtils.getVal("extraObs");
      cc    = AdminUtils.getVal("extraCC")    || "ADM Geral";

      // Se tipo tem CC automático e o campo ficou vazio, aplica
      if (!cc && this.CC_POR_TIPO[tipo.toLowerCase()]) {
        cc = this.CC_POR_TIPO[tipo.toLowerCase()];
      }
      if (!cc) cc = "ADM Geral";
    }

    if (!nome) { AdminUtils.toast("Informe o nome do extra.", "error"); return; }
    if (tipo === "prestador" && !cc) {
      AdminUtils.toast("Informe o centro de custo do prestador.", "error"); return;
    }

    try {
      await SP.init();
      // Grava em Extras com CC
      await SP._addExtraPedidoCC(semanaId, dia, nome, tipo, opcao, obs, cc, SP.getUserName());
      AdminUtils.closeModal("modalExtra");
      AdminUtils.toast("Extra adicionado.", "success");
      await this._carregar(semanaId);
    } catch (e) {
      AdminUtils.toast("Erro ao salvar extra: " + e.message, "error");
    }
  },

  async excluir(id) {
    if (!confirm("Excluir este extra?")) return;
    try {
      await SP.init();
      const semanaId = AdminState.getSemanaId();
      const item     = this._lista.find(e => String(e.id) === String(id));
      if (item) await SP.deleteExtraComPedido(item);
      else await SP.deleteExtra(id);
      this._lista = this._lista.filter(e => String(e.id) !== String(id));
      this._render();
      AdminUtils.toast("Extra excluído.", "success");
    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + e.message, "error");
    }
  },

  // ── Bindings ─────────────────────────────────────────────────
  _bindFiltros(semanaId) {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundExt) { el.dataset.boundExt = "1"; el.addEventListener(ev, fn); }
    };
    bind("fExtraTexto", "input",  () => this._render());
    bind("fExtraDia",   "change", () => this._render());
    bind("fExtraTipo",  "change", () => this._render());
    bind("btnLimparExtras", "click", () => {
      ["fExtraTexto", "fExtraDia", "fExtraTipo"].forEach(id => AdminUtils.setVal(id, ""));
      this._render();
    });
  },

  _bindBotoes(semanaId) {
    const bind = (id, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundExtBtn) { el.dataset.boundExtBtn = "1"; el.addEventListener("click", fn); }
    };
    bind("btnAdicionarExtra",    () => this.abrirModal());
    bind("btnExtraInvestigador", () => this.abrirModal("investigador"));
    bind("btnExtraGuarda",       () => this.abrirModal("guarda"));
    bind("salvarExtra",          () => this.salvar());
    bind("cancelModalExtra",     () => AdminUtils.closeModal("modalExtra"));
    bind("closeModalExtra",      () => AdminUtils.closeModal("modalExtra"));

    // Atualiza CC automaticamente quando tipo muda
    const tipoEl = document.getElementById("extraTipo");
    if (tipoEl && !tipoEl.dataset.boundExtBtn) {
      tipoEl.dataset.boundExtBtn = "1";
      tipoEl.addEventListener("change", () => this._atualizarCCPorTipo());
    }
  }
};
