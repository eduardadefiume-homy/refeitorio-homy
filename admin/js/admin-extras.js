// ============================================================
// admin-extras.js — Extras / Visitantes do Admin Homy
// v: base-operacional-extras-v11-1-20260702
//
// Regra v11.1:
// - Guarda e Investigadores são fixos operacionais idempotentes.
// - Botão de fixo nunca cria duplicidade: se existe, reutiliza/atualiza.
// - Fixo não é excluído; quando não houver marmita, marcar "Não vai almoçar".
// - Extra eventual pode ser cancelado, não deletado fisicamente.
// ============================================================

const AdminExtras = window.AdminExtras = {
  _lista: [],

  async load(semanaId) {
    this._bindFiltros(semanaId);
    this._bindBotoes(semanaId);
    await this._carregar(semanaId);
  },

  _norm(v) {
    return AdminUtils.norm ? AdminUtils.norm(v) : String(v || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
  },

  _pick(obj, ...keys) {
    if (SP.pick) return SP.pick(obj, ...keys);
    for (const k of keys) {
      const v = obj?.[k];
      if (v !== undefined && v !== null && String(v).trim() !== "") return v;
    }
    return "";
  },

  _esc(v) {
    return AdminUtils.esc ? AdminUtils.esc(v) : String(v ?? "");
  },

  _diaSelecionado() {
    return this._norm(AdminUtils.getVal("fExtraDia")) || this._norm(AdminUtils.getVal("extraDia")) || "segunda";
  },

  _statusExtra(e) {
    const st = this._pick(e, "Status", "status");
    if (st) return String(st);
    return "Confirmado";
  },

  _isNaoVaiAlmocar(e) {
    const st = this._norm(this._statusExtra(e));
    const obs = this._norm(this._pick(e, "Observacao", "Observação", "observacao"));
    return st.includes("nao vai almocar") || st.includes("não vai almoçar") || obs.includes("nao vai almocar") || obs.includes("não vai almoçar");
  },

  _isCancelado(e) {
    const st = this._norm(this._statusExtra(e));
    return ["cancelado", "bloqueado", "inativo", "excluido", "excluído"].includes(st);
  },

  _isFixo(e) {
    const R = window.HomyRefeitorioRegras;
    if (R?.isExtraFixoOperacional) return R.isExtraFixoOperacional(e);
    const nome = this._norm(this._pick(e, "Nome", "Title", "Colaborador_nome"));
    const tipo = this._norm(this._pick(e, "tipo", "Tipo", "Origem"));
    const obs = this._norm(this._pick(e, "Observacao", "Observação", "observacao"));
    return obs.includes("fixooperacional:") || nome === "guarda" || /^investigador\s*[123]$/.test(nome) || tipo.includes("fixo operacional");
  },

  _tipoBadge(tipo, fixo) {
    const t = this._esc(tipo || "—");
    return `<span class="badge ${fixo ? "badge-blue" : "badge-yellow"}">${fixo ? "FIXO" : t}</span>${fixo ? ` <span class="badge badge-yellow">${t}</span>` : ""}`;
  },

  _statusBadge(e) {
    const st = this._statusExtra(e);
    const n = this._norm(st);
    if (n.includes("nao vai almocar") || n.includes("não vai almoçar")) return `<span class="badge badge-red">NÃO VAI ALMOÇAR</span>`;
    if (["cancelado", "bloqueado", "inativo"].includes(n)) return `<span class="badge badge-red">${this._esc(st)}</span>`;
    return `<span class="badge badge-green">${this._esc(st || "Confirmado")}</span>`;
  },

  async _carregar(semanaId) {
    const tbody = document.getElementById("extrasTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const todos = await SP.getExtras(semanaId);

      // Visualmente mantém uma linha por chave fixa; não altera SharePoint ao carregar.
      const vistos = new Set();
      this._lista = (todos || []).filter(e => {
        const key = this._isFixo(e) && window.HomyRefeitorioRegras?.chaveExtraFixoRegistro
          ? window.HomyRefeitorioRegras.chaveExtraFixoRegistro(e, semanaId)
          : "";
        if (key) {
          // Se houver sujeira histórica duplicada, mostra a primeira ativa/mais antiga.
          if (vistos.has(key)) return false;
          vistos.add(key);
        }
        return !this._isCancelado(e);
      });

      const diaOrd = { segunda: 1, terca: 2, quarta: 3, quinta: 4, sexta: 5 };
      this._lista.sort((a, b) => {
        const da = diaOrd[this._norm(this._pick(a, "Dia"))] || 9;
        const db = diaOrd[this._norm(this._pick(b, "Dia"))] || 9;
        if (da !== db) return da - db;
        const fa = this._isFixo(a) ? 0 : 1;
        const fb = this._isFixo(b) ? 0 : 1;
        if (fa !== fb) return fa - fb;
        return this._norm(this._pick(a, "Nome", "Title")).localeCompare(this._norm(this._pick(b, "Nome", "Title")), "pt-BR");
      });

      this._render();
    } catch (e) {
      console.error("[Extras]", e);
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell" style="color:#ff8080">Erro: ${this._esc(e.message || e)}</td></tr>`;
    }
  },

  _render() {
    const tbody = document.getElementById("extrasTable");
    if (!tbody) return;

    const txt  = this._norm(AdminUtils.getVal("fExtraTexto"));
    const dia  = this._norm(AdminUtils.getVal("fExtraDia"));
    const tipo = this._norm(AdminUtils.getVal("fExtraTipo"));

    const lista = this._lista.filter(e => {
      const all = this._norm([
        this._pick(e, "Nome", "Title"), this._pick(e, "Observacao", "Observação"),
        this._pick(e, "tipo", "Tipo"), this._pick(e, "Dia"), this._statusExtra(e)
      ].join(" "));
      return (!txt || all.includes(txt)) &&
             (!dia  || this._norm(this._pick(e, "Dia")) === dia) &&
             (!tipo || this._norm(this._pick(e, "tipo", "Tipo")).includes(tipo));
    });

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhum extra encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(e => {
      const id   = this._esc(this._pick(e, "id", "ID") || "");
      const nome = this._esc(this._pick(e, "Nome", "Title") || "Extra");
      const tipo = this._esc(this._pick(e, "tipo", "Tipo") || "—");
      const dia  = this._esc(this._pick(e, "Dia") || "—");
      const opc  = this._esc(this._pick(e, "Opcao", "Opção") || "principal");
      const obs  = this._esc(this._pick(e, "Observacao", "Observação") || "—");
      const fixo = this._isFixo(e);
      const naoVai = this._isNaoVaiAlmocar(e);

      const acoes = fixo
        ? `<button class="btn-icon ${naoVai ? "success" : "danger"}" title="${naoVai ? "Reativar fixo" : "Marcar como não vai almoçar"}" onclick="${naoVai ? `AdminExtras.reativarFixo('${id}')` : `AdminExtras.marcarNaoVaiAlmocar('${id}')`}">${naoVai ? "↩️" : "🚫"}</button>`
        : `<button class="btn-icon danger" title="Cancelar extra" onclick="AdminExtras.cancelar('${id}')">🗑️</button>`;

      return `<tr>
        <td>${nome}<div style="margin-top:.25rem">${this._statusBadge(e)}</div></td>
        <td>${this._tipoBadge(tipo, fixo)}</td>
        <td>${dia}</td>
        <td><span class="badge badge-blue">${opc}</span></td>
        <td style="font-size:.78rem;color:rgba(143,170,210,.7)">${obs}</td>
        <td><div class="table-actions">${acoes}</div></td>
      </tr>`;
    }).join("");
  },

  // ── Fixos operacionais ─────────────────────────────────────
  async garantirFixosDia() {
    const semanaId = AdminState.getSemanaId();
    const dia = this._diaSelecionado();
    if (!dia) { AdminUtils.toast("Selecione um dia para garantir os fixos.", "error"); return; }
    if (!confirm(`Garantir fixos operacionais de ${dia} (${semanaId})?\n\nIsso cria somente o que estiver faltando e não duplica o que já existe.`)) return;

    try {
      await SP.init();
      const r = await SP.garantirExtrasFixosDia(semanaId, dia, { usuario: SP.getUserName?.() });
      AdminUtils.toast(`Fixos do dia conferidos: ${r.criados || 0} criado(s), ${r.atualizados || 0} atualizado(s), ${r.ignorados || 0} sem alteração.`, "success");
      SP.clearListCache?.("Extras");
      SP.clearListCache?.("Pedidos");
      await this._carregar(semanaId);
    } catch (e) {
      console.error("[Extras] garantirFixosDia", e);
      AdminUtils.toast("Erro ao garantir fixos: " + (e.message || e), "error");
    }
  },

  async garantirFixosSemana() {
    const semanaId = AdminState.getSemanaId();
    if (!confirm(`Garantir fixos operacionais da semana ${semanaId}?\n\nIsso cria somente o que estiver faltando de segunda a sexta e não duplica registros existentes.`)) return;

    try {
      await SP.init();
      const r = await SP.garantirExtrasFixosSemana(semanaId, { usuario: SP.getUserName?.() });
      AdminUtils.toast(`Fixos da semana conferidos: ${r.criados || 0} criado(s), ${r.atualizados || 0} atualizado(s), ${r.ignorados || 0} sem alteração.`, "success");
      SP.clearListCache?.("Extras");
      SP.clearListCache?.("Pedidos");
      await this._carregar(semanaId);
    } catch (e) {
      console.error("[Extras] garantirFixosSemana", e);
      AdminUtils.toast("Erro ao garantir fixos da semana: " + (e.message || e), "error");
    }
  },

  async marcarNaoVaiAlmocar(id) {
    const item = this._lista.find(e => String(this._pick(e, "id", "ID")) === String(id));
    if (!item) return;
    const nome = this._pick(item, "Nome", "Title") || "extra fixo";
    if (!confirm(`Marcar ${nome} como NÃO VAI ALMOÇAR neste dia?\n\nNão será contado na produção, mas o registro fixo será preservado.`)) return;

    try {
      await SP.init();
      await SP.marcarExtraFixoNaoVaiAlmocar(item, { usuario: SP.getUserName?.() });
      AdminUtils.toast("Fixo marcado como não vai almoçar.", "success");
      await this._carregar(AdminState.getSemanaId());
    } catch (e) {
      console.error("[Extras] marcarNaoVaiAlmocar", e);
      AdminUtils.toast("Erro ao marcar não vai almoçar: " + (e.message || e), "error");
    }
  },

  async reativarFixo(id) {
    const item = this._lista.find(e => String(this._pick(e, "id", "ID")) === String(id));
    if (!item) return;
    const nome = this._pick(item, "Nome", "Title") || "extra fixo";
    if (!confirm(`Reativar ${nome} como Principal/Confirmado?`)) return;

    try {
      await SP.init();
      await SP.reativarExtraFixo(item, { usuario: SP.getUserName?.() });
      AdminUtils.toast("Fixo reativado.", "success");
      await this._carregar(AdminState.getSemanaId());
    } catch (e) {
      console.error("[Extras] reativarFixo", e);
      AdminUtils.toast("Erro ao reativar fixo: " + (e.message || e), "error");
    }
  },

  // ── Extra eventual ──────────────────────────────────────────
  abrirModal() {
    const modal = document.getElementById("modalExtra");
    if (!modal) return;
    modal.dataset.predefinido = "";

    ["extraNome", "extraTipo", "extraOpcao", "extraObs"].forEach(id => {
      const el = document.getElementById(id);
      if (el?.closest(".form-group")) el.closest(".form-group").style.display = "";
      if (el) el.value = "";
    });

    AdminUtils.openModal("modalExtra");
  },

  async salvar() {
    const semanaId = AdminState.getSemanaId();
    const dia = AdminUtils.getVal("extraDia") || this._diaSelecionado() || "segunda";
    const nome = AdminUtils.getVal("extraNome");
    const tipo = AdminUtils.getVal("extraTipo") || "visitante";
    const opcao = AdminUtils.getVal("extraOpcao") || "principal";
    const obs = AdminUtils.getVal("extraObs");

    if (!nome) { AdminUtils.toast("Informe o nome do extra.", "error"); return; }

    try {
      await SP.init();
      await SP.addExtraPedido(semanaId, dia, nome, tipo, opcao, obs, SP.getUserName());
      AdminUtils.closeModal("modalExtra");
      AdminUtils.toast("Extra eventual adicionado.", "success");
      await this._carregar(semanaId);
    } catch (e) {
      console.error("[Extras] salvar", e);
      AdminUtils.toast("Erro ao salvar extra: " + (e.message || e), "error");
    }
  },

  async cancelar(id) {
    const item = this._lista.find(e => String(this._pick(e, "id", "ID")) === String(id));
    if (!item) return;

    if (this._isFixo(item)) {
      AdminUtils.toast("Extra fixo não é excluído. Use 'Não vai almoçar'.", "error");
      return;
    }

    if (!confirm("Cancelar este extra eventual?\n\nO registro não será apagado; ficará cancelado para auditoria.")) return;
    try {
      await SP.init();
      await SP.cancelarExtraComPedido(item, "Cancelado pelo módulo Extras.");
      this._lista = this._lista.filter(e => String(this._pick(e, "id", "ID")) !== String(id));
      this._render();
      AdminUtils.toast("Extra cancelado.", "success");
    } catch (e) {
      console.error("[Extras] cancelar", e);
      AdminUtils.toast("Erro ao cancelar: " + (e.message || e), "error");
    }
  },

  // compatibilidade com HTML/cache antigo
  abrirModalFixo(predefinido) {
    if (predefinido === "guarda") return this.garantirFixosDia();
    if (predefinido === "investigador") return this.garantirFixosDia();
    return this.abrirModal();
  },
  excluir(id) { return this.cancelar(id); },

  _bindFiltros() {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundExt) { el.dataset.boundExt = "1"; el.addEventListener(ev, fn); }
    };
    bind("fExtraTexto", "input",  () => this._render());
    bind("fExtraDia",   "change", () => this._render());
    bind("fExtraTipo",  "change", () => this._render());
    bind("btnFiltrarExtras", "click", () => this._render());
    bind("btnLimparExtras",  "click", () => {
      ["fExtraTexto", "fExtraDia", "fExtraTipo"].forEach(id => AdminUtils.setVal(id, ""));
      this._render();
    });
  },

  _bindBotoes() {
    const bind = (id, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundExtBtn) { el.dataset.boundExtBtn = "1"; el.addEventListener("click", fn); }
    };

    bind("btnGarantirFixosDia",    () => this.garantirFixosDia());
    bind("btnGarantirFixosSemana", () => this.garantirFixosSemana());
    bind("btnAdicionarExtra",      () => this.abrirModal());

    // compatibilidade se o HTML antigo ainda estiver em cache
    bind("btnExtraInvestigador",   () => this.garantirFixosDia());
    bind("btnExtraGuarda",         () => this.garantirFixosDia());

    bind("salvarExtra",            () => this.salvar());
    bind("cancelModalExtra",       () => AdminUtils.closeModal("modalExtra"));
    bind("closeModalExtra",        () => AdminUtils.closeModal("modalExtra"));
  }
};
