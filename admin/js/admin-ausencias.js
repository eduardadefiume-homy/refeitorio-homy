// admin-ausencias.js — Gestão de Ausências do Admin Homy
// Luana cadastra períodos de ausência (férias, atestado, falta)
// Colaborador aparece travado na cozinha durante o período

const AdminAusencias = window.AdminAusencias = {

  _lista: [],

  async load() {
    this._bindFiltros();
    this._bindBotoes();
    await this._carregar();
  },

  async _carregar() {
    const tbody = document.getElementById("ausenciasTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const todos = await SP.getAusencias(false);

      // Ordena por data início (mais recentes primeiro)
      todos.sort((a, b) => {
        const da = new Date(SP.pick(a, "Data_Inicio") || 0);
        const db = new Date(SP.pick(b, "Data_Inicio") || 0);
        return db - da;
      });

      this._lista = todos;
      this._render();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
    }
  },

  _render() {
    const tbody = document.getElementById("ausenciasTable");
    if (!tbody) return;

    const txt    = AdminUtils.norm(AdminUtils.getVal("fAusenciaTexto") || "");
    const status = AdminUtils.getVal("fAusenciaStatus") || "";
    const hoje   = new Date(); hoje.setHours(0,0,0,0);

    const lista = this._lista.filter(a => {
      const nome   = AdminUtils.norm(SP.pick(a, "Colaborador_nome") || "");
      const motivo = AdminUtils.norm(SP.pick(a, "Motivo")           || "");
      if (txt && !nome.includes(txt) && !motivo.includes(txt)) return false;
      if (status) {
        const s = this._calcStatus(a, hoje);
        if (s !== status) return false;
      }
      return true;
    });

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhuma ausência encontrada.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(a => {
      const id    = AdminUtils.esc(a.id || "");
      const nome  = AdminUtils.esc(SP.pick(a, "Colaborador_nome") || "—");
      const ini   = AdminUtils.esc(this._fmtData(SP.pick(a, "Data_Inicio")) || "—");
      const fim   = AdminUtils.esc(this._fmtData(SP.pick(a, "Data_Fim"))    || "—");
      const motivo = AdminUtils.esc(SP.pick(a, "Motivo")     || "—");
      const obs   = AdminUtils.esc(SP.pick(a, "Observacao")  || "—");
      const s     = this._calcStatus(a, hoje);
      const sLabel = {ativa:"Ativa", futura:"Futura", encerrada:"Encerrada"}[s] || s;
      const sColor = {ativa:"rgba(192,40,28,.7)", futura:"rgba(60,140,255,.6)", encerrada:"rgba(100,100,100,.5)"}[s] || "";

      return `<tr>
        <td style="font-weight:600">${nome}</td>
        <td>${ini}</td>
        <td>${fim}</td>
        <td><span class="badge badge-yellow">${motivo}</span></td>
        <td><span class="badge" style="background:${sColor};color:#fff;border:none">${sLabel}</span></td>
        <td><div class="table-actions">
          <button class="btn-icon danger" title="Excluir" onclick="AdminAusencias.excluir('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  _calcStatus(a, hoje = new Date()) {
    const ini = new Date(SP.pick(a, "Data_Inicio") || 0);
    const fim = new Date(SP.pick(a, "Data_Fim")    || 0);
    ini.setHours(0,0,0,0); fim.setHours(23,59,59,999);
    if (!SP.isTrue(SP.pick(a, "Ativo"))) return "encerrada";
    if (hoje < ini) return "futura";
    if (hoje > fim) return "encerrada";
    return "ativa";
  },

  _fmtData(v) {
    if (!v) return "";
    const d = new Date(v);
    if (isNaN(d)) return v;
    return d.toLocaleDateString("pt-BR", {day:"2-digit", month:"2-digit", year:"numeric"});
  },

  // ── Modal ─────────────────────────────────────────────────────
  async abrirModal() {
    // Popula select de colaboradores
    const sel = document.getElementById("ausenciaColaborador");
    if (sel && sel.options.length <= 1) {
      try {
        await SP.init();
        const colabs = await SP.getColaboradores();
        colabs.sort((a, b) =>
          (SP.pick(a,"Nome","Title")||"").localeCompare(SP.pick(b,"Nome","Title")||"", "pt-BR")
        );
        sel.innerHTML = '<option value="">Selecione o colaborador...</option>' +
          colabs.map(c =>
            `<option value="${AdminUtils.esc(c.id)}" data-nome="${AdminUtils.esc(SP.pick(c,"Nome","Title")||"")}">${AdminUtils.esc(SP.pick(c,"Nome","Title")||"")}</option>`
          ).join("");
      } catch (e) {
        console.warn("[ausencias] carregar colaboradores:", e);
      }
    }

    // Reseta campos
    if (sel) sel.value = "";
    ["ausenciaInicio","ausenciaFim","ausenciaObs"].forEach(id => AdminUtils.setVal(id,""));
    AdminUtils.setVal("ausenciaMotivo","ferias");

    AdminUtils.openModal("modalAusencia");
  },

  async salvar() {
    const sel   = document.getElementById("ausenciaColaborador");
    const colId = sel?.value || "";
    const colNome = sel?.selectedOptions?.[0]?.dataset?.nome || "";
    const inicio  = AdminUtils.getVal("ausenciaInicio");
    const fim     = AdminUtils.getVal("ausenciaFim");
    const motivo  = AdminUtils.getVal("ausenciaMotivo") || "ferias";
    const obs     = AdminUtils.getVal("ausenciaObs");

    if (!colId)   { AdminUtils.toast("Selecione o colaborador.", "error"); return; }
    if (!inicio)  { AdminUtils.toast("Informe a data de início.", "error"); return; }
    if (!fim)     { AdminUtils.toast("Informe a data de fim.", "error"); return; }
    if (new Date(inicio) > new Date(fim)) {
      AdminUtils.toast("Data de início deve ser anterior ao fim.", "error"); return;
    }

    try {
      await SP.init();
      await SP.createAusencia({
        colaboradorId:   colId,
        colaboradorNome: colNome,
        dataInicio:      inicio,
        dataFim:         fim,
        motivo,
        observacao:      obs,
        ativo:           true,
        criadoPor:       SP.getUserName()
      });
      AdminUtils.closeModal("modalAusencia");
      AdminUtils.toast("✅ Ausência cadastrada.", "success");
      await this._carregar();
    } catch (e) {
      AdminUtils.toast("Erro ao salvar: " + e.message, "error");
    }
  },

  async excluir(id) {
    if (!confirm("Excluir esta ausência?")) return;
    try {
      await SP.init();
      await SP.deleteAusencia(id);
      this._lista = this._lista.filter(a => String(a.id) !== String(id));
      this._render();
      AdminUtils.toast("Ausência removida.", "success");
    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + e.message, "error");
    }
  },

  // ── Bindings ──────────────────────────────────────────────────
  _bindFiltros() {
    const bind = (id, ev) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundAus) {
        el.dataset.boundAus = "1";
        el.addEventListener(ev, () => this._render());
      }
    };
    bind("fAusenciaTexto",  "input");
    bind("fAusenciaStatus", "change");

    const btnLimpar = document.getElementById("btnLimparAusencias");
    if (btnLimpar && !btnLimpar.dataset.boundAus) {
      btnLimpar.dataset.boundAus = "1";
      btnLimpar.addEventListener("click", () => {
        AdminUtils.setVal("fAusenciaTexto","");
        AdminUtils.setVal("fAusenciaStatus","");
        this._render();
      });
    }
  },

  _bindBotoes() {
    const bindBtn = (id, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundAusBtn) {
        el.dataset.boundAusBtn = "1";
        el.addEventListener("click", fn);
      }
    };
    bindBtn("btnAdicionarAusencia",  () => this.abrirModal());
    bindBtn("salvarAusencia",        () => this.salvar());
    bindBtn("cancelModalAusencia",   () => AdminUtils.closeModal("modalAusencia"));
    bindBtn("closeModalAusencia",    () => AdminUtils.closeModal("modalAusencia"));
  }
};
