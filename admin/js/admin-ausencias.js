// ============================================================
// admin-ausencias.js — Férias, afastamentos, ausências e bloqueios
// Correção: lista SharePoint resiliente + renderização completa
// ============================================================

const AdminAusencias = window.AdminAusencias = {
  _lista: [],
  _colaboradores: [],
  _editandoId: null,
  _listName: null,

  // O erro atual veio porque o código buscava só "Ausencias_Refeitorio".
  // Aqui tentamos os nomes mais prováveis sem quebrar a tela.
  LIST_NAMES: [
    "Ausencias_Refeitorio",
    "Ausências_Refeitorio",
    "Ausencias do Refeitorio",
    "Ausências do Refeitório",
    "Ausencias do Refeitório",
    "Ausencias Refeitorio",
    "Ausências Refeitório",
    "Ausencias",
    "Ausências"
  ],

  async load() {
    this._bindControles();
    await this._carregarColaboradores();
    await this._carregar();
  },

  // Compatibilidade com códigos antigos que chamavam carregar/salvar/editar
  async carregar() { await this._carregar(); return this._lista; },
  async salvar(dados) { return this._salvarDireto(dados); },
  async editar(id, dados) { return this._updateAusencia(id, dados); },

  async _descobrirLista() {
    if (this._listName) return this._listName;

    let ultimoErro = null;
    for (const nome of this.LIST_NAMES) {
      try {
        await SP.getListId(nome);
        this._listName = nome;
        console.info(`[Ausências] Lista conectada: ${nome}`);
        return nome;
      } catch (e) {
        ultimoErro = e;
      }
    }

    throw new Error(
      "Lista de ausências não encontrada no SharePoint. " +
      "Crie ou renomeie a lista para 'Ausencias do Refeitorio' ou 'Ausencias_Refeitorio'. " +
      (ultimoErro?.message || "")
    );
  },

  async _getAusencias() {
    const lista = await this._descobrirLista();
    return SP.getItems(lista);
  },

  async _createAusencia(fields) {
    const lista = await this._descobrirLista();
    return SP.createItem(lista, fields);
  },

  async _updateAusencia(id, fields) {
    const lista = await this._descobrirLista();
    return SP.updateItem(lista, id, fields);
  },

  async _deleteAusencia(id) {
    const lista = await this._descobrirLista();
    return SP.deleteItem(lista, id);
  },

  async _carregarColaboradores() {
    try {
      await SP.init();
      if (SP.getTodosColaboradores) this._colaboradores = await SP.getTodosColaboradores();
      else if (SP.getColaboradores) this._colaboradores = await SP.getColaboradores();
      else this._colaboradores = [];
      this._popularSelectColaboradores();
    } catch (e) {
      console.warn("[Ausências] Não foi possível carregar colaboradores:", e);
      this._colaboradores = [];
    }
  },

  async _carregar() {
    const tbody = this._tbody();
    if (tbody) tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      this._lista = await this._getAusencias();
      this._render();
    } catch (e) {
      console.error("[Ausências]", e);
      if (tbody) {
        tbody.innerHTML = `<tr><td colspan="6" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message || e)}</td></tr>`;
      }
      AdminUtils.toast("Erro ao carregar ausências: " + (e.message || e), "error");
    }
  },

  _tbody() {
    return document.getElementById("ausenciasTable") ||
           document.getElementById("ausenciasTableBody") ||
           document.getElementById("ausenciasRefeitorioTable") ||
           document.querySelector("#mod-ausencias tbody");
  },

  _getEl(...ids) {
    for (const id of ids) {
      const el = document.getElementById(id);
      if (el) return el;
    }
    return null;
  },

  _val(...ids) {
    const el = this._getEl(...ids);
    return (el?.value || "").trim();
  },

  _setVal(valor, ...ids) {
    const el = this._getEl(...ids);
    if (el) el.value = valor ?? "";
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

  _dateBR(v) {
    const iso = this._dateISO(v);
    if (!iso) return "—";
    const [a, m, d] = iso.split("-");
    return `${d}/${m}/${a}`;
  },

  _ativo(item) {
    const v = this._pick(item, "Ativo", "ativo", "Status");
    const n = AdminUtils.norm(v);
    if (["inativo", "cancelado", "excluido", "excluído", "false", "nao", "não", "0"].includes(n)) return false;
    if (v === false || v === 0) return false;
    return true;
  },

  _statusLabel(item) {
    return this._ativo(item) ? "Ativo" : "Inativo";
  },

  _render() {
    const tbody = this._tbody();
    if (!tbody) return;

    const busca = AdminUtils.norm(this._val("fAusenciaTexto", "ausenciaBusca", "searchAusencia", "filtroAusenciaTexto"));
    const statusFiltro = AdminUtils.norm(this._val("fAusenciaStatus", "ausenciaStatus", "filtroAusenciaStatus"));

    let lista = [...this._lista];

    if (busca) {
      lista = lista.filter(a => AdminUtils.norm([
        this._pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title"),
        this._pick(a, "Centro_Custo", "Setor", "Departamento"),
        this._pick(a, "Motivo", "motivo"),
        this._pick(a, "Observacao", "Observação", "Obs")
      ].join(" ")).includes(busca));
    }

    if (statusFiltro && statusFiltro !== "todos") {
      lista = lista.filter(a => {
        const ativo = this._ativo(a);
        if (["ativo", "sim", "true"].includes(statusFiltro)) return ativo;
        if (["inativo", "cancelado", "nao", "não", "false"].includes(statusFiltro)) return !ativo;
        return AdminUtils.norm(this._statusLabel(a)) === statusFiltro || AdminUtils.norm(this._pick(a, "Status")) === statusFiltro;
      });
    }

    lista.sort((a, b) => {
      const da = this._dateISO(this._pick(a, "Data_Inicio", "Inicio", "DataInicio"));
      const db = this._dateISO(this._pick(b, "Data_Inicio", "Inicio", "DataInicio"));
      return db.localeCompare(da);
    });

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhuma ausência encontrada.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(a => {
      const id = AdminUtils.esc(a.id || "");
      const nome = AdminUtils.esc(this._pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title") || "—");
      const ini = this._dateBR(this._pick(a, "Data_Inicio", "Inicio", "DataInicio"));
      const fim = this._dateBR(this._pick(a, "Data_Fim", "Fim", "DataFim"));
      const motivo = AdminUtils.esc(this._formatMotivo(this._pick(a, "Motivo", "motivo") || "—"));
      const ativo = this._ativo(a);
      const status = ativo ? "Ativo" : "Inativo";
      return `<tr>
        <td>${nome}</td>
        <td>${ini}</td>
        <td>${fim}</td>
        <td>${motivo}</td>
        <td><span class="badge ${ativo ? "badge-green" : "badge-red"}">${status}</span></td>
        <td><div class="table-actions">
          <button class="btn-icon" title="Editar" onclick="AdminAusencias.abrirEdicao('${id}')">✏️</button>
          ${ativo
            ? `<button class="btn-icon danger" title="Inativar" onclick="AdminAusencias.inativar('${id}')">🚫</button>`
            : `<button class="btn-icon" title="Reativar" onclick="AdminAusencias.reativar('${id}')">↩️</button>`}
        </div></td>
      </tr>`;
    }).join("");
  },

  _formatMotivo(v) {
    const n = AdminUtils.norm(v);
    const map = {
      ferias: "Férias",
      atestado: "Atestado",
      falta: "Falta",
      licenca: "Licença",
      afastamento: "Afastamento",
      nao_vai_almocar: "Não vai almoçar",
      "nao vai almocar": "Não vai almoçar",
      homy_office: "Homy Office",
      banco_horas: "Banco de horas",
      outro: "Outro"
    };
    return map[n] || v;
  },

  _popularSelectColaboradores() {
    const sel = this._getEl("ausenciaColaborador", "ausenciaColaboradorId", "selectAusenciaColaborador");
    if (!sel || sel.dataset.populadoAus) return;

    const opts = this._colaboradores
      .filter(c => SP.isTrue ? SP.isTrue(this._pick(c, "Ativo")) : true)
      .sort((a, b) => String(this._pick(a, "Nome", "Title")).localeCompare(String(this._pick(b, "Nome", "Title"))))
      .map(c => {
        const id = AdminUtils.esc(c.id || this._pick(c, "ID") || "");
        const nome = AdminUtils.esc(this._pick(c, "Nome", "Title") || "Sem nome");
        const cc = AdminUtils.esc(this._pick(c, "Centro_Custo", "Setor", "Departamento") || "");
        return `<option value="${id}" data-nome="${nome}" data-cc="${cc}">${nome}${cc ? " — " + cc : ""}</option>`;
      }).join("");

    sel.innerHTML = `<option value="">Selecione...</option>${opts}`;
    sel.dataset.populadoAus = "1";
  },

  abrirNovo() {
    this._editandoId = null;
    this._setVal("", "ausenciaId", "editAusenciaId");
    this._setVal("", "ausenciaColaborador", "ausenciaColaboradorId", "selectAusenciaColaborador");
    this._setVal("", "ausenciaDataInicio", "ausenciaInicio");
    this._setVal("", "ausenciaDataFim", "ausenciaFim");
    this._setVal("nao_vai_almocar", "ausenciaMotivo");
    this._setVal("", "ausenciaObs", "ausenciaObservacao");

    const t = document.querySelector("#modalAusencia .modal-title, #modalAusenciaRefeitorio .modal-title");
    if (t) t.textContent = "Nova ausência";

    if (document.getElementById("modalAusencia")) AdminUtils.openModal("modalAusencia");
    else if (document.getElementById("modalAusenciaRefeitorio")) AdminUtils.openModal("modalAusenciaRefeitorio");
    else AdminUtils.toast("Modal de ausência não encontrado no HTML.", "error");
  },

  abrirEdicao(id) {
    const item = this._lista.find(a => String(a.id) === String(id));
    if (!item) { AdminUtils.toast("Ausência não encontrada.", "error"); return; }

    this._editandoId = id;
    this._setVal(id, "ausenciaId", "editAusenciaId");
    this._setVal(this._pick(item, "Colaborador_id", "ColaboradorId") || "", "ausenciaColaborador", "ausenciaColaboradorId", "selectAusenciaColaborador");
    this._setVal(this._dateISO(this._pick(item, "Data_Inicio", "Inicio", "DataInicio")), "ausenciaDataInicio", "ausenciaInicio");
    this._setVal(this._dateISO(this._pick(item, "Data_Fim", "Fim", "DataFim")), "ausenciaDataFim", "ausenciaFim");
    this._setVal(this._pick(item, "Motivo", "motivo") || "nao_vai_almocar", "ausenciaMotivo");
    this._setVal(this._pick(item, "Observacao", "Observação", "Obs") || "", "ausenciaObs", "ausenciaObservacao");

    const t = document.querySelector("#modalAusencia .modal-title, #modalAusenciaRefeitorio .modal-title");
    if (t) t.textContent = "Editar ausência";

    if (document.getElementById("modalAusencia")) AdminUtils.openModal("modalAusencia");
    else if (document.getElementById("modalAusenciaRefeitorio")) AdminUtils.openModal("modalAusenciaRefeitorio");
  },

  async salvarModal() {
    const colaboradorId = this._val("ausenciaColaborador", "ausenciaColaboradorId", "selectAusenciaColaborador");
    const sel = this._getEl("ausenciaColaborador", "ausenciaColaboradorId", "selectAusenciaColaborador");
    const opt = sel?.selectedOptions?.[0];
    const colab = this._colaboradores.find(c => String(c.id) === String(colaboradorId));

    const colaboradorNome = opt?.dataset?.nome || this._pick(colab, "Nome", "Title") || this._val("ausenciaColaboradorNome", "ausenciaNome");
    const centroCusto = opt?.dataset?.cc || this._pick(colab, "Centro_Custo", "Setor", "Departamento") || this._val("ausenciaCentroCusto", "ausenciaSetor");
    const dataInicio = this._val("ausenciaDataInicio", "ausenciaInicio");
    const dataFim = this._val("ausenciaDataFim", "ausenciaFim");
    const motivo = this._val("ausenciaMotivo") || "nao_vai_almocar";
    const obs = this._val("ausenciaObs", "ausenciaObservacao");

    if (!colaboradorNome && !colaboradorId) { AdminUtils.toast("Informe o colaborador.", "error"); return; }
    if (!dataInicio || !dataFim) { AdminUtils.toast("Informe início e fim.", "error"); return; }
    if (dataInicio > dataFim) { AdminUtils.toast("Data início maior que data fim.", "error"); return; }

    const fields = {
      Title: `${colaboradorNome || colaboradorId} - ${this._formatMotivo(motivo)}`,
      Colaborador_id: String(colaboradorId || ""),
      Colaborador_nome: colaboradorNome || "",
      Centro_Custo: centroCusto || "",
      Data_Inicio: dataInicio,
      Data_Fim: dataFim,
      Motivo: motivo,
      Observacao: obs || "",
      Ativo: true,
      Criado_Por: SP.getUserName ? SP.getUserName() : "Admin"
    };

    try {
      await SP.init();
      if (this._editandoId) {
        await this._updateAusencia(this._editandoId, fields);
        AdminUtils.toast("Ausência atualizada.", "success");
      } else {
        await this._createAusencia(fields);
        AdminUtils.toast("Ausência cadastrada.", "success");
      }

      if (document.getElementById("modalAusencia")) AdminUtils.closeModal("modalAusencia");
      if (document.getElementById("modalAusenciaRefeitorio")) AdminUtils.closeModal("modalAusenciaRefeitorio");
      this._editandoId = null;
      await this._carregar();
    } catch (e) {
      AdminUtils.toast("Erro ao salvar ausência: " + (e.message || e), "error");
    }
  },

  async _salvarDireto(dados) {
    const fields = {
      Title: dados.Title || `${dados.colaboradorNome || dados.Colaborador_nome || "Ausência"} - ${dados.motivo || dados.Motivo || ""}`,
      Colaborador_id: String(dados.colaboradorId || dados.Colaborador_id || ""),
      Colaborador_nome: dados.colaboradorNome || dados.Colaborador_nome || dados.Nome || "",
      Centro_Custo: dados.centroCusto || dados.Centro_Custo || "",
      Data_Inicio: dados.dataInicio || dados.Data_Inicio,
      Data_Fim: dados.dataFim || dados.Data_Fim,
      Motivo: dados.motivo || dados.Motivo || "nao_vai_almocar",
      Observacao: dados.observacao || dados.Observacao || "",
      Ativo: dados.ativo ?? dados.Ativo ?? true,
      Criado_Por: dados.criadoPor || dados.Criado_Por || (SP.getUserName ? SP.getUserName() : "Admin")
    };
    return this._createAusencia(fields);
  },

  async inativar(id) {
    if (!confirm("Inativar esta ausência?")) return;
    try {
      await this._updateAusencia(id, { Ativo: false, Status: "Inativo" });
      AdminUtils.toast("Ausência inativada.", "success");
      await this._carregar();
    } catch (e) {
      AdminUtils.toast("Erro ao inativar: " + (e.message || e), "error");
    }
  },

  async reativar(id) {
    try {
      await this._updateAusencia(id, { Ativo: true, Status: "Ativo" });
      AdminUtils.toast("Ausência reativada.", "success");
      await this._carregar();
    } catch (e) {
      AdminUtils.toast("Erro ao reativar: " + (e.message || e), "error");
    }
  },

  async excluir(id) {
    if (!confirm("Excluir definitivamente esta ausência?")) return;
    try {
      await this._deleteAusencia(id);
      AdminUtils.toast("Ausência excluída.", "success");
      await this._carregar();
    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + (e.message || e), "error");
    }
  },

  _bindControles() {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundAus) {
        el.dataset.boundAus = "1";
        el.addEventListener(ev, fn);
      }
    };

    ["fAusenciaTexto", "ausenciaBusca", "searchAusencia", "filtroAusenciaTexto"].forEach(id => {
      bind(id, "input", () => this._render());
    });
    ["fAusenciaStatus", "ausenciaStatus", "filtroAusenciaStatus"].forEach(id => {
      bind(id, "change", () => this._render());
    });

    ["btnLimparAusencias", "limparAusencias", "btnLimparFiltroAusencias"].forEach(id => {
      bind(id, "click", () => {
        this._setVal("", "fAusenciaTexto", "ausenciaBusca", "searchAusencia", "filtroAusenciaTexto");
        this._setVal("", "fAusenciaStatus", "ausenciaStatus", "filtroAusenciaStatus");
        this._render();
      });
    });

    ["btnNovaAusencia", "btnNovaAusenciaRefeitorio", "btnAdicionarAusencia"].forEach(id => {
      bind(id, "click", () => this.abrirNovo());
    });

    ["salvarAusencia", "btnSalvarAusencia", "salvarAusenciaRefeitorio"].forEach(id => {
      bind(id, "click", () => this.salvarModal());
    });

    ["cancelarAusencia", "btnCancelarAusencia"].forEach(id => {
      bind(id, "click", () => {
        if (document.getElementById("modalAusencia")) AdminUtils.closeModal("modalAusencia");
        if (document.getElementById("modalAusenciaRefeitorio")) AdminUtils.closeModal("modalAusenciaRefeitorio");
      });
    });
  }
};
