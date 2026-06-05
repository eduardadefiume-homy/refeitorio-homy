// admin-colaboradores.js — Colaboradores do Admin Homy

const AdminColaboradores = window.AdminColaboradores = {

  _lista: [],
  _editandoId: null,

  async load() {
    await this._carregar();
    this._bindBotoes();
  },

  async _carregar() {
    const tbody = document.getElementById("colabTable") || document.getElementById("colaboradoresTableBody");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="5" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      this._lista = await SP.getTodosColaboradores();
      this._render();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="5" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
    }
  },

  _render() {
    const tbody = document.getElementById("colabTable") || document.getElementById("colaboradoresTableBody");
    if (!tbody) return;

    const busca = AdminUtils.norm(AdminUtils.getVal("searchColab"));
    const lista = busca
      ? this._lista.filter(c => AdminUtils.norm([SP.pick(c, "Nome", "Title"), SP.pick(c, "Departamento")].join(" ")).includes(busca))
      : this._lista;

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="5" class="empty-cell">Nenhum colaborador encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(c => {
      const id    = AdminUtils.esc(c.id || "");
      const nome  = AdminUtils.esc(SP.pick(c, "Nome", "Title")       || "—");
      const dept  = AdminUtils.esc(SP.pick(c, "Departamento")        || "—");
      const tipo  = AdminUtils.esc(SP.pick(c, "tipo")                || "Colaborador");
      const ativo = SP.isTrue(SP.pick(c, "Ativo"));
      return `<tr>
        <td>${nome}</td>
        <td>${dept}</td>
        <td><span class="badge badge-blue">${tipo}</span></td>
        <td><span class="badge ${ativo ? "badge-green" : "badge-red"}">${ativo ? "Ativo" : "Inativo"}</span></td>
        <td><div class="table-actions">
          <button class="btn-icon" title="Editar"     onclick="AdminColaboradores.abrirEdicao('${id}')">✏️</button>
          <button class="btn-icon danger" title="Desativar" onclick="AdminColaboradores.desativar('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  // ── Modal ────────────────────────────────────────────────────
  abrirNovo() {
    this._editandoId = null;
    this._limparModal();
    const t = document.querySelector("#modalColaborador .modal-title");
    if (t) t.textContent = "Novo colaborador";
    AdminUtils.openModal("modalColaborador");
    setTimeout(() => document.getElementById("colabNome")?.focus(), 80);
  },

  abrirEdicao(id) {
    this._editandoId = id;
    const c = this._lista.find(x => String(x.id) === String(id));
    if (!c) { AdminUtils.toast("Colaborador não encontrado.", "error"); return; }

    AdminUtils.setVal("colabNome",         SP.pick(c, "Nome", "Title")    || "");
    AdminUtils.setVal("colabDepartamento", SP.pick(c, "Departamento")     || "");
    AdminUtils.setVal("colabCentroCusto",  SP.pick(c, "Centro_Custo")     || "");
    AdminUtils.setVal("colabEmail",        SP.pick(c, "Email")            || "");
    AdminUtils.setVal("colabTipo",         SP.pick(c, "tipo")             || "colaborador");

    const t = document.querySelector("#modalColaborador .modal-title");
    if (t) t.textContent = "Editar colaborador";
    AdminUtils.openModal("modalColaborador");
  },

  _limparModal() {
    ["colabNome", "colabDepartamento", "colabCentroCusto", "colabEmail"].forEach(id => AdminUtils.setVal(id, ""));
    AdminUtils.setVal("colabTipo", "colaborador");
  },

  async salvar() {
    const nome        = AdminUtils.getVal("colabNome");
    const departamento = AdminUtils.getVal("colabDepartamento");
    const centroCusto  = AdminUtils.getVal("colabCentroCusto");
    const email        = AdminUtils.getVal("colabEmail");
    const tipo         = AdminUtils.getVal("colabTipo") || "colaborador";

    if (!nome) { AdminUtils.toast("Informe o nome.", "error"); return; }

    try {
      await SP.init();
      if (this._editandoId) {
        await SP.updateColaborador(this._editandoId, { nome, departamento, centroCusto, email, tipo });
        AdminUtils.toast("Colaborador atualizado.", "success");
      } else {
        await SP.createColaborador({ nome, departamento, centroCusto, email, tipo });
        AdminUtils.toast("Colaborador cadastrado.", "success");
      }
      AdminUtils.closeModal("modalColaborador");
      this._editandoId = null;
      await this._carregar();
    } catch (e) {
      AdminUtils.toast("Erro ao salvar: " + e.message, "error");
    }
  },

  async desativar(id) {
    if (!confirm("Desativar este colaborador?")) return;
    try {
      await SP.init();
      await SP.desativarColaborador(id);
      const c = this._lista.find(x => String(x.id) === String(id));
      if (c) c.Ativo = false;
      this._render();
      AdminUtils.toast("Colaborador desativado.", "success");
    } catch (e) {
      AdminUtils.toast("Erro: " + e.message, "error");
    }
  },

  // ── Importação Excel ─────────────────────────────────────────
  async importarExcel(file) {
    if (!file) return;
    if (typeof XLSX === "undefined") { AdminUtils.toast("Biblioteca XLSX não carregou.", "error"); return; }

    try {
      const buf  = await file.arrayBuffer();
      const wb   = XLSX.read(buf, { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });

      const norm = v => AdminUtils.norm(v).replace(/[^a-z0-9]/g, "");
      const get  = (row, ...names) => {
        const keys = Object.keys(row);
        for (const n of names) {
          const k = keys.find(k => norm(k) === norm(n));
          if (k && String(row[k]).trim()) return String(row[k]).trim();
        }
        return "";
      };

      const lista = rows.map(row => ({
        nome:         get(row, "Nome", "Colaborador", "Funcionario", "Nome completo"),
        departamento: get(row, "Departamento", "Setor", "Area"),
        centroCusto:  get(row, "Centro_Custo", "Centro de Custo", "CC"),
        email:        get(row, "Email", "E-mail", "Mail"),
        tipo:         get(row, "tipo", "Tipo") || "colaborador"
      })).filter(c => c.nome);

      if (!lista.length) { AdminUtils.toast("Nenhum colaborador encontrado na planilha.", "error"); return; }
      if (!confirm(`${lista.length} colaboradores encontrados. Importar?`)) return;

      await SP.init();
      let ok = 0, falhas = 0;
      for (const c of lista) {
        try { await SP.createColaborador(c); ok++; }
        catch (e) { console.warn("Falha:", c, e); falhas++; }
      }

      AdminUtils.toast(`Importados: ${ok}${falhas ? ` / ${falhas} falhas` : ""}.`, falhas ? "error" : "success");
      await this._carregar();
    } catch (e) {
      AdminUtils.toast("Erro ao importar: " + e.message, "error");
    }
  },

  // ── Bindings ─────────────────────────────────────────────────
  _bindBotoes() {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundCol) { el.dataset.boundCol = "1"; el.addEventListener(ev, fn); }
    };

    bind("btnNovoColaborador",   "click",  () => this.abrirNovo());
    bind("salvarColaborador",    "click",  () => this.salvar());
    bind("cancelarColaborador",  "click",  () => AdminUtils.closeModal("modalColaborador"));
    bind("searchColab",          "input",  () => this._render());

    bind("btnImportarExcel", "click", () => {
      const inp = document.getElementById("excelInput") || document.getElementById("inputImportarColaboradoresExcel");
      if (inp) { inp.value = ""; inp.click(); }
    });

    const excelInputs = ["excelInput", "inputImportarColaboradoresExcel"];
    excelInputs.forEach(id => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundCol) {
        el.dataset.boundCol = "1";
        el.addEventListener("change", () => this.importarExcel(el.files[0]));
      }
    });
  }
};
