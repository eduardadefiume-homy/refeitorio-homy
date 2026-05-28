// ============================================================
// admin-colaboradores.js — CRUD de colaboradores
// Prioridade 1: Centro_Custo + Editar + Excluir/Desativar
// ============================================================

window.AdminColaboradores = {
  lista: [],
  editandoId: null,

  async carregar() {
    if (!window.SP) {
      console.error("SP não encontrado. Verifique se sharepoint.js foi carregado antes dos módulos admin.");
      return;
    }

    const getFn = SP.getTodosColaboradores ? "getTodosColaboradores" : "getColaboradores";
    this.lista = await SP[getFn]();
    this.renderizarTabela();
  },

  renderizarTabela() {
    const tbody =
      document.getElementById("colaboradoresTableBody") ||
      document.getElementById("tbodyColaboradores") ||
      document.querySelector("[data-colaboradores-tbody]");

    if (!tbody) {
      console.warn("Tabela de colaboradores não encontrada no HTML atual.");
      return;
    }

    const busca = AdminUtils.normalizarTexto(
      document.getElementById("buscaColaborador")?.value || ""
    );

    const filtrados = this.lista.filter(c => {
      const nome = AdminUtils.normalizarTexto(c.Nome || c.Title || "");
      const departamento = AdminUtils.normalizarTexto(c.Departamento || "");
      const centro = AdminUtils.normalizarTexto(c.Centro_Custo || "");
      return !busca || nome.includes(busca) || departamento.includes(busca) || centro.includes(busca);
    });

    if (!filtrados.length) {
      tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;opacity:.55;padding:2rem;">Nenhum colaborador encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = filtrados.map(c => {
      const ativo = SP.isTrue ? SP.isTrue(c.Ativo) : String(c.Ativo).toLowerCase() !== "false";
      return `
        <tr>
          <td>${c.Nome || c.Title || ""}</td>
          <td>${c.Departamento || ""}</td>
          <td>${c.Centro_Custo || "-"}</td>
          <td><span class="badge badge-blue">${c.tipo || "Colaborador"}</span></td>
          <td><span class="badge ${ativo ? "badge-green" : "badge-red"}">${ativo ? "ATIVO" : "INATIVO"}</span></td>
          <td>
            <div class="table-actions">
              <button class="btn-icon" title="Editar" onclick="AdminColaboradores.abrirEditar('${c.id}')">✏️</button>
              <button class="btn-icon danger" title="Desativar" onclick="AdminColaboradores.desativar('${c.id}')">🗑️</button>
            </div>
          </td>
        </tr>
      `;
    }).join("");
  },

  abrirNovo() {
    this.editandoId = null;
    this.preencherModal({});
    AdminUtils.abrirModal("modalColaborador");
  },

  abrirEditar(id) {
    const c = this.lista.find(x => String(x.id) === String(id));
    if (!c) {
      AdminUtils.toast("Colaborador não encontrado para edição.", "warning");
      return;
    }

    this.editandoId = id;
    this.preencherModal(c);
    AdminUtils.abrirModal("modalColaborador");
  },

  preencherModal(c) {
    const set = (id, value) => {
      const el = document.getElementById(id);
      if (el) el.value = value || "";
    };

    set("colabNome", c.Nome || c.Title || "");
    set("colabDepartamento", c.Departamento || "");
    set("colabEmail", c.Email || "");
    set("colabTipo", c.tipo || "Colaborador");
    set("colabCentroCusto", c.Centro_Custo || "");

    const titulo = document.getElementById("modalColaboradorTitulo");
    if (titulo) titulo.textContent = this.editandoId ? "EDITAR COLABORADOR" : "NOVO COLABORADOR";
  },

  obterDadosFormulario() {
    const val = id => document.getElementById(id)?.value?.trim() || "";

    return {
      nome: val("colabNome"),
      departamento: val("colabDepartamento"),
      email: val("colabEmail"),
      tipo: val("colabTipo") || "Colaborador",
      centroCusto: val("colabCentroCusto"),
      Centro_Custo: val("colabCentroCusto")
    };
  },

  async salvar() {
    const dados = this.obterDadosFormulario();

    if (!dados.nome) {
      AdminUtils.toast("Informe o nome do colaborador.", "warning");
      return;
    }

    if (!dados.centroCusto) {
      AdminUtils.toast("Informe o Centro de Custo do colaborador.", "warning");
      return;
    }

    if (this.editandoId) {
      await SP.updateColaborador(this.editandoId, dados);
      AdminUtils.toast("Colaborador atualizado com sucesso.", "success");
    } else {
      await SP.createColaborador(dados);
      AdminUtils.toast("Colaborador cadastrado com sucesso.", "success");
    }

    AdminUtils.fecharModal("modalColaborador");
    await this.carregar();
  },

  async desativar(id) {
    const ok = confirm("Deseja desativar este colaborador? Ele não será removido do histórico.");
    if (!ok) return;

    await SP.desativarColaborador(id);
    AdminUtils.toast("Colaborador desativado.", "success");
    await this.carregar();
  }
};

// Compatibilidade com onclicks antigos do admin/index.html
window.abrirModalColaborador = () => AdminColaboradores.abrirNovo();
window.salvarColaborador = () => AdminColaboradores.salvar();
window.carregarColaboradores = () => AdminColaboradores.carregar();
