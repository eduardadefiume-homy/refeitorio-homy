// ============================================================
// admin-colaboradores.js — CRUD de colaboradores
// Correção: renderização robusta da tabela Colaboradores
// ============================================================

window.AdminColaboradores = {
  lista: [],
  editandoId: null,

  async carregar() {
    try {
      if (!window.SP) {
        throw new Error("SP não encontrado. Verifique o carregamento do sharepoint.js.");
      }

      const getFn = typeof SP.getTodosColaboradores === "function"
        ? "getTodosColaboradores"
        : "getColaboradores";

      this.lista = await SP[getFn]();
      this.renderizarTabela();
    } catch (erro) {
      console.error("Erro ao carregar colaboradores:", erro);
      this.renderizarErro(erro);
    }
  },

  encontrarTbody() {
    return (
      document.getElementById("colaboradoresTableBody") ||
      document.getElementById("tbodyColaboradores") ||
      document.querySelector("[data-colaboradores-tbody]") ||
      document.querySelector("#colaboradores tbody") ||
      document.querySelector("#module-colaboradores tbody") ||
      Array.from(document.querySelectorAll("tbody")).find(tb =>
        tb.innerText && tb.innerText.toLowerCase().includes("carregando colaboradores")
      )
    );
  },

  renderizarErro(erro) {
    const tbody = this.encontrarTbody();
    if (!tbody) return;
    tbody.innerHTML = `
      <tr>
        <td colspan="6" style="text-align:center;color:#ff6060;padding:2rem;">
          Erro ao carregar colaboradores: ${erro.message || erro}
        </td>
      </tr>
    `;
  },

  renderizarTabela() {
    const tbody = this.encontrarTbody();

    if (!tbody) {
      console.warn("Tabela de colaboradores não encontrada no HTML atual.");
      return;
    }

    const busca = (document.getElementById("buscaColaborador")?.value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();

    const normalizar = valor => String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();

    const filtrados = this.lista.filter(c => {
      const nome = normalizar(c.Nome || c.Title || "");
      const departamento = normalizar(c.Departamento || "");
      const centro = normalizar(c.Centro_Custo || "");
      return !busca || nome.includes(busca) || departamento.includes(busca) || centro.includes(busca);
    });

    if (!filtrados.length) {
      tbody.innerHTML = `
        <tr>
          <td colspan="6" style="text-align:center;opacity:.55;padding:2rem;">
            Nenhum colaborador encontrado.
          </td>
        </tr>
      `;
      return;
    }

    tbody.innerHTML = filtrados.map(c => {
      const ativo = typeof SP.isTrue === "function"
        ? SP.isTrue(c.Ativo)
        : String(c.Ativo).toLowerCase() !== "false";

      return `
        <tr>
          <td>${c.Nome || c.Title || ""}</td>
          <td>${c.Departamento || ""}</td>
          <td>${c.Centro_Custo || "-"}</td>
          <td><span class="badge badge-blue">${c.tipo || c.Tipo || "Colaborador"}</span></td>
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
    this.abrirModal();
  },

  abrirModal() {
    const modal =
      document.getElementById("modalColaborador") ||
      document.getElementById("modalNovoColaborador");
    if (modal) modal.classList.add("open");
  },

  fecharModal() {
    const modal =
      document.getElementById("modalColaborador") ||
      document.getElementById("modalNovoColaborador");
    if (modal) modal.classList.remove("open");
  },

  abrirEditar(id) {
    const c = this.lista.find(x => String(x.id) === String(id));
    if (!c) {
      alert("Colaborador não encontrado para edição.");
      return;
    }

    this.editandoId = id;
    this.preencherModal(c);
    this.abrirModal();
  },

  preencherModal(c) {
    const set = (ids, value) => {
      for (const id of ids) {
        const el = document.getElementById(id);
        if (el) {
          el.value = value || "";
          return;
        }
      }
    };

    set(["colabNome", "nomeColaborador"], c.Nome || c.Title || "");
    set(["colabDepartamento", "departamentoColaborador"], c.Departamento || "");
    set(["colabEmail", "emailColaborador"], c.Email || "");
    set(["colabTipo", "tipoColaborador"], c.tipo || c.Tipo || "Colaborador");
    set(["colabCentroCusto", "centroCustoColaborador"], c.Centro_Custo || "");

    const titulo =
      document.getElementById("modalColaboradorTitulo") ||
      document.querySelector("#modalColaborador .modal-title") ||
      document.querySelector("#modalNovoColaborador .modal-title");

    if (titulo) titulo.textContent = this.editandoId ? "EDITAR COLABORADOR" : "NOVO COLABORADOR";
  },

  obterDadosFormulario() {
    const val = ids => {
      for (const id of ids) {
        const el = document.getElementById(id);
        if (el) return el.value.trim();
      }
      return "";
    };

    const centroCusto = val(["colabCentroCusto", "centroCustoColaborador"]);

    return {
      nome: val(["colabNome", "nomeColaborador"]),
      departamento: val(["colabDepartamento", "departamentoColaborador"]),
      email: val(["colabEmail", "emailColaborador"]),
      tipo: val(["colabTipo", "tipoColaborador"]) || "Colaborador",
      centroCusto,
      Centro_Custo: centroCusto
    };
  },

  async salvar() {
    const dados = this.obterDadosFormulario();

    if (!dados.nome) {
      alert("Informe o nome do colaborador.");
      return;
    }

    if (!dados.centroCusto) {
      alert("Informe o Centro de Custo do colaborador.");
      return;
    }

    if (this.editandoId) {
      await SP.updateColaborador(this.editandoId, dados);
      alert("Colaborador atualizado com sucesso.");
    } else {
      await SP.createColaborador(dados);
      alert("Colaborador cadastrado com sucesso.");
    }

    this.fecharModal();
    await this.carregar();
  },

  async desativar(id) {
    const ok = confirm("Deseja desativar este colaborador? Ele não será removido do histórico.");
    if (!ok) return;

    await SP.desativarColaborador(id);
    await this.carregar();
  }
};

// Compatibilidade com funções antigas do HTML
window.abrirModalColaborador = () => AdminColaboradores.abrirNovo();
window.salvarColaborador = () => AdminColaboradores.salvar();
window.carregarColaboradores = () => AdminColaboradores.carregar();

// Reforço: quando clicar na aba Colaboradores, recarrega a tabela
document.addEventListener("click", function(e) {
  const alvo = e.target.closest("[onclick], .nav-item, button, a");
  if (!alvo) return;

  const texto = (alvo.innerText || alvo.getAttribute("onclick") || "").toLowerCase();
  if (texto.includes("colaboradores")) {
    setTimeout(() => AdminColaboradores.carregar(), 250);
  }
});

// Se a tela já abrir na aba colaboradores ou após o carregamento inicial
document.addEventListener("DOMContentLoaded", function() {
  setTimeout(() => {
    const ativo = document.querySelector(".nav-item.active, .module.active");
    const texto = ativo ? ativo.innerText.toLowerCase() : "";
    if (texto.includes("colaboradores")) {
      AdminColaboradores.carregar();
    }
  }, 800);
});
