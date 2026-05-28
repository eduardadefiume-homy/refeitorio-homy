// ============================================================
// admin-colaboradores.js — CRUD correto de colaboradores
// Correção: editar atualiza item existente; excluir remove de verdade
// ============================================================

window.AdminColaboradores = {
  lista: [],
  editandoId: null,

  async carregar() {
    try {
      if (!window.SP) throw new Error("SP não encontrado.");

      const getFn = typeof SP.getTodosColaboradores === "function"
        ? "getTodosColaboradores"
        : "getColaboradores";

      this.lista = await SP[getFn]();
      this.renderizarTabela();
    } catch (erro) {
      console.error("Erro ao carregar colaboradores:", erro);
      this.renderizarMensagem(`Erro ao carregar colaboradores: ${erro.message || erro}`, true);
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
        (tb.innerText || "").toLowerCase().includes("colaborador")
      )
    );
  },

  renderizarMensagem(msg, erro = false) {
    const tbody = this.encontrarTbody();
    if (!tbody) return;

    tbody.innerHTML = `
      <tr>
        <td colspan="6" style="text-align:center;${erro ? "color:#ff6060;" : "opacity:.55;"}padding:2rem;">
          ${msg}
        </td>
      </tr>
    `;
  },

  renderizarTabela() {
    const tbody = this.encontrarTbody();

    if (!tbody) {
      console.warn("Tabela de colaboradores não encontrada.");
      return;
    }

    const busca = this.normalizar(document.getElementById("buscaColaborador")?.value || "");

    const filtrados = this.lista.filter(c => {
      const nome = this.normalizar(c.Nome || c.Title || "");
      const departamento = this.normalizar(c.Departamento || "");
      const centro = this.normalizar(c.Centro_Custo || "");
      const email = this.normalizar(c.Email || "");
      return !busca || nome.includes(busca) || departamento.includes(busca) || centro.includes(busca) || email.includes(busca);
    });

    if (!filtrados.length) {
      this.renderizarMensagem("Nenhum colaborador encontrado.");
      return;
    }

    tbody.innerHTML = filtrados.map(c => {
      const ativo = typeof SP.isTrue === "function"
        ? SP.isTrue(c.Ativo)
        : String(c.Ativo).toLowerCase() !== "false";

      return `
        <tr data-colaborador-id="${c.id}">
          <td>${this.escape(c.Nome || c.Title || "")}</td>
          <td>${this.escape(c.Departamento || "")}</td>
          <td>${this.escape(c.Centro_Custo || "-")}</td>
          <td><span class="badge badge-blue">${this.escape(c.tipo || c.Tipo || "Colaborador")}</span></td>
          <td><span class="badge ${ativo ? "badge-green" : "badge-red"}">${ativo ? "ATIVO" : "INATIVO"}</span></td>
          <td>
            <div class="table-actions">
              <button type="button" class="btn-icon" title="Editar" onclick="AdminColaboradores.abrirEditar('${c.id}')">✏️</button>
              <button type="button" class="btn-icon danger" title="Excluir" onclick="AdminColaboradores.excluir('${c.id}')">🗑️</button>
            </div>
          </td>
        </tr>
      `;
    }).join("");
  },

  normalizar(valor) {
    return String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  },

  escape(valor) {
    return String(valor ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  },

  abrirNovo() {
    this.editandoId = null;
    this.preencherModal({});
    this.setTituloModal("NOVO COLABORADOR");
    this.abrirModal();
  },

  abrirEditar(id) {
    const c = this.lista.find(x => String(x.id) === String(id));

    if (!c) {
      alert("Colaborador não encontrado para edição.");
      return;
    }

    this.editandoId = String(id);
    this.preencherModal(c);
    this.setTituloModal("EDITAR COLABORADOR");
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
    this.editandoId = null;
  },

  setTituloModal(texto) {
    const titulo =
      document.getElementById("modalColaboradorTitulo") ||
      document.querySelector("#modalColaborador .modal-title") ||
      document.querySelector("#modalNovoColaborador .modal-title") ||
      Array.from(document.querySelectorAll(".modal-title,h2,h3")).find(el =>
        (el.innerText || "").toUpperCase().includes("COLABORADOR")
      );

    if (titulo) titulo.textContent = texto;
  },

  preencherModal(c) {
    this.setCampo(["colabNome", "nomeColaborador"], c.Nome || c.Title || "");
    this.setCampo(["colabDepartamento", "departamentoColaborador"], c.Departamento || "");
    this.setCampo(["colabEmail", "emailColaborador"], c.Email || "");
    this.setCampo(["colabTipo", "tipoColaborador"], c.tipo || c.Tipo || "Colaborador");
    this.setCampo(["colabCentroCusto", "centroCustoColaborador"], c.Centro_Custo || "");
  },

  setCampo(ids, valor) {
    for (const id of ids) {
      const el = document.getElementById(id);
      if (el) {
        el.value = valor || "";
        return;
      }
    }
  },

  getCampo(ids) {
    for (const id of ids) {
      const el = document.getElementById(id);
      if (el) return el.value.trim();
    }
    return "";
  },

  obterDadosFormulario() {
    const centroCusto = this.getCampo(["colabCentroCusto", "centroCustoColaborador"]);

    return {
      nome: this.getCampo(["colabNome", "nomeColaborador"]),
      departamento: this.getCampo(["colabDepartamento", "departamentoColaborador"]),
      email: this.getCampo(["colabEmail", "emailColaborador"]),
      tipo: this.getCampo(["colabTipo", "tipoColaborador"]) || "Colaborador",
      centroCusto,
      Centro_Custo: centroCusto
    };
  },

  async salvar() {
    try {
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
    } catch (erro) {
      console.error("Erro ao salvar colaborador:", erro);
      alert(`Erro ao salvar colaborador: ${erro.message || erro}`);
    }
  },

  async excluir(id) {
    const colaborador = this.lista.find(x => String(x.id) === String(id));
    const nome = colaborador ? (colaborador.Nome || colaborador.Title || "") : "";

    const ok = confirm(
      `Deseja EXCLUIR definitivamente este colaborador${nome ? `: ${nome}` : ""}?\n\n` +
      "Use isto somente para cadastros duplicados ou criados por engano."
    );

    if (!ok) return;

    try {
      if (typeof SP.deleteColaborador === "function") {
        await SP.deleteColaborador(id);
      } else {
        await SP.deleteItem("Colaboradores", id);
      }

      await this.carregar();
      alert("Colaborador excluído com sucesso.");
    } catch (erro) {
      console.error("Erro ao excluir colaborador:", erro);
      alert(`Erro ao excluir colaborador: ${erro.message || erro}`);
    }
  },

  async desativar(id) {
    const ok = confirm("Deseja apenas desativar este colaborador?");
    if (!ok) return;

    await SP.desativarColaborador(id);
    await this.carregar();
  }
};

// Sobrescreve funções antigas do HTML para impedir duplicação
window.abrirModalColaborador = () => AdminColaboradores.abrirNovo();
window.salvarColaborador = () => AdminColaboradores.salvar();
window.carregarColaboradores = () => AdminColaboradores.carregar();
window.fecharModalColaborador = () => AdminColaboradores.fecharModal();

document.addEventListener("click", function(e) {
  const alvo = e.target.closest("[onclick], .nav-item, button, a");
  if (!alvo) return;

  const texto = `${alvo.innerText || ""} ${alvo.getAttribute("onclick") || ""}`.toLowerCase();

  if (texto.includes("colaboradores")) {
    setTimeout(() => AdminColaboradores.carregar(), 250);
  }
});

document.addEventListener("input", function(e) {
  if (e.target && e.target.id === "buscaColaborador") {
    AdminColaboradores.renderizarTabela();
  }
});

document.addEventListener("DOMContentLoaded", function() {
  setTimeout(() => {
    const textoPagina = (document.body.innerText || "").toLowerCase();
    if (textoPagina.includes("colaboradores")) {
      AdminColaboradores.carregar();
    }
  }, 1000);
});
