// ============================================================
// admin-colaboradores.js — CRUD de colaboradores
// Inclui Centro_Custo
// ============================================================

window.AdminColaboradores = {
  async carregar() {
    console.log("Carregar colaboradores.");
  },

  async salvar(dados) {
    return SP.createColaborador({
      nome: dados.nome,
      departamento: dados.departamento,
      email: dados.email,
      tipo: dados.tipo,
      centroCusto: dados.centroCusto || dados.Centro_Custo || ""
    });
  },

  async editar(id, dados) {
    return SP.updateColaborador(id, dados);
  },

  async excluirOuDesativar(id) {
    return SP.desativarColaborador(id);
  }
};
