// ============================================================
// admin-valores.js — Cadastro de valores da refeição
// Valor Vascon e valor descontado dos funcionários
// ============================================================

window.AdminValores = {
  async carregar() {
    if (!SP.getValoresRefeicao) {
      console.warn("SP.getValoresRefeicao ainda não encontrado no sharepoint.js.");
      return [];
    }
    return SP.getValoresRefeicao();
  },

  async salvar(dados) {
    return SP.createValorRefeicao(dados);
  },

  async editar(id, dados) {
    return SP.updateValorRefeicao(id, dados);
  }
};
