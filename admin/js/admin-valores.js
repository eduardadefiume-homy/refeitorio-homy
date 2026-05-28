// ============================================================
// admin-valores.js — Cadastro financeiro refeições
// ============================================================

window.AdminValores = {
  lista: [],

  async carregar() {
    if (!window.SP || !SP.getValoresRefeicao) {
      console.warn("Funções de valores ainda não encontradas.");
      return;
    }

    this.lista = await SP.getValoresRefeicao();
    console.log("Valores carregados:", this.lista);
  },

  async salvar(dados) {
    return SP.createValorRefeicao(dados);
  },

  async editar(id, dados) {
    return SP.updateValorRefeicao(id, dados);
  }
};
