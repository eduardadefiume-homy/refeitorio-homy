// ============================================================
// admin-operacao-dia.js — Operação diária da Luana/cozinha
// ============================================================

window.AdminOperacaoDia = {
  async carregar() {
    console.log("Carregar operação do dia.");
  },

  async cancelarPedido(id, observacao = "") {
    return SP.updatePedido(id, {
      Status: "Cancelado",
      Observacao: observacao,
      Origem: "Admin",
      Alterado_Por: window.AdminState?.usuarioAtual?.name || ""
    });
  }
};
