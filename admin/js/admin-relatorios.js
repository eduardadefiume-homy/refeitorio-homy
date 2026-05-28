// ============================================================
// admin-relatorios.js — Relatórios financeiros e operacionais
// ============================================================

window.AdminRelatorios = {
  async porPeriodo(dataInicio, dataFim) {
    console.log("Relatório por período", dataInicio, dataFim);
  },

  async porCentroCusto(dataInicio, dataFim) {
    console.log("Relatório por centro de custo", dataInicio, dataFim);
  },

  async porColaborador(dataInicio, dataFim) {
    console.log("Relatório por colaborador", dataInicio, dataFim);
  }
};
