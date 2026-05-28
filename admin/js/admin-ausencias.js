// ============================================================
// admin-ausencias.js — Férias, afastamentos, ausências e bloqueios
// ============================================================

window.AdminAusencias = {
  async carregar() {
    if (!SP.getAusenciasRefeitorio) {
      console.warn("SP.getAusenciasRefeitorio ainda não encontrado no sharepoint.js.");
      return [];
    }
    return SP.getAusenciasRefeitorio();
  },

  async salvar(dados) {
    return SP.createAusenciaRefeitorio(dados);
  },

  async editar(id, dados) {
    return SP.updateAusenciaRefeitorio(id, dados);
  }
};
