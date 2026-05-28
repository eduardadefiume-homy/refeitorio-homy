// ============================================================
// admin-core.js — Inicialização, login e navegação do Admin
// ============================================================

window.AdminCore = {
  async init() {
    console.log("AdminCore iniciado.");
    // Nesta fase, o index.html legado continua comandando a tela.
    // Vamos migrar as chamadas para cá módulo por módulo, sem quebrar o que já funciona.
  },

  irParaAba(nomeModulo) {
    if (typeof window.showModule === "function") {
      window.showModule(nomeModulo);
      return;
    }
    document.querySelectorAll(".module").forEach(m => m.classList.remove("active"));
    const el = document.getElementById(nomeModulo);
    if (el) el.classList.add("active");
  }
};

document.addEventListener("DOMContentLoaded", () => {
  if (window.AdminCore) window.AdminCore.init();
});
