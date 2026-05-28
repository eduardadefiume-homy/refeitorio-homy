// ============================================================
// admin-utils.js — Funções utilitárias reutilizáveis
// ============================================================

window.AdminUtils = {
  moeda(valor) {
    const n = Number(valor || 0);
    return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  },

  dataBR(valor) {
    if (!valor) return "";
    const d = new Date(valor);
    if (Number.isNaN(d.getTime())) return "";
    return d.toLocaleDateString("pt-BR");
  },

  periodoTexto(inicio, fim) {
    return `Período de ${this.dataBR(inicio)} a ${this.dataBR(fim)}`;
  },

  normalizarTexto(valor) {
    return String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  },

  abrirModal(id) {
    const modal = document.getElementById(id);
    if (modal) modal.classList.add("open");
  },

  fecharModal(id) {
    const modal = document.getElementById(id);
    if (modal) modal.classList.remove("open");
  },

  toast(msg, type = "info") {
    if (typeof window.toast === "function") {
      window.toast(msg, type);
      return;
    }
    console.log(`[${type}] ${msg}`);
  }
};
