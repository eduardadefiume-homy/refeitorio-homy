// admin-utils.js — Utilitários compartilhados do Admin Homy
// Sem dependências externas. Carregado antes de todos os outros módulos.

const AdminUtils = window.AdminUtils = {

  // ── Toast ──────────────────────────────────────────────────
  toast(msg, type = "info") {
    const container = document.getElementById("toast");
    if (!container) { console.warn("[toast]", msg); return; }
    const el = document.createElement("div");
    el.className = `toast-item toast-${type}`;
    el.textContent = msg;
    container.appendChild(el);
    setTimeout(() => el.remove(), 3800);
  },

  // ── Semana ISO ──────────────────────────────────────────────
  getSemanaId(date = new Date()) {
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
    const w1 = new Date(d.getFullYear(), 0, 4);
    const wn = 1 + Math.round(((d - w1) / 86400000 - 3 + ((w1.getDay() + 6) % 7)) / 7);
    return `${d.getFullYear()}-W${String(wn).padStart(2, "0")}`;
  },

  getSemanaIdByOffset(offset = 0) {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    const day = d.getDay();
    const diffMon = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + diffMon + offset * 7);
    return this.getSemanaId(d);
  },

  getMondayByOffset(offset = 0) {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    const day = d.getDay();
    const diffMon = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + diffMon + offset * 7);
    return d;
  },

  formatSemana(semanaId) {
    const [year, week] = semanaId.split("-W");
    return `Semana ${week} / ${year}`;
  },

  formatDateBR(date) {
    return new Date(date).toLocaleDateString("pt-BR", {
      day: "2-digit", month: "2-digit", year: "numeric"
    });
  },

  formatDateTimeBR(date) {
    return new Date(date).toLocaleString("pt-BR");
  },

  // ── Texto ───────────────────────────────────────────────────
  norm(v) {
    return String(v || "").normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
  },

  esc(v) {
    return String(v ?? "").replace(/[&<>"']/g, c =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
    );
  },

  pick(obj, ...keys) {
    for (const k of keys) {
      if (obj && obj[k] !== undefined && obj[k] !== null) return obj[k];
    }
    return "";
  },

  moeda(v) {
    if (v === null || v === undefined || String(v).trim() === "") return null;
    let s = String(v).replace(/R\$/gi, "").replace(/\s/g, "").trim();
    if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
    s = s.replace(/[^0-9.-]/g, "");
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  },

  // ── DOM ─────────────────────────────────────────────────────
  setTxt(id, val) {
    const el = document.getElementById(id);
    if (el) el.textContent = val ?? "—";
  },

  getVal(id) {
    return (document.getElementById(id)?.value || "").trim();
  },

  setVal(id, val) {
    const el = document.getElementById(id);
    if (el) el.value = val ?? "";
  },

  // ── Badge de status ─────────────────────────────────────────
  badgeClass(status) {
    const s = this.norm(status);
    if (["confirmado", "extra", "aprovado"].includes(s)) return "badge-green";
    if (["cancelado", "afastado", "ferias", "bloqueado", "travado",
         "nao vai almocar"].includes(s)) return "badge-red";
    return "badge-yellow";
  },

  badge(status, label) {
    const cls = this.badgeClass(status);
    return `<span class="badge ${cls}">${this.esc(label || status)}</span>`;
  },

  // ── Dia da semana ───────────────────────────────────────────
  DIAS: ["segunda", "terca", "quarta", "quinta", "sexta"],
  DIA_LABEL: { segunda: "Segunda", terca: "Terça", quarta: "Quarta", quinta: "Quinta", sexta: "Sexta" },
  DIA_HOJE() {
    const map = ["domingo", "segunda", "terca", "quarta", "quinta", "sexta", "sabado"];
    const d = map[new Date().getDay()];
    return this.DIAS.includes(d) ? d : "segunda";
  },

  // ── Modal genérico ──────────────────────────────────────────
  openModal(id) {
    const m = document.getElementById(id);
    if (!m) return;
    m.style.display = "flex";
    m.style.opacity = "";
    m.style.pointerEvents = "";
    m.removeAttribute("aria-hidden");
    m.classList.add("open");
  },

  closeModal(id) {
    const m = document.getElementById(id);
    if (!m) return;
    m.classList.remove("open", "show", "active");
    m.style.display = "none";
    m.setAttribute("aria-hidden", "true");
  },

  bindModalClose() {
    document.addEventListener("click", e => {
      if (e.target?.classList?.contains("modal-overlay")) this.closeModal(e.target.id);
      if (e.target?.classList?.contains("modal-close")) {
        const overlay = e.target.closest(".modal-overlay");
        if (overlay) this.closeModal(overlay.id);
      }
    }, true);

    document.addEventListener("keydown", e => {
      if (e.key === "Escape") {
        document.querySelectorAll(".modal-overlay.open").forEach(m => this.closeModal(m.id));
      }
    });
  }
};
