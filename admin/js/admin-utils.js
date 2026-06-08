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
// ============================================================
// Extensão AdminUtils — Exportação Excel formatada padrão Homy
// Requer XLSX carregado no admin/index.html
// ============================================================

Object.assign(AdminUtils, {

  exportarExcelFormatado(config) {
    if (!window.XLSX) {
      this.toast("Biblioteca XLSX não carregada.", "error");
      return;
    }

    const {
      nomeArquivo = "relatorio-homy.xlsx",
      nomeAba = "Relatório",
      titulo = "Relatório Homy",
      periodo = "",
      colunas = [],
      linhas = [],
      totais = null
    } = config || {};

    const aoa = [];

    aoa.push([titulo]);
    if (periodo) aoa.push([periodo]);
    aoa.push([]);
    aoa.push(colunas.map(c => c.label));

    linhas.forEach(linha => {
      aoa.push(colunas.map(c => {
        const valor = linha[c.key];
        return valor === null || valor === undefined ? "" : valor;
      }));
    });

    if (totais) {
      aoa.push([]);
      aoa.push(colunas.map(c => {
        if (totais[c.key] !== undefined) return totais[c.key];
        if (c.totalLabel) return c.totalLabel;
        return "";
      }));
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    const totalCols = colunas.length;
    const headerRowIndex = periodo ? 4 : 3;
    const headerRowNumber = headerRowIndex;
    const lastRow = aoa.length;

    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(totalCols - 1, 0) } }
    ];

    if (periodo) {
      ws["!merges"].push({
        s: { r: 1, c: 0 },
        e: { r: 1, c: Math.max(totalCols - 1, 0) }
      });
    }

    ws["!cols"] = colunas.map(c => ({
      wch: c.width || 18
    }));

    ws["!autofilter"] = {
      ref: XLSX.utils.encode_range({
        s: { r: headerRowIndex - 1, c: 0 },
        e: { r: Math.max(lastRow - 1, headerRowIndex - 1), c: totalCols - 1 }
      })
    };

    ws["!freeze"] = {
      xSplit: 0,
      ySplit: headerRowIndex,
      topLeftCell: `A${headerRowIndex + 1}`,
      activePane: "bottomLeft",
      state: "frozen"
    };

    const range = XLSX.utils.decode_range(ws["!ref"]);

    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[cellRef]) continue;

        ws[cellRef].s = ws[cellRef].s || {};

        ws[cellRef].s.font = {
          name: "Calibri",
          sz: 11,
          color: { rgb: "111111" }
        };

        ws[cellRef].s.alignment = {
          vertical: "center",
          horizontal: "left",
          wrapText: true
        };

        ws[cellRef].s.border = {
          top: { style: "thin", color: { rgb: "D9E2F3" } },
          bottom: { style: "thin", color: { rgb: "D9E2F3" } },
          left: { style: "thin", color: { rgb: "D9E2F3" } },
          right: { style: "thin", color: { rgb: "D9E2F3" } }
        };

        if (R === 0) {
          ws[cellRef].s.font = {
            name: "Calibri",
            sz: 16,
            bold: true,
            color: { rgb: "FFFFFF" }
          };
          ws[cellRef].s.fill = {
            fgColor: { rgb: "0B1A35" }
          };
          ws[cellRef].s.alignment = {
            vertical: "center",
            horizontal: "center"
          };
        }

        if (periodo && R === 1) {
          ws[cellRef].s.font = {
            name: "Calibri",
            sz: 11,
            bold: true,
            color: { rgb: "1A3A6E" }
          };
          ws[cellRef].s.fill = {
            fgColor: { rgb: "EAF2FF" }
          };
          ws[cellRef].s.alignment = {
            vertical: "center",
            horizontal: "center"
          };
        }

        if (R === headerRowIndex - 1) {
          ws[cellRef].s.font = {
            name: "Calibri",
            sz: 11,
            bold: true,
            color: { rgb: "FFFFFF" }
          };
          ws[cellRef].s.fill = {
            fgColor: { rgb: "1A3A6E" }
          };
          ws[cellRef].s.alignment = {
            vertical: "center",
            horizontal: "center",
            wrapText: true
          };
        }

        if (totais && R === range.e.r) {
          ws[cellRef].s.font = {
            name: "Calibri",
            sz: 11,
            bold: true,
            color: { rgb: "111111" }
          };
          ws[cellRef].s.fill = {
            fgColor: { rgb: "D9EAD3" }
          };
        }
      }
    }

    ws["!rows"] = aoa.map((_, index) => {
      if (index === 0) return { hpt: 26 };
      if (periodo && index === 1) return { hpt: 20 };
      if (index === headerRowIndex - 1) return { hpt: 22 };
      return { hpt: 18 };
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, nomeAba.slice(0, 31));

    XLSX.writeFile(wb, nomeArquivo, {
      bookType: "xlsx",
      cellStyles: true
    });
  },

  limparTextoRelatorio(valor) {
    const texto = String(valor ?? "").trim();
    if (!texto || texto === "—" || texto === "-") return "";
    return texto;
  },

  nomeArquivoSeguro(valor) {
    return String(valor || "relatorio")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .toLowerCase();
  }

});
