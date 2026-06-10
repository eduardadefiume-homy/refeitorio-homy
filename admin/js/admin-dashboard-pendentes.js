// ============================================================
// admin-dashboard-pendentes.js
// Botão de trava de pendentes no Dashboard
// Regra preservada: quem não marcou vira Principal Confirmado
// ============================================================

const AdminDashboardPendentes = window.AdminDashboardPendentes = {

  ID_BOX: "controleTravaPendentesUnico",

  init() {
    this.inserirBotao();
  },

  normalizar(valor) {
    return String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  },

  normalizarDiaParaSP(dia) {
    const n = this.normalizar(dia);

    if (n.startsWith("seg")) return "segunda";
    if (n.startsWith("ter")) return "terca";
    if (n.startsWith("qua")) return "quarta";
    if (n.startsWith("qui")) return "quinta";
    if (n.startsWith("sex")) return "sexta";

    return n;
  },

  limparControlesAntigos() {
    document.querySelectorAll("[id^='controleTravaPendentes']").forEach(el => {
      if (el.id !== this.ID_BOX) el.remove();
    });
  },

  estaNoDashboard() {
    const ativoModulo = document.querySelector("#mod-dashboard.module.active");
    if (ativoModulo) return true;

    const ativo = document.querySelector(".nav-item.active");
    if (this.normalizar(ativo?.innerText || "").includes("dashboard")) return true;

    const titulo = document.getElementById("topbarTitle") || document.querySelector(".topbar-title, h1, h2");
    return this.normalizar(titulo?.innerText || "").includes("dashboard");
  },

  encontrarFaixaSemanaAtual() {
    const ids = ["semanaLabel", "dashboardSemanaLabel", "dashSemanaLabel"];

    for (const id of ids) {
      const el = document.getElementById(id);
      if (el) return el.closest(".info-box, .alert, .card, div") || el;
    }

    const candidatos = Array.from(document.querySelectorAll("#mod-dashboard div, .module.active div"));

    const achado = candidatos
      .filter(el => {
        const texto = el.innerText || "";
        return (
          texto.includes("Semana:") ||
          texto.includes("Semana atual:") ||
          texto.includes("dados exibidos") ||
          /\d{4}-W\d{1,2}/i.test(texto)
        );
      })
      .sort((a, b) => {
        const areaA = a.offsetWidth * a.offsetHeight;
        const areaB = b.offsetWidth * b.offsetHeight;
        return areaA - areaB;
      })[0];

    return achado || document.querySelector("#mod-dashboard .module-body") || document.querySelector("#mod-dashboard");
  },

  obterSemanaAtual() {
    if (window.AdminState && typeof AdminState.getSemanaId === "function") {
      return AdminState.getSemanaId();
    }

    if (window.SP && typeof SP.getCurrentWeekId === "function") {
      return SP.getCurrentWeekId();
    }

    const texto = document.body.innerText || "";
    const match = texto.match(/(\d{4}-W\d{1,2})/i);
    if (match) return match[1];

    return "";
  },

  inserirBotao() {
    this.limparControlesAntigos();

    if (!this.estaNoDashboard()) return;
    if (document.getElementById(this.ID_BOX)) return;

    const faixaSemana = this.encontrarFaixaSemanaAtual();
    if (!faixaSemana) return;

    const diaSelecionado = localStorage.getItem("diaTravaPendentes") || "Segunda";

    const box = document.createElement("div");
    box.id = this.ID_BOX;
    box.style.cssText = `
      width: 100%;
      margin: 10px 0 0 0;
      padding: 10px 14px;
      border: 1px solid rgba(192,40,28,.35);
      border-radius: 10px;
      background: rgba(192,40,28,.08);
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      box-sizing: border-box;
      min-height: 54px;
      overflow: visible;
      position: static;
      clear: both;
      z-index: 20;
    `;

    box.innerHTML = `
      <div style="display:flex;align-items:center;gap:12px;min-width:0;flex:1;flex-wrap:wrap;">
        <div style="display:flex;flex-direction:column;gap:4px;width:145px;flex:0 0 145px;">
          <label class="form-label" style="font-size:10px;margin:0;">DIA</label>
          <select id="diaTravaPendentesUnico" class="form-select" style="height:34px;padding:4px 8px;">
            <option value="Segunda">Segunda</option>
            <option value="Terça">Terça</option>
            <option value="Quarta">Quarta</option>
            <option value="Quinta">Quinta</option>
            <option value="Sexta">Sexta</option>
          </select>
        </div>

        <div style="font-size:12px;color:#ffcf8a;line-height:1.3;white-space:normal;min-width:220px;flex:1;">
          Após o prazo, preenche automaticamente como <b>Principal</b> quem ficou pendente.
        </div>
      </div>

      <button type="button" class="btn-danger" id="btnTravarPendentesUnico"
        style="height:34px;padding:0 14px;white-space:nowrap;flex:0 0 auto;">
        🔒 Travar pendentes
      </button>
    `;

    faixaSemana.insertAdjacentElement("afterend", box);

    const select = document.getElementById("diaTravaPendentesUnico");
    if (select) {
      select.value = diaSelecionado;
      select.addEventListener("change", function () {
        localStorage.setItem("diaTravaPendentes", this.value);
      });
    }

    document.getElementById("btnTravarPendentesUnico")?.addEventListener("click", () => {
      this.travarPendentesComoPrincipal();
    });
  },

  async salvarPedidoPrincipal({ semanaId, colaborador, dia }) {
    const colaboradorId = String(colaborador.id || colaborador.ID || "");
    const colaboradorNome = colaborador.Nome || colaborador.Title || "";
    const centroCusto = colaborador.Centro_Custo || "";
    const diaNormalizado = this.normalizarDiaParaSP(dia);

    if (!window.SP || typeof SP.savePedido !== "function") {
      throw new Error("SP.savePedido não encontrado.");
    }

    return SP.savePedido({
      semanaId,
      colaboradorId,
      colaboradorNome,
      centroCusto,
      dia: diaNormalizado,
      opcao: "principal",
      nomePrato: "Principal",
      confirmado: true,
      status: "Confirmado",
      origem: "Admin",
      dataHora: new Date().toISOString(),
      observacao: "Preenchido automaticamente após prazo de marcação.",
      Alterado_Por: SP.getUserEmail ? SP.getUserEmail() : ""
    });
  },

  async travarPendentesComoPrincipal() {
    try {
      if (!window.SP) {
        alert("SP não encontrado. Recarregue a página e tente novamente.");
        return;
      }

      await SP.init();

      const select = document.getElementById("diaTravaPendentesUnico");
      const dia = select?.value || localStorage.getItem("diaTravaPendentes") || "Segunda";
      const semanaId = this.obterSemanaAtual();
      const diaNormalizado = this.normalizarDiaParaSP(dia);

      if (!semanaId) {
        alert("Não foi possível identificar a semana.");
        return;
      }

      const ok = confirm(
        `Confirmar trava dos pendentes?\n\n` +
        `Semana: ${semanaId}\n` +
        `Dia: ${dia}\n\n` +
        `Quem não marcou será criado automaticamente como Principal.`
      );

      if (!ok) return;

      const btn = document.getElementById("btnTravarPendentesUnico");

      if (btn) {
        btn.disabled = true;
        btn.textContent = "Travando...";
      }

      const colaboradores = await (
        SP.getTodosColaboradores
          ? SP.getTodosColaboradores()
          : SP.getColaboradores()
      );

      const pedidosSemana = await SP.getPedidos(semanaId);

      const ativos = colaboradores.filter(c => {
        if (SP.isTrue) return SP.isTrue(c.Ativo);
        return String(c.Ativo).toLowerCase() !== "false";
      });

      let criados = 0;
      let jaExistiam = 0;
      let falhas = 0;

      for (const c of ativos) {
        const colaboradorId = String(c.id || c.ID || "");

        if (!colaboradorId) continue;

        const jaTemPedido = pedidosSemana.some(p =>
          String(p.Colaborador_id || p.colaboradorId || "") === colaboradorId &&
          this.normalizar(p.Dia) === this.normalizar(diaNormalizado)
        );

        if (jaTemPedido) {
          jaExistiam++;
          continue;
        }

        try {
          await this.salvarPedidoPrincipal({
            semanaId,
            colaborador: c,
            dia: diaNormalizado
          });

          criados++;
        } catch (erro) {
          falhas++;
          console.error("Falha ao travar colaborador:", c.Nome || c.Title, erro);
        }
      }

      alert(
        `${criados} pendente(s) preenchido(s) como Principal para ${dia}.` +
        `\n${jaExistiam} já tinham registro.` +
        (falhas ? `\n${falhas} falha(s). Verifique o console.` : "")
      );

      if (window.AdminDashboard && window.AdminState) {
        await AdminDashboard.load(AdminState.getSemanaId());
      }

    } catch (erro) {
      console.error("Erro ao travar pendentes:", erro);
      alert(`Erro ao travar pendentes: ${erro.message || erro}`);
    } finally {
      const btn = document.getElementById("btnTravarPendentesUnico");
      if (btn) {
        btn.disabled = false;
        btn.textContent = "🔒 Travar pendentes";
      }
    }
  }
};

document.addEventListener("DOMContentLoaded", () => {
  setTimeout(() => AdminDashboardPendentes.init(), 500);
});
