// ============================================================
// admin-operacao-dia.js — Operação diária da Luana/cozinha
// Regra: travar pendentes como Principal
// ============================================================

window.AdminOperacaoDia = window.AdminOperacaoDia || {};

Object.assign(window.AdminOperacaoDia, {
  diasSemana: ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"],

  initTravaPendentes() {
    this.inserirControleTravaPendentes();

    document.addEventListener("click", e => {
      const item = e.target.closest(".nav-item, [onclick], button, a");
      if (!item) return;

      const texto = `${item.innerText || ""} ${item.getAttribute("onclick") || ""}`.toLowerCase();

      if (texto.includes("dashboard")) {
        setTimeout(() => this.inserirControleTravaPendentes(), 300);
      }
    });
  },

  estaNoDashboard() {
    const ativo = document.querySelector(".module.active, .nav-item.active");
    const textoAtivo = (ativo?.innerText || "").toLowerCase();

    if (textoAtivo.includes("dashboard")) return true;

    const titulo = document.querySelector(".topbar-title, h1, h2");
    const textoTitulo = (titulo?.innerText || "").toLowerCase();

    return textoTitulo.includes("dashboard");
  },

  inserirControleTravaPendentes() {
    if (!this.estaNoDashboard()) return;
    if (document.getElementById("controleTravaPendentes")) return;

    const avisoSemana =
      Array.from(document.querySelectorAll("div,span,p")).find(el =>
        (el.innerText || "").includes("Semana atual")
      );

    const box = document.createElement("div");
    box.id = "controleTravaPendentes";
    box.style.cssText = `
      margin: 1rem 0;
      padding: 1rem;
      border: 1px solid rgba(192,40,28,.35);
      border-radius: 12px;
      background: rgba(192,40,28,.08);
      display: grid;
      grid-template-columns: 220px 1fr auto;
      gap: .8rem;
      align-items: end;
    `;

    box.innerHTML = `
      <div class="form-group">
        <label class="form-label">DIA PARA TRAVAR</label>
        <select id="diaTravaPendentes" class="form-select">
          <option value="Segunda">Segunda</option>
          <option value="Terça">Terça</option>
          <option value="Quarta">Quarta</option>
          <option value="Quinta">Quinta</option>
          <option value="Sexta">Sexta</option>
        </select>
      </div>

      <div style="font-size:.78rem;color:#ffcf8a;line-height:1.45;">
        Quando o prazo de marcação acabar, a Luana pode travar os pendentes.
        Quem não marcou será preenchido automaticamente como <b>Principal</b>.
      </div>

      <button type="button" class="btn-danger" onclick="AdminOperacaoDia.confirmarTravaPendentes()">
        🔒 Travar pendentes como Principal
      </button>
    `;

    if (avisoSemana) {
      avisoSemana.insertAdjacentElement("afterend", box);
    } else {
      const content = document.querySelector(".content") || document.querySelector("main") || document.body;
      content.appendChild(box);
    }
  },

  obterSemanaAtual() {
    if (typeof getSemanaId === "function") return getSemanaId();

    const texto = document.body.innerText || "";
    const match = texto.match(/(\d{4}-W\d{1,2})/i);

    if (match) return match[1];

    const hoje = new Date();
    const semana = this.numeroSemanaISO(hoje);
    return `${hoje.getFullYear()}-W${String(semana).padStart(2, "0")}`;
  },

  numeroSemanaISO(data) {
    const d = new Date(Date.UTC(data.getFullYear(), data.getMonth(), data.getDate()));
    const dia = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dia);
    const anoInicio = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - anoInicio) / 86400000) + 1) / 7);
  },

  async confirmarTravaPendentes() {
    const dia = document.getElementById("diaTravaPendentes")?.value || "Segunda";
    const semanaId = this.obterSemanaAtual();

    const ok = confirm(
      `Confirmar trava dos pendentes?\n\n` +
      `Semana: ${semanaId}\n` +
      `Dia: ${dia}\n\n` +
      `Todos os colaboradores ativos sem marcação neste dia serão preenchidos como Principal.`
    );

    if (!ok) return;

    try {
      const total = await this.travarPendentesComoPrincipal(semanaId, dia);
      alert(`${total} pendente(s) preenchido(s) como Principal para ${dia}.`);
      location.reload();
    } catch (erro) {
      console.error("Erro ao travar pendentes:", erro);
      alert(`Erro ao travar pendentes: ${erro.message || erro}`);
    }
  },

  async travarPendentesComoPrincipal(semanaId, dia) {
    if (!window.SP) throw new Error("SP não encontrado.");

    const colaboradores = await (SP.getTodosColaboradores ? SP.getTodosColaboradores() : SP.getColaboradores());
    const pedidosSemana = await SP.getPedidos(semanaId);

    const ativos = colaboradores.filter(c => {
      const ativo = SP.isTrue ? SP.isTrue(c.Ativo) : String(c.Ativo).toLowerCase() !== "false";
      return ativo;
    });

    let criados = 0;

    for (const c of ativos) {
      const colaboradorId = String(c.id);

      const jaTemPedido = pedidosSemana.some(p =>
        String(p.Colaborador_id) === colaboradorId &&
        this.normalizar(p.Dia) === this.normalizar(dia)
      );

      if (jaTemPedido) continue;

      await SP.savePedido(
        semanaId,
        colaboradorId,
        c.Nome || c.Title || "",
        dia,
        "Principal",
        "Principal"
      );

      const pedidosColaborador = await SP.getPedidoColaborador(semanaId, colaboradorId);

      const pedidoCriado = pedidosColaborador.find(p =>
        this.normalizar(p.Dia) === this.normalizar(dia)
      );

      if (pedidoCriado) {
        await SP.updatePedido(pedidoCriado.id, {
          Centro_Custo: c.Centro_Custo || "",
          Status: "Confirmado",
          Confirmado: true,
          Origem: "Admin",
          Observacao: "Preenchido automaticamente após prazo de marcação.",
          Alterado_Por: SP.getUserEmail ? SP.getUserEmail() : ""
        });
      }

      criados++;
    }

    return criados;
  },

  normalizar(valor) {
    return String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  }
});

document.addEventListener("DOMContentLoaded", () => {
  setTimeout(() => AdminOperacaoDia.initTravaPendentes(), 800);
});
