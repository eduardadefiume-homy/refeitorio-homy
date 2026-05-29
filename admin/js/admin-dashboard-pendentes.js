// ============================================================
// admin-dashboard-pendentes.js
// Posição final + select estável
// ============================================================

(function () {
  const ID_BOX = "controleTravaPendentesUnico";
  let diaSelecionado = localStorage.getItem("diaTravaPendentes") || "Segunda";

  function normalizar(valor) {
    return String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  }

  function limparControlesAntigos() {
    document.querySelectorAll("[id^='controleTravaPendentes']").forEach(el => {
      if (el.id !== ID_BOX) el.remove();
    });
  }

  function estaNoDashboard() {
    const titulo = document.querySelector(".topbar-title, h1, h2");
    if (normalizar(titulo?.innerText || "").includes("dashboard")) return true;

    const ativo = document.querySelector(".nav-item.active, .module.active");
    return normalizar(ativo?.innerText || "").includes("dashboard");
  }

  function encontrarFaixaSemanaAtual() {
    const candidatos = Array.from(document.querySelectorAll("div"));

    return candidatos
      .filter(el => {
        const texto = el.innerText || "";
        return texto.includes("Semana atual:") && texto.includes("dados exibidos");
      })
      .sort((a, b) => {
        const areaA = a.offsetWidth * a.offsetHeight;
        const areaB = b.offsetWidth * b.offsetHeight;
        return areaA - areaB;
      })[0] || null;
  }

  function numeroSemanaISO(data) {
    const d = new Date(Date.UTC(data.getFullYear(), data.getMonth(), data.getDate()));
    const dia = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dia);
    const anoInicio = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - anoInicio) / 86400000) + 1) / 7);
  }

  function obterSemanaAtual() {
    if (typeof window.getSemanaId === "function") return window.getSemanaId();

    const texto = document.body.innerText || "";
    const match = texto.match(/(\d{4}-W\d{1,2})/i);
    if (match) return match[1];

    const hoje = new Date();
    const semana = numeroSemanaISO(hoje);
    return `${hoje.getFullYear()}-W${String(semana).padStart(2, "0")}`;
  }

  function inserirBotao() {
    limparControlesAntigos();

    if (!estaNoDashboard()) return;

    const existente = document.getElementById(ID_BOX);
    if (existente) return;

    const faixaSemana = encontrarFaixaSemanaAtual();
    if (!faixaSemana) return;

    const box = document.createElement("div");
    box.id = ID_BOX;
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
      max-height: 64px;
      overflow: visible;
      position: static;
      clear: both;
      z-index: 20;
    `;

    box.innerHTML = `
      <div style="display:flex;align-items:center;gap:12px;min-width:0;flex:1;">
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

        <div style="font-size:12px;color:#ffcf8a;line-height:1.3;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
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
    select.value = diaSelecionado;

    select.addEventListener("change", function () {
      diaSelecionado = this.value;
      localStorage.setItem("diaTravaPendentes", diaSelecionado);
    });

    document.getElementById("btnTravarPendentesUnico").addEventListener("click", travarPendentesComoPrincipal);
  }

  async function travarPendentesComoPrincipal() {
    try {
      if (!window.SP) {
        alert("SP não encontrado. Recarregue a página e tente novamente.");
        return;
      }

      const dia = document.getElementById("diaTravaPendentesUnico")?.value || diaSelecionado || "Segunda";
      diaSelecionado = dia;
      localStorage.setItem("diaTravaPendentes", diaSelecionado);

      const semanaId = obterSemanaAtual();

      const ok = confirm(
        `Confirmar trava dos pendentes?\n\n` +
        `Semana: ${semanaId}\n` +
        `Dia: ${dia}\n\n` +
        `Quem não marcou será criado automaticamente como Principal.`
      );

      if (!ok) return;

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
          normalizar(p.Dia) === normalizar(dia)
        );

        if (jaTemPedido) continue;

        await SP.savePedido(semanaId, colaboradorId, c.Nome || c.Title || "", dia, "Principal", "Principal");

        const pedidosColaborador = await SP.getPedidoColaborador(semanaId, colaboradorId);
        const pedidoCriado = pedidosColaborador.find(p => normalizar(p.Dia) === normalizar(dia));

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

      alert(`${criados} pendente(s) preenchido(s) como Principal para ${dia}.`);
      location.reload();
    } catch (erro) {
      console.error("Erro ao travar pendentes:", erro);
      alert(`Erro ao travar pendentes: ${erro.message || erro}`);
    }
  }

  function iniciar() {
    let tentativas = 0;

    const timer = setInterval(() => {
      tentativas++;
      inserirBotao();

      if (document.getElementById(ID_BOX) || tentativas >= 20) {
        clearInterval(timer);
      }
    }, 300);

    document.addEventListener("click", () => {
      setTimeout(() => {
        if (estaNoDashboard() && !document.getElementById(ID_BOX)) inserirBotao();
      }, 400);
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", iniciar);
  } else {
    iniciar();
  }
})();
