// ============================================================
// admin-dashboard-pendentes.js
// Botão único e compacto: Travar pendentes como Principal
// ============================================================

(function () {
  const ID_BOX = "controleTravaPendentesUnico";

  function normalizar(valor) {
    return String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  }

  function removerDuplicados() {
    ["controleTravaPendentes", "controleTravaPendentesForcado"].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.remove();
    });

    document.querySelectorAll("button").forEach(btn => {
      const texto = normalizar(btn.innerText || "");
      const paiAntigo = btn.closest("[id^='controleTravaPendentes']");
      if (texto.includes("travar pendentes") && paiAntigo && paiAntigo.id !== ID_BOX) {
        paiAntigo.remove();
      }
    });
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

  function estaNoDashboard() {
    const titulo = document.querySelector(".topbar-title, h1, h2");
    const textoTitulo = normalizar(titulo?.innerText || "");
    if (textoTitulo.includes("dashboard")) return true;

    const ativo = document.querySelector(".nav-item.active, .module.active");
    const textoAtivo = normalizar(ativo?.innerText || "");
    return textoAtivo.includes("dashboard");
  }

  function encontrarReferenciaSemana() {
    return Array.from(document.querySelectorAll("div,span,p")).find(el =>
      (el.innerText || "").includes("Semana atual")
    );
  }

  function inserirBotao() {
    removerDuplicados();

    if (!estaNoDashboard()) return;
    if (document.getElementById(ID_BOX)) return;

    const referencia = encontrarReferenciaSemana();
    const content = document.querySelector(".content") || document.querySelector("main") || document.body;

    const box = document.createElement("div");
    box.id = ID_BOX;
    box.style.cssText = `
      margin: .8rem 0 1rem 0;
      padding: .8rem 1rem;
      border: 1px solid rgba(192,40,28,.35);
      border-radius: 10px;
      background: rgba(192,40,28,.08);
      display: flex;
      align-items: center;
      gap: .8rem;
      justify-content: space-between;
      max-width: 100%;
    `;

    box.innerHTML = `
      <div style="display:flex;align-items:center;gap:.7rem;min-width:0;">
        <div class="form-group" style="margin:0;min-width:150px;">
          <label class="form-label" style="margin-bottom:.25rem;">DIA</label>
          <select id="diaTravaPendentesUnico" class="form-select" style="height:38px;">
            <option value="Segunda">Segunda</option>
            <option value="Terça">Terça</option>
            <option value="Quarta">Quarta</option>
            <option value="Quinta">Quinta</option>
            <option value="Sexta">Sexta</option>
          </select>
        </div>

        <div style="font-size:.76rem;color:#ffcf8a;line-height:1.35;white-space:normal;">
          Após o prazo, preenche automaticamente como <b>Principal</b> quem ficou pendente.
        </div>
      </div>

      <button type="button" class="btn-danger" id="btnTravarPendentesUnico" style="white-space:nowrap;height:38px;padding:0 .9rem;">
        🔒 Travar pendentes
      </button>
    `;

    if (referencia) referencia.insertAdjacentElement("afterend", box);
    else content.appendChild(box);

    document.getElementById("btnTravarPendentesUnico").addEventListener("click", travarPendentesComoPrincipal);
  }

  async function travarPendentesComoPrincipal() {
    try {
      if (!window.SP) {
        alert("SP não encontrado. Recarregue a página e tente novamente.");
        return;
      }

      const dia = document.getElementById("diaTravaPendentesUnico")?.value || "Segunda";
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
    inserirBotao();

    const observer = new MutationObserver(() => inserirBotao());
    observer.observe(document.body, { childList: true, subtree: true });

    document.addEventListener("click", () => setTimeout(inserirBotao, 250));
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", iniciar);
  else iniciar();
})();
