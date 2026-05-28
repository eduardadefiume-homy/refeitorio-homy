// ============================================================
// admin-operacao-dia.js — Operação diária da Luana/cozinha
// Regra: travar pendentes como Principal
// ============================================================

window.AdminOperacaoDia = window.AdminOperacaoDia || {};

Object.assign(window.AdminOperacaoDia, {
  async travarPendentesComoPrincipal(semanaId, dia) {
    if (!window.SP) throw new Error("SP não encontrado.");

    const colaboradores = await (SP.getTodosColaboradores ? SP.getTodosColaboradores() : SP.getColaboradores());
    const pedidosSemana = await SP.getPedidos(semanaId);

    const ativos = colaboradores.filter(c => SP.isTrue ? SP.isTrue(c.Ativo) : String(c.Ativo).toLowerCase() !== "false");

    let criados = 0;

    for (const c of ativos) {
      const jaTemPedido = pedidosSemana.some(p =>
        String(p.Colaborador_id) === String(c.id) &&
        String(p.Dia || "").toLowerCase() === String(dia || "").toLowerCase()
      );

      if (jaTemPedido) continue;

      await SP.savePedido(
        semanaId,
        c.id,
        c.Nome || c.Title || "",
        dia,
        "Principal",
        "Principal"
      );

      const pedidosAtualizados = await SP.getPedidoColaborador(semanaId, c.id);
      const pedidoCriado = pedidosAtualizados.find(p =>
        String(p.Dia || "").toLowerCase() === String(dia || "").toLowerCase()
      );

      if (pedidoCriado) {
        await SP.updatePedido(pedidoCriado.id, {
          Centro_Custo: c.Centro_Custo || "",
          Status: "Confirmado",
          Origem: "Admin",
          Observacao: "Preenchido automaticamente após prazo de marcação.",
          Alterado_Por: SP.getUserEmail ? SP.getUserEmail() : ""
        });
      }

      criados++;
    }

    alert(`${criados} pendente(s) preenchido(s) automaticamente como Principal.`);
    return criados;
  }
});
