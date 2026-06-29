// ============================================================
// admin-pos-ausencia.js — Regra pós-ausência · Operação do Dia
// v: base-limpa-pos-ausencia-v10-4-20260629
//
// Carregar depois de admin-operacao-dia.js.
// Objetivo: impedir que pedido antigo de ausência encerrada vire Principal
// na tela de Operação do Dia antes do travamento/fechamento.
// ============================================================
(function (global) {
  "use strict";

  function instalar() {
    const AdminOperacao = global.AdminOperacao;
    if (!AdminOperacao || AdminOperacao.__posAusenciaV10) return false;

    if (typeof AdminOperacao._normalizarPedidosObsoletosOperacao !== "function") return false;

    AdminOperacao.__posAusenciaV10 = true;
    AdminOperacao.__posAusenciaOriginalNormalizar = AdminOperacao.__posAusenciaOriginalNormalizar || AdminOperacao._normalizarPedidosObsoletosOperacao;
    AdminOperacao._normalizarPedidosObsoletosOperacao = function _normalizarPedidosObsoletosOperacaoV10(lista, ausencias, semanaId, dia) {
      const grupos = new Map();
      const especiais = [];

      for (const p of (lista || [])) {
        if (this._isExtraPedido(p) || this._isRefeicaoAdicionalColaborador(p)) {
          especiais.push(p);
          continue;
        }
        const key = this._pedidoKeyOperacao(p);
        if (!grupos.has(key)) grupos.set(key, []);
        grupos.get(key).push(p);
      }

      const saida = [...especiais];

      for (const grupo of grupos.values()) {
        const base = grupo[0];
        const aus = this._ausenciaDoPedidoNoDia(base, ausencias, semanaId, dia);

        if (aus) {
          const escolhido = grupo.find(p => this._isAusenteOperacao(p)) || grupo.sort((a, b) => this._compararPedidoOperacao(b, a))[0];
          saida.push({ ...escolhido, _ausenciaOperacao: aus });
          continue;
        }

        const normais = grupo.filter(p =>
          !this._pedidoAusenciaObsoleta(p, ausencias, semanaId, dia) &&
          !["cancelado", "bloqueado"].includes(this._norm(this._pick(p, "Status", "status")))
        );
        if (normais.length) {
          normais.sort((a, b) => this._compararPedidoOperacao(b, a));
          saida.push(normais[0]);
          continue;
        }

        // v10: se só restou pedido antigo de férias/ausência e ela não cobre mais o dia,
        // NÃO converter para Principal aqui. Mostra como pendente para a Operação.
        const stale = grupo.sort((a, b) => this._compararPedidoOperacao(b, a))[0];
        saida.push({
          ...stale,
          _virtualPendente: true,
          _ausenciaEncerradaAguardandoMarcacao: true,
          Status: "Pendente",
          Confirmado: false,
          Nome_Prato: "Sem marcação",
          Origem: "Ausência encerrada - aguardando marcação",
          Observacao: "Ausência encerrada. Colaborador liberado para escolher; Principal só será aplicado no travamento/fechamento se ninguém marcar."
        });
      }

      return this._deduplicarPedidosOperacao(saida);
    };

    console.info("[Homy v10] Regra pós-ausência aplicada na Operação do Dia.");
    return true;
  }

  function boot() {
    if (instalar()) return;
    let tentativas = 0;
    const t = setInterval(() => {
      if (instalar() || ++tentativas > 30) clearInterval(t);
    }, 250);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();
})(window);
