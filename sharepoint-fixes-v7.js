// ============================================================
// sharepoint-fixes-v7.js — Correções seguras em runtime para SP
// Refeitório Homy · Integridade v7
//
// Carregar depois de sharepoint.js.
// Objetivo: corrigir regras críticas sem substituir o sharepoint.js inteiro.
// ============================================================
(function aplicarCorrecoesSharePointV7(global) {
  "use strict";

  const SP = global.SP;
  if (!SP) {
    console.error("[Homy v7] SP não encontrado. Carregue sharepoint.js antes de sharepoint-fixes-v7.js.");
    return;
  }

  if (SP._homyIntegridadeV7) {
    console.info("[Homy v7] Correções SP já estavam aplicadas.");
    return;
  }

  const original = {
    criarOuAtualizarPedidoRetornoPrincipal: SP._criarOuAtualizarPedidoRetornoPrincipal,
    normalizarPedidosPorAusenciasSemana: SP._normalizarPedidosPorAusenciasSemana,
    addExtraPedidoCC: SP._addExtraPedidoCC,
    pedidoPreferidoGrupoColaborador: SP._pedidoPreferidoGrupoColaborador
  };

  Object.assign(SP, {
    _homyIntegridadeV7: true,
    _homyIntegridadeV7Original: original,
    _homyExtraLocksV7: Object.create(null),

    _homyV7DataISO(valor) {
      if (!valor) return "";
      if (valor instanceof Date && !isNaN(valor)) return valor.toISOString().slice(0, 10);
      const s = String(valor || "").trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
      const d = new Date(s);
      return d && !isNaN(d) ? d.toISOString().slice(0, 10) : "";
    },

    _homyV7HojeISO() {
      return new Date().toISOString().slice(0, 10);
    },

    _homyV7DiaOperacionalPassado(dataOuDiaInfo) {
      const data = this._homyV7DataISO(dataOuDiaInfo?.data || dataOuDiaInfo);
      if (!data) return false;
      return data < this._homyV7HojeISO();
    },

    _homyV7PodeReceberRetornoAutomatico(diaInfo, options = {}) {
      if (options && options.permitirRetornoRetroativo === true) return true;
      return !this._homyV7DiaOperacionalPassado(diaInfo);
    },

    _homyV7Norm(valor) {
      if (typeof this.norm === "function") return this.norm(valor);
      return String(valor || "")
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .toLowerCase().trim();
    },

    _homyV7StatusBloqueiaProducao(pedido) {
      const s = this._homyV7Norm(this.pick?.(pedido, "Status", "status") ?? pedido?.Status ?? pedido?.status);
      return [
        "cancelado", "bloqueado", "nao vai almocar", "nao_vai_almocar",
        "ausente", "ferias", "afastado", "atestado", "licenca",
        "banco horas", "banco_horas", "homy office", "homy_office", "falta"
      ].includes(s);
    },

    _homyV7PedidoProdutivo(pedido) {
      const status = this._homyV7Norm(this.pick?.(pedido, "Status", "status") ?? pedido?.Status ?? pedido?.status);
      const origem = this._homyV7Norm(this.pick?.(pedido, "Origem", "origem", "tipo", "Tipo") ?? pedido?.Origem ?? pedido?.origem);
      if (status === "travado" && origem.includes("travamento")) return true;
      if (this._homyV7StatusBloqueiaProducao(pedido)) return false;
      const confirmado = this.pick?.(pedido, "Confirmado", "confirmado") ?? pedido?.Confirmado ?? pedido?.confirmado;
      return status === "confirmado" || status === "aprovado" || status === "extra" || this.isTrue?.(confirmado) === true;
    },

    _homyV7ExtraKey(semanaId, dia, nome, tipo, opcao) {
      return [semanaId, dia, nome, tipo, opcao || "principal"]
        .map(v => this._homyV7Norm(v))
        .join("|");
    }
  });

  // 1) Retorno automático não pode retroagir.
  if (typeof original.criarOuAtualizarPedidoRetornoPrincipal === "function") {
    SP._criarOuAtualizarPedidoRetornoPrincipal = async function(semanaId, diaInfo, ausencia, colaborador, pedidos, options = {}) {
      if (!this._homyV7PodeReceberRetornoAutomatico(diaInfo, options)) {
        return { criado: 0, atualizado: 0, ignorado: 1, motivo: "retorno automático bloqueado para dia operacional passado" };
      }
      return original.criarOuAtualizarPedidoRetornoPrincipal.call(this, semanaId, diaInfo, ausencia, colaborador, pedidos, options);
    };
  }

  // 1b) A normalização antiga também podia converter ausência antiga em Principal.
  // Para preservar histórico, ela passa a atuar somente hoje/futuro.
  if (typeof original.normalizarPedidosPorAusenciasSemana === "function") {
    SP._normalizarPedidosPorAusenciasSemana = async function(semanaId, pedidos, ausencias, dias, colabPorKey, options = {}) {
      if (options && options.permitirRetornoRetroativo === true) {
        return original.normalizarPedidosPorAusenciasSemana.call(this, semanaId, pedidos, ausencias, dias, colabPorKey, options);
      }
      const diasSeguros = (dias || []).filter(d => !this._homyV7DiaOperacionalPassado(d));
      if (!diasSeguros.length) return { pedidosAtualizados: 0, pedidosCancelados: 0, retornosIgnoradosPorDiaPassado: (dias || []).length };
      return original.normalizarPedidosPorAusenciasSemana.call(this, semanaId, pedidos, ausencias, diasSeguros, colabPorKey, options);
    };
  }

  // 2) Extra idempotente: evita criação paralela/duplicada de Guarda, Investigador, Prestador e Refeição Extra.
  if (typeof original.addExtraPedidoCC === "function") {
    SP._addExtraPedidoCC = async function(semanaId, dia, nome, tipo, opcao = "principal", observacao = "", centroCusto = "", adicionadoPor = "") {
      const key = this._homyV7ExtraKey(semanaId, dia, nome, tipo, opcao);
      if (this._homyExtraLocksV7[key]) return this._homyExtraLocksV7[key];

      const run = (async () => {
        try {
          this.clearListCache?.("Pedidos");
          this.clearListCache?.("Extras");

          const pedidos = await this.getPedidos?.(semanaId, { force: true }).catch(() => []) || [];
          const existenteAtivo = (pedidos || []).find(p => {
            const bate = typeof this._pedidoExtraBate === "function"
              ? this._pedidoExtraBate(p, semanaId, dia, nome, tipo, opcao, "")
              : false;
            return bate && this._homyV7PedidoProdutivo(p);
          });

          if (existenteAtivo) {
            return { extra: null, pedido: existenteAtivo, criadoPedido: false, reaproveitado: true };
          }

          return original.addExtraPedidoCC.call(this, semanaId, dia, nome, tipo, opcao, observacao, centroCusto, adicionadoPor);
        } finally {
          delete this._homyExtraLocksV7[key];
        }
      })();

      this._homyExtraLocksV7[key] = run;
      return run;
    };
  }

  // 3) Duplicidade: pedido produtivo válido vence ausência/cancelado.
  if (typeof original.pedidoPreferidoGrupoColaborador === "function") {
    SP._pedidoPreferidoGrupoColaborador = function(grupo) {
      const lista = [...(grupo || [])];
      if (!lista.length) return null;
      return lista.sort((a, b) => {
        const ap = this._homyV7PedidoProdutivo(a) ? 1 : 0;
        const bp = this._homyV7PedidoProdutivo(b) ? 1 : 0;
        if (ap !== bp) return bp - ap;

        const ac = this._homyV7StatusBloqueiaProducao(a) ? 1 : 0;
        const bc = this._homyV7StatusBloqueiaProducao(b) ? 1 : 0;
        if (ac !== bc) return ac - bc;

        const ta = new Date(this.pick?.(a, "Modified", "modified", "Data_Hora", "Created", "created") || 0).getTime() || Number(this.pick?.(a, "id", "ID") || 0) || 0;
        const tb = new Date(this.pick?.(b, "Modified", "modified", "Data_Hora", "Created", "created") || 0).getTime() || Number(this.pick?.(b, "id", "ID") || 0) || 0;
        return tb - ta;
      })[0];
    };
  }

  console.info("[Homy v7] Correções SP aplicadas: retorno sem retroagir, extra idempotente e duplicidade segura.");
})(window);
