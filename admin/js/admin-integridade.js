// ============================================================
// admin-integridade.js — Integridade dos Dados · Admin Homy
// v: base-centralizada-integridade-v10-17-20260701
//
// Objetivo:
// - Ler dados reais do SharePoint via SP/sharepoint.js.
// - Comparar produção calculada x referência operacional quando existir.
// - Identificar causas prováveis das diferenças.
// - Gerar plano de correção SOMENTE LEITURA.
// - NÃO grava no SharePoint nesta fase.
// ============================================================
(function (global) {
  "use strict";

  const AdminIntegridade = global.AdminIntegridade = {
    _ultimoResultado: null,
    _loading: false,
    _instalado: false,
    _modalAberto: false,
    _cache: new Map(),
    _config: {
      containerId: "dashAlertas",
      checarSomenteAteHoje: true,
      maxItensResumo: 6,
      maxDetalhesPorGrupo: 80,
      cacheMs: 15000,
      referenciaOperacional: {
        "2026-W26": {
          fonte: "Luana",
          valores: {
            segunda: { principal: 44, light: 4, carne: 17, massa: 0, lanche: 0, total: 65 },
            terca:   { principal: 49, light: 6, carne: 8,  massa: 6, lanche: 0, total: 69 },
            quarta:  { principal: 45, light: 3, carne: 17, massa: 5, lanche: 0, total: 70 },
            quinta:  { principal: 57, light: 5, carne: 4,  massa: 4, lanche: 0, total: 70 }
          }
        }
      }
    },

    // ============================================================
    // API PÚBLICA
    // ============================================================
    async load(semanaId, options = {}) {
      const cfg = { ...this._config, ...(options || {}) };
      const semana = semanaId || this._semanaAtual();
      const container = document.getElementById(cfg.containerId);
      if (!container || !semana) return null;
      if (this._loading) return this._ultimoResultado;

      this._loading = true;
      try {
        this._ensureStyle();
        this._ensureModal();

        const cacheKey = `${semana}|${cfg.checarSomenteAteHoje ? "ate-hoje" : "semana-inteira"}`;
        const cached = this._cache.get(cacheKey);
        if (cached && Date.now() - cached.ts < cfg.cacheMs) {
          this._ultimoResultado = cached.resultado;
          this.renderResumo(cached.resultado, cfg);
          if (this._modalAberto) this._atualizarModal(cached.resultado);
          return cached.resultado;
        }

        const resultado = await this.verificarSemana(semana, cfg);
        this._ultimoResultado = resultado;
        this._cache.set(cacheKey, { ts: Date.now(), resultado });
        this.renderResumo(resultado, cfg);
        if (this._modalAberto) this._atualizarModal(resultado);
        return resultado;
      } catch (e) {
        console.error("[AdminIntegridade]", e);
        this.renderErro(e, cfg);
        return null;
      } finally {
        this._loading = false;
      }
    },

    async verificarSemana(semanaId, options = {}) {
      await this._ensureSP();
      const dados = await this._buscarDados(semanaId);
      const dias = this._diasParaVerificar(semanaId, options);
      const referenciaOperacional = this._referenciaSemana(semanaId, options);

      const detalhesProducaoPorDia = {};
      const resumoProducaoPorDia = {};
      const resumoRetiradaPorDia = {};
      const comparativoPorDia = {};
      const irregularidades = [];

      for (const diaInfo of dias) {
        const det = this._calcularDetalhesProducaoDia(dados.pedidos, diaInfo.dia);
        detalhesProducaoPorDia[diaInfo.dia] = det;
        resumoProducaoPorDia[diaInfo.dia] = det.resumo;
        resumoRetiradaPorDia[diaInfo.dia] = this._calcularResumoRetiradaDia(dados.checkins, diaInfo.dia);
      }

      irregularidades.push(...this._verificarPedidosQuebrados(dados, semanaId));
      irregularidades.push(...this._verificarPedidosBloqueadosConfirmados(dados, semanaId));
      irregularidades.push(...this._verificarDuplicidadesNormais(dados, semanaId));
      irregularidades.push(...this._verificarDuplicidadesExtras(dados, semanaId));

      for (const diaInfo of dias) {
        irregularidades.push(...this._verificarAusenciasVigentesComPedidoConfirmado(dados, semanaId, diaInfo));
        irregularidades.push(...this._verificarRetornosEsperados(dados, semanaId, diaInfo));
        const check = resumoRetiradaPorDia[diaInfo.dia];
        const prod = resumoProducaoPorDia[diaInfo.dia];
        if (prod.total > 0 && check.total === 0 && this._dataPassadaOuHoje(diaInfo.data)) {
          irregularidades.push(this._issue({
            id: `checkin-sem-dados-dia-${diaInfo.dia}`,
            tipo: "checkin-sem-dados-dia",
            severidade: "info",
            nome: this._diaLabel(diaInfo.dia),
            semanaId,
            dia: diaInfo.dia,
            mensagem: `A lista CheckIn retornou 0 retirada(s) para ${this._diaLabel(diaInfo.dia)}. Não usar CheckIn como conferência histórica desse dia sem validar os campos.`,
            acaoSugerida: "Validar se a Cozinha gravou Semana_id, Dia, Retirou e Data_Hora_Retirada corretamente."
          }));
        }
      }

      for (const diaInfo of dias) {
        comparativoPorDia[diaInfo.dia] = this._compararDiaComReferencia(
          diaInfo.dia,
          resumoProducaoPorDia[diaInfo.dia],
          resumoRetiradaPorDia[diaInfo.dia],
          referenciaOperacional?.valores?.[diaInfo.dia] || null
        );
        if (comparativoPorDia[diaInfo.dia].temReferencia && !comparativoPorDia[diaInfo.dia].bateProducaoComReferencia) {
          const dif = comparativoPorDia[diaInfo.dia].diferencas;
          const partes = Object.entries(dif).filter(([_, v]) => v !== 0).map(([k, v]) => `${k} ${v > 0 ? "+" : ""}${v}`);
          irregularidades.push(this._issue({
            id: `diferenca-referencia-${diaInfo.dia}`,
            tipo: "diferenca-referencia-operacional",
            severidade: Math.abs(dif.total || 0) > 1 ? "critico" : "alerta",
            nome: this._diaLabel(diaInfo.dia),
            semanaId,
            dia: diaInfo.dia,
            mensagem: `Produção calculada não bate com referência Luana. Diferenças: ${partes.join(", ")}.`,
            acaoSugerida: "Abrir o Plano de Correção e revisar os registros sugeridos antes de qualquer gravação."
          }));
        }
      }

      const planoCorrecao = this._gerarPlanoCorrecao({
        semanaId,
        dias,
        dados,
        detalhesProducaoPorDia,
        comparativoPorDia,
        referenciaOperacional
      });

      irregularidades.push(...planoCorrecao.irregularidadesInfo);
      irregularidades.sort((a, b) => {
        const peso = { critico: 0, alerta: 1, atencao: 2, info: 3 };
        return (peso[a.severidade] ?? 9) - (peso[b.severidade] ?? 9) ||
          String(a.nome || "").localeCompare(String(b.nome || ""), "pt-BR", { sensitivity: "base" });
      });

      const porSeveridade = irregularidades.reduce((acc, item) => {
        acc[item.severidade] = (acc[item.severidade] || 0) + 1;
        return acc;
      }, {});

      return {
        semanaId,
        geradoEm: new Date().toISOString(),
        observacao: "Prévia de integridade. Nenhuma correção é aplicada sem confirmação explícita.",
        referenciaOperacional,
        diasVerificados: dias,
        dadosResumo: {
          pedidos: dados.pedidos.length,
          ausencias: dados.ausencias.length,
          colaboradores: dados.colaboradores.length,
          extras: dados.extras.length,
          cardapio: dados.cardapio.length,
          checkins: dados.checkins.length
        },
        resumoProducaoPorDia,
        resumoRetiradaPorDia,
        comparativoPorDia,
        detalhesProducaoPorDia,
        planoCorrecao,
        irregularidades,
        porSeveridade,
        status: irregularidades.some(i => i.severidade === "critico") ? "critico" :
          irregularidades.some(i => i.severidade === "alerta" || i.severidade === "atencao") ? "atencao" : "ok"
      };
    },

    renderResumo(resultado, options = {}) {
      const container = document.getElementById(options.containerId || this._config.containerId);
      if (!container || !resultado) return;

      let box = document.getElementById("adminIntegridadeResumo");
      if (!box) {
        box = document.createElement("div");
        box.id = "adminIntegridadeResumo";
        container.appendChild(box);
      }

      const total = resultado.irregularidades.length;
      const criticos = resultado.porSeveridade.critico || 0;
      const alertas = resultado.porSeveridade.alerta || 0;
      const atencoes = resultado.porSeveridade.atencao || 0;
      const infos = resultado.porSeveridade.info || 0;
      const plano = resultado.planoCorrecao || { totais: {} };
      const classe = resultado.status === "ok" ? "ok" : resultado.status === "critico" ? "critico" : "atencao";
      const icone = resultado.status === "ok" ? "✅" : resultado.status === "critico" ? "🔴" : "⚠️";

      const difs = Object.values(resultado.comparativoPorDia || {})
        .filter(d => d.temReferencia && !d.bateProducaoComReferencia)
        .map(d => `${this._diaLabel(d.dia)} ${d.diferencas.total > 0 ? "+" : ""}${d.diferencas.total}`);

      const amostra = resultado.irregularidades.slice(0, options.maxItensResumo || this._config.maxItensResumo);
      const lista = amostra.length ? `
        <div class="integridade-preview-list">
          ${amostra.map(i => `
            <div class="integridade-preview-row ${this._esc(i.severidade)}">
              <div class="integridade-preview-main">${this._esc(i.nome || i.titulo || "Registro")}</div>
              <div class="integridade-preview-sub">${this._esc(i.mensagem || i.tipo || "Irregularidade")}</div>
            </div>
          `).join("")}
        </div>` : "";

      box.innerHTML = `
        <div class="admin-integridade-card ${classe}">
          <div class="admin-integridade-head">
            <div>
              <div class="admin-integridade-title">${icone} Integridade dos Dados</div>
              <div class="admin-integridade-sub">Semana ${this._esc(resultado.semanaId)} · ${this._fmtDataHora(resultado.geradoEm)} · somente leitura</div>
            </div>
            <div class="admin-integridade-head-actions">
              <button class="btn-secondary admin-integridade-btn" onclick="AdminIntegridade.abrirDetalhes()">Abrir detalhes</button>
              <button class="btn-secondary admin-integridade-btn" onclick="AdminIntegridade.abrirCorrecaoAssistida()">Aplicar correção assistida</button>
              <button class="btn-secondary admin-integridade-btn" onclick="AdminIntegridade.baixarDiagnostico()">Baixar JSON</button>
            </div>
          </div>

          <div class="admin-integridade-status">${this._esc(total ? `${total} irregularidade(s) encontradas.` : "Integridade dos dados sem irregularidades críticas.")}</div>
          ${difs.length ? `<div class="integridade-dif-alert">Diferença contra referência Luana: ${this._esc(difs.join(" · "))}</div>` : ""}
          <div class="integridade-plano-mini">
            🛠 Plano sugerido: ${this._num(plano.totais?.cancelar)} cancelar · ${this._num(plano.totais?.reativar)} reativar · ${this._num(plano.totais?.revisar)} revisar
          </div>

          <div class="admin-integridade-pills">
            <span class="admin-integridade-pill critico">🔴 ${criticos} crítico(s)</span>
            <span class="admin-integridade-pill alerta">🟠 ${alertas} alerta(s)</span>
            <span class="admin-integridade-pill atencao">🟡 ${atencoes} atenção</span>
            <span class="admin-integridade-pill info">🔵 ${infos} info</span>
          </div>
          ${lista}
        </div>`;
    },

    renderErro(e, options = {}) {
      const container = document.getElementById(options.containerId || this._config.containerId);
      if (!container) return;
      let box = document.getElementById("adminIntegridadeResumo");
      if (!box) {
        box = document.createElement("div");
        box.id = "adminIntegridadeResumo";
        container.appendChild(box);
      }
      box.innerHTML = `
        <div class="admin-integridade-card critico">
          <div class="admin-integridade-title">🔴 Integridade dos Dados</div>
          <div class="admin-integridade-status">Não foi possível verificar a integridade agora.</div>
          <div class="admin-integridade-error">${this._esc(e.message || e)}</div>
        </div>`;
    },

    abrirDetalhes() {
      this._ensureModal();
      const overlay = document.getElementById("adminIntegridadeModalOverlay");
      if (!overlay) return;
      this._modalAberto = true;
      this._atualizarModal(this._ultimoResultado);
      overlay.classList.add("open");
      document.body.classList.add("admin-integridade-modal-open");
      const scroller = document.getElementById("adminIntegridadeModalScroller");
      if (scroller) scroller.scrollTop = 0;
    },

    fecharDetalhes() {
      const overlay = document.getElementById("adminIntegridadeModalOverlay");
      if (overlay) overlay.classList.remove("open");
      document.body.classList.remove("admin-integridade-modal-open");
      this._modalAberto = false;
    },

    copiarDiagnostico() {
      const r = this._ultimoResultado;
      if (!r) return;
      const texto = JSON.stringify(r, null, 2);
      navigator.clipboard?.writeText(texto)
        .then(() => this._toast("Diagnóstico copiado.", "success"))
        .catch(() => {
          console.log("Diagnóstico de integridade:", r);
          this._toast("Não foi possível copiar; diagnóstico enviado ao console.", "warning");
        });
    },

    baixarDiagnostico() {
      const r = this._ultimoResultado;
      if (!r) return;
      const blob = new Blob([JSON.stringify(r, null, 2)], { type: "application/json;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `integridade-${r.semanaId}-${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 800);
    },


    async abrirCorrecaoAssistida() {
      try {
        if (global.AdminCorrecaoIntegridade?.abrirSemana) {
          await global.AdminCorrecaoIntegridade.abrirSemana();
          return;
        }
        this._toast("Módulo de Correção Assistida ainda não carregado. Recarregue o Admin e tente novamente.", "warning");
      } catch (e) {
        console.error("[AdminIntegridade] Correção Assistida", e);
        this._toast("Erro ao abrir correção assistida: " + (e.message || e), "error");
      }
    },

    baixarPlanoCorrecao() {
      const r = this._ultimoResultado;
      if (!r) return;
      const blob = new Blob([JSON.stringify(r.planoCorrecao || {}, null, 2)], { type: "application/json;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `plano-correcao-${r.semanaId}-${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 800);
    },

    instalarNoDashboard() {
      if (this._instalado) return true;
      if (!global.AdminDashboard || typeof global.AdminDashboard._renderAlertas !== "function") return false;

      const original = global.AdminDashboard._renderAlertas;
      if (original.__adminIntegridadeWrapped) {
        this._instalado = true;
        return true;
      }

      const self = this;
      const wrapped = async function _renderAlertasComIntegridade(resumo) {
        const retorno = original.call(this, resumo);
        if (retorno && typeof retorno.then === "function") await retorno;
        const semanaId = global.AdminState?.getSemanaId?.() || resumo?.semanaId || self._semanaAtual();
        await self.load(semanaId).catch(e => console.warn("[AdminIntegridade] render ignorado:", e));
      };
      wrapped.__adminIntegridadeWrapped = true;
      global.AdminDashboard._renderAlertas = wrapped;
      this._instalado = true;
      return true;
    },

    // ============================================================
    // BUSCA DE DADOS
    // ============================================================
    async _ensureSP() {
      if (!global.SP) throw new Error("SP/sharepoint.js não está carregado.");
      if (typeof SP.init === "function") await SP.init();
    },

    async _buscarDados(semanaId) {
      const [pedidos, ausencias, colaboradores, extras, cardapio, checkins] = await Promise.all([
        this._safeList(() => SP.getPedidos?.(semanaId, { reparar: false, force: true }), "Pedidos"),
        this._safeList(() => SP.getAusencias?.(false), "Ausencias do Refeitorio"),
        this._safeList(() => SP.getTodosColaboradores?.(), "Colaboradores"),
        this._safeList(() => SP.getExtras?.(semanaId), "Extras"),
        this._safeList(() => SP.getCardapio?.(semanaId), "Cardapio"),
        this._safeList(() => SP.getItems?.("CheckIn", { force: true }), "CheckIn")
      ]);

      return {
        pedidos: (pedidos || []).filter(p => String(this._pick(p, "Semana_id", "Semana") || "") === String(semanaId)),
        ausencias: ausencias || [],
        colaboradores: (colaboradores || []).filter(c => this._colaboradorAtivo(c)),
        extras: extras || [],
        cardapio: cardapio || [],
        checkins: (checkins || []).filter(c => String(this._pick(c, "Semana_id", "Semana") || "") === String(semanaId))
      };
    },

    async _safeList(fn, nomeLista) {
      try {
        if (typeof fn !== "function") return [];
        const r = await fn();
        return Array.isArray(r) ? r : [];
      } catch (e) {
        console.warn(`[AdminIntegridade] Falha ao ler ${nomeLista}:`, e);
        return [];
      }
    },

    // ============================================================
    // PLANO DE CORREÇÃO — SOMENTE LEITURA
    // ============================================================
    _gerarPlanoCorrecao(ctx) {
      const acoes = [];
      const irregularidadesInfo = [];
      const usados = new Set();
      const opcoes = ["principal", "light", "carne", "massa", "lanche"];

      // 1) Extras duplicados: mesma semana + dia + nome + origem/tipo + opção.
      const gruposExtras = this._gruparExtrasProducao(ctx.dados.pedidos);
      for (const grupo of gruposExtras.values()) {
        if (grupo.length < 2) continue;
        const ordenados = grupo.slice().sort((a, b) => this._timestamp(a) - this._timestamp(b));
        const manter = ordenados[0];
        for (const p of ordenados.slice(1)) {
          const id = this._id(p);
          if (!id || usados.has(id)) continue;
          usados.add(id);
          acoes.push(this._acaoCancelar(p, "extra-duplicado", `Extra duplicado: manter ID ${this._id(manter)} e cancelar este registro.`));
        }
      }

      // 2) Diferença contra referência: primeiro resolver déficits por reativação de pedido cancelado compatível.
      for (const diaInfo of ctx.dias) {
        const comp = ctx.comparativoPorDia[diaInfo.dia];
        if (!comp?.temReferencia || comp.bateProducaoComReferencia) continue;
        const detalhes = ctx.detalhesProducaoPorDia[diaInfo.dia];
        const excluidos = detalhes?.excluidos || [];

        for (const opcao of opcoes) {
          const diff = comp.diferencas[opcao] || 0;
          if (diff >= 0) continue;
          const quantidade = Math.abs(diff);
          const candidatos = excluidos
            .filter(p => this._norm(this._opcao(p)) === opcao)
            .filter(p => this._statusCanceladoOuDuplicado(p))
            .sort((a, b) => this._timestamp(b) - this._timestamp(a));

          for (const p of candidatos.slice(0, quantidade)) {
            const id = this._id(p);
            if (!id || usados.has(id)) continue;
            usados.add(id);
            acoes.push(this._acaoReativar(p, "deficit-opcao", `Referência Luana indica falta de ${quantidade} ${opcao} em ${this._diaLabel(diaInfo.dia)}.`));
          }
        }
      }

      // 3) Diferenças positivas: cancelar primeiro retorno automático retroativo e extras duplicados já mapeados.
      for (const diaInfo of ctx.dias) {
        const comp = ctx.comparativoPorDia[diaInfo.dia];
        if (!comp?.temReferencia || comp.bateProducaoComReferencia) continue;
        const detalhes = ctx.detalhesProducaoPorDia[diaInfo.dia];
        const incluidos = detalhes?.incluidos || [];

        for (const opcao of opcoes) {
          let sobra = comp.diferencas[opcao] || 0;
          if (sobra <= 0) continue;

          const candidatosRetroativos = incluidos
            .filter(p => this._norm(this._opcao(p)) === opcao)
            .filter(p => this._retornoAutomaticoRetroativo(p, diaInfo.data))
            .sort((a, b) => this._timestamp(b) - this._timestamp(a));

          for (const p of candidatosRetroativos) {
            if (sobra <= 0) break;
            const id = this._id(p);
            if (!id || usados.has(id)) continue;
            usados.add(id);
            acoes.push(this._acaoCancelar(p, "retorno-automatico-retroativo", `Retorno automático criado/alterado após o dia operacional ${this._br(diaInfo.data)}. Não deve retroagir.`));
            sobra--;
          }

          const candidatosDuplicados = incluidos
            .filter(p => this._norm(this._opcao(p)) === opcao)
            .filter(p => this._isExtraPedido(p))
            .filter(p => this._temOutroExtraIgualMaisAntigo(p, incluidos))
            .sort((a, b) => this._timestamp(b) - this._timestamp(a));

          for (const p of candidatosDuplicados) {
            if (sobra <= 0) break;
            const id = this._id(p);
            if (!id || usados.has(id)) continue;
            usados.add(id);
            acoes.push(this._acaoCancelar(p, "extra-duplicado", "Extra especial duplicado no mesmo dia/opção."));
            sobra--;
          }

          if (sobra > 0) {
            irregularidadesInfo.push(this._issue({
              id: `plano-revisar-${diaInfo.dia}-${opcao}`,
              tipo: "plano-revisao-manual",
              severidade: "atencao",
              nome: `${this._diaLabel(diaInfo.dia)} · ${opcao}`,
              semanaId: ctx.semanaId,
              dia: diaInfo.dia,
              mensagem: `Ainda sobra(m) ${sobra} pedido(s) em ${opcao} após as sugestões automáticas seguras.`,
              acaoSugerida: "Comparar a lista de incluídos do dia com a conferência da Luana antes de cancelar qualquer outro pedido."
            }));
            acoes.push({
              tipo: "revisar",
              motivo: "sobra-sem-candidato-seguro",
              semanaId: ctx.semanaId,
              dia: diaInfo.dia,
              opcao,
              quantidade: sobra,
              mensagem: `Revisar manualmente ${sobra} pedido(s) de ${opcao} em ${this._diaLabel(diaInfo.dia)}.`
            });
          }
        }
      }

      const simulado = this._simularPlano(ctx.resumoProducaoPorDia || {}, acoes);
      const porDia = this._agruparAcoesPorDia(acoes);
      const totais = {
        cancelar: acoes.filter(a => a.acao === "cancelar").length,
        reativar: acoes.filter(a => a.acao === "reativar").length,
        revisar: acoes.filter(a => a.acao === "revisar").length,
        total: acoes.length
      };

      return {
        status: "somente-leitura",
        mensagem: "Plano sugerido. Nenhuma correção automática foi executada.",
        totais,
        acoes,
        porDia,
        resultadoSimuladoPorDia: simulado,
        irregularidadesInfo
      };
    },

    _acaoCancelar(p, motivo, justificativa) {
      return {
        acao: "cancelar",
        motivo,
        pedidoId: this._id(p),
        nome: this._nome(p),
        dia: this._dia(p),
        opcao: this._norm(this._opcao(p) || "principal"),
        statusAtual: this._status(p),
        confirmadoAtual: this._isTrue(this._pick(p, "Confirmado", "confirmado")),
        origemAtual: this._origem(p),
        justificativa,
        camposSugeridos: {
          Status: "Cancelado",
          Confirmado: false,
          Origem: "Correção de integridade",
          Observacao: `Correção de integridade: ${justificativa}`
        },
        pedido: this._resumoPedido(p)
      };
    },

    _acaoReativar(p, motivo, justificativa) {
      const opcao = this._norm(this._opcao(p) || "principal");
      return {
        acao: "reativar",
        motivo,
        pedidoId: this._id(p),
        nome: this._nome(p),
        dia: this._dia(p),
        opcao,
        statusAtual: this._status(p),
        confirmadoAtual: this._isTrue(this._pick(p, "Confirmado", "confirmado")),
        origemAtual: this._origem(p),
        justificativa,
        camposSugeridos: {
          Status: "Confirmado",
          Confirmado: true,
          Origem: "Correção de integridade",
          Opcao: opcao,
          Observacao: `Correção de integridade: ${justificativa}`
        },
        pedido: this._resumoPedido(p)
      };
    },

    _simularPlano(resumoPorDia, acoes) {
      const out = JSON.parse(JSON.stringify(resumoPorDia || {}));
      for (const a of acoes || []) {
        if (!a.dia || !a.opcao || !out[a.dia] || a.acao === "revisar") continue;
        if (a.acao === "cancelar") {
          out[a.dia].total = Math.max(0, (out[a.dia].total || 0) - 1);
          out[a.dia][a.opcao] = Math.max(0, (out[a.dia][a.opcao] || 0) - 1);
        }
        if (a.acao === "reativar") {
          out[a.dia].total = (out[a.dia].total || 0) + 1;
          out[a.dia][a.opcao] = (out[a.dia][a.opcao] || 0) + 1;
        }
      }
      return out;
    },

    _agruparAcoesPorDia(acoes) {
      return (acoes || []).reduce((acc, a) => {
        const dia = a.dia || "sem-dia";
        if (!acc[dia]) acc[dia] = [];
        acc[dia].push(a);
        return acc;
      }, {});
    },

    // ============================================================
    // VERIFICAÇÕES
    // ============================================================
    _verificarPedidosQuebrados(dados, semanaId) {
      const out = [];
      for (const p of dados.pedidos || []) {
        const nome = this._nome(p);
        const dia = this._dia(p);
        const opcao = this._opcao(p);
        const semana = this._pick(p, "Semana_id", "Semana");
        const id = this._id(p);
        const faltas = [];
        if (!semana) faltas.push("Semana_id");
        if (!dia) faltas.push("Dia");
        if (!nome && !this._colabId(p)) faltas.push("Colaborador");
        if (!opcao && !this._pedidoAusencia(p)) faltas.push("Opcao");
        if (faltas.length) {
          out.push(this._issue({
            tipo: "pedido-incompleto",
            severidade: "atencao",
            nome: nome || `Pedido ${id || "sem ID"}`,
            semanaId,
            dia,
            pedido: p,
            pedidoId: id,
            mensagem: `Pedido com campo(s) obrigatório(s) ausente(s): ${faltas.join(", ")}.`,
            acaoSugerida: "Revisar o registro em Pedidos antes de considerar na produção."
          }));
        }
      }
      return out;
    },

    _verificarPedidosBloqueadosConfirmados(dados, semanaId) {
      const out = [];
      for (const p of dados.pedidos || []) {
        const confirmado = this._isTrue(this._pick(p, "Confirmado", "confirmado"));
        const status = this._status(p);
        if (!confirmado || !this._statusBloqueiaProducao(status)) continue;
        if (this._isTravamentoAutomaticoConfirmado(p)) continue;
        out.push(this._issue({
          tipo: "pedido-bloqueado-confirmado",
          severidade: "critico",
          nome: this._nome(p),
          semanaId,
          dia: this._dia(p),
          pedido: p,
          pedidoId: this._id(p),
          mensagem: `Pedido está Confirmado=true, mas Status=${status || "—"} bloqueia produção.`,
          acaoSugerida: "Ajustar Confirmado=false ou reprocessar ausência/cancelamento."
        }));
      }
      return out;
    },

    _verificarDuplicidadesNormais(dados, semanaId) {
      const grupos = new Map();
      const out = [];
      for (const p of dados.pedidos || []) {
        if (this._isExtraPedido(p) || this._isPedidoAdicionalColaborador(p)) continue;
        const dia = this._norm(this._dia(p));
        const keyColab = this._colabKey(p);
        if (!dia || !keyColab) continue;
        const key = `${semanaId}|${dia}|${keyColab}`;
        if (!grupos.has(key)) grupos.set(key, []);
        grupos.get(key).push(p);
      }
      for (const [key, grupo] of grupos.entries()) {
        if (grupo.length < 2) continue;
        const ordenados = grupo.slice().sort((a, b) => this._compararPreferido(a, b));
        const preferido = ordenados[0];
        const validos = grupo.filter(p => this._pedidoContaProducao(p));
        out.push(this._issue({
          tipo: "duplicidade-pedido",
          severidade: validos.length > 1 ? "critico" : "info",
          nome: this._nome(preferido),
          semanaId,
          dia: this._dia(preferido),
          pedidoId: this._id(preferido),
          mensagem: `${grupo.length} pedidos encontrados para o mesmo colaborador/dia. ${validos.length > 1 ? "Mais de um conta produção." : "Duplicidade técnica encontrada."}`,
          acaoSugerida: validos.length > 1 ? "Manter apenas o pedido produtivo correto." : "Não afeta produção se obsoletos estão bloqueados/cancelados.",
          detalhes: { chave: key, pedidoPreferido: this._resumoPedido(preferido), todos: ordenados.map(p => this._resumoPedido(p)) }
        }));
      }
      return out;
    },

    _verificarDuplicidadesExtras(dados, semanaId) {
      const grupos = this._gruparExtrasProducao(dados.pedidos || []);
      const out = [];
      for (const [key, grupo] of grupos.entries()) {
        if (grupo.length < 2) continue;
        const ordenados = grupo.slice().sort((a, b) => this._timestamp(a) - this._timestamp(b));
        out.push(this._issue({
          tipo: "duplicidade-extra",
          severidade: "critico",
          nome: this._nome(ordenados[0]) || "Extra",
          semanaId,
          dia: this._dia(ordenados[0]),
          pedidoId: this._id(ordenados[0]),
          mensagem: `${grupo.length} extras/pedidos especiais iguais encontrados no mesmo dia/opção.`,
          acaoSugerida: `Manter ID ${this._id(ordenados[0])} e cancelar os demais no plano de correção.`,
          detalhes: { chave: key, manter: this._resumoPedido(ordenados[0]), duplicados: ordenados.slice(1).map(p => this._resumoPedido(p)) }
        }));
      }
      return out;
    },

    _verificarAusenciasVigentesComPedidoConfirmado(dados, semanaId, diaInfo) {
      const out = [];
      const dataRef = diaInfo.data;
      const pedidosDia = (dados.pedidos || []).filter(p => this._norm(this._dia(p)) === this._norm(diaInfo.dia));
      for (const a of dados.ausencias || []) {
        if (!this._ausenciaAtiva(a) || !this._ausenciaCobreData(a, dataRef)) continue;
        const keyAus = this._colabKey({ Colaborador_id: this._pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId"), Colaborador_nome: this._pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title") });
        if (!keyAus) continue;
        const confirmados = pedidosDia.filter(p => this._colabKey(p) === keyAus && this._pedidoContaProducao(p));
        for (const p of confirmados) {
          out.push(this._issue({
            tipo: "ausencia-vigente-com-pedido-confirmado",
            severidade: "critico",
            nome: this._nome(p) || this._nome(a),
            semanaId,
            dia: diaInfo.dia,
            data: dataRef,
            pedidoId: this._id(p),
            ausenciaId: this._id(a),
            mensagem: `Ausência vigente cobre ${this._br(dataRef)}, mas existe pedido confirmado na produção.`,
            acaoSugerida: "Reprocessar ausência para o dia ou alterar pedido para status de ausência."
          }));
        }
      }
      return out;
    },

    _verificarRetornosEsperados(dados, semanaId, diaInfo) {
      const out = [];
      const dataRef = diaInfo.data;
      const pedidosDia = (dados.pedidos || []).filter(p => this._norm(this._dia(p)) === this._norm(diaInfo.dia));
      const pedidosPorColab = new Map();
      for (const p of pedidosDia) {
        if (this._isExtraPedido(p) || this._isPedidoAdicionalColaborador(p)) continue;
        const key = this._colabKey(p);
        if (!key) continue;
        const atual = pedidosPorColab.get(key);
        if (!atual || this._compararPreferido(p, atual) < 0) pedidosPorColab.set(key, p);
      }
      for (const c of dados.colaboradores || []) {
        const key = this._colabKey(c);
        if (!key) continue;
        if (this._ausenciaVigenteParaKey(dados.ausencias, key, dataRef)) continue;
        const encerradas = (dados.ausencias || []).filter(a => {
          const keyAus = this._colabKey({ Colaborador_id: this._pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId"), Colaborador_nome: this._pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title") });
          return keyAus === key && this._ausenciaAtiva(a) && this._ausenciaFim(a) && this._ausenciaFim(a) < dataRef;
        });
        if (!encerradas.length) continue;
        encerradas.sort((a, b) => String(this._ausenciaFim(b)).localeCompare(String(this._ausenciaFim(a))));
        const pedido = pedidosPorColab.get(key);
        if (!pedido) {
          out.push(this._issue({
            tipo: "ausencia-encerrada-sem-retorno",
            severidade: "critico",
            nome: this._nome(c),
            semanaId,
            dia: diaInfo.dia,
            data: dataRef,
            ausenciaId: this._id(encerradas[0]),
            mensagem: `Ausência encerrou em ${this._br(this._ausenciaFim(encerradas[0]))}, mas não existe pedido para ${this._br(dataRef)}.`,
            acaoSugerida: "Gerar retorno automático apenas se o dia não estiver encerrado. Para dia passado, usar correção assistida."
          }));
        } else if (this._pedidoAusencia(pedido) && !this._pedidoContaProducao(pedido)) {
          out.push(this._issue({
            tipo: "retorno-bloqueado-por-status-antigo",
            severidade: "critico",
            nome: this._nome(c) || this._nome(pedido),
            semanaId,
            dia: diaInfo.dia,
            data: dataRef,
            pedidoId: this._id(pedido),
            ausenciaId: this._id(encerradas[0]),
            mensagem: `Ausência encerrou em ${this._br(this._ausenciaFim(encerradas[0]))}, mas o pedido continua como ${this._status(pedido) || "ausência"}.`,
            acaoSugerida: "Atualizar para Confirmado/Principal somente com confirmação se o dia já passou."
          }));
        }
      }
      return out;
    },

    // ============================================================
    // CÁLCULOS
    // ============================================================
    _calcularDetalhesProducaoDia(pedidos, dia) {
      const lista = (pedidos || []).filter(p => this._norm(this._dia(p)) === this._norm(dia));
      const dedup = this._deduplicar(lista);
      const incluidos = [];
      const excluidos = [];
      const resumo = { total: 0, principal: 0, light: 0, carne: 0, massa: 0, lanche: 0, outros: 0, semOpcao: 0 };
      for (const p of dedup) {
        const item = this._resumoPedido(p);
        item.motivo = this._motivoPedidoProducao(p);
        item.categoria = this._categoriaPedido(p);
        if (this._pedidoContaProducao(p)) {
          incluidos.push(item);
          const op = this._norm(this._opcao(p) || "principal");
          if (op && Object.prototype.hasOwnProperty.call(resumo, op)) resumo[op]++;
          else if (op) resumo.outros++;
          else resumo.semOpcao++;
          resumo.total++;
        } else {
          excluidos.push(item);
        }
      }
      return { dia, resumo, incluidos, excluidos };
    },

    _calcularResumoRetiradaDia(checkins, dia) {
      const lista = (checkins || []).filter(c => this._norm(this._dia(c)) === this._norm(dia));
      const validos = lista.filter(c => this._isTrue(this._pick(c, "Retirou", "retirou", "Confirmado", "confirmado")) || this._pick(c, "Data_Hora_Retirada", "DataRetirada"));
      const resumo = { total: 0, principal: 0, light: 0, carne: 0, massa: 0, lanche: 0, outros: 0, semOpcao: 0 };
      for (const c of validos) {
        const op = this._norm(this._opcao(c) || this._pick(c, "Opcao_Pedido", "opcaoPedido") || "");
        if (op && Object.prototype.hasOwnProperty.call(resumo, op)) resumo[op]++;
        else if (op) resumo.outros++;
        else resumo.semOpcao++;
        resumo.total++;
      }
      return resumo;
    },

    _compararDiaComReferencia(dia, producao, retirada, referencia) {
      const base = { principal: 0, light: 0, carne: 0, massa: 0, lanche: 0, total: 0 };
      const ref = referencia ? { ...base, ...referencia } : null;
      const prod = { ...base, ...(producao || {}) };
      const ret = { ...base, ...(retirada || {}) };
      const diferencas = {};
      for (const k of Object.keys(base)) diferencas[k] = ref ? (prod[k] || 0) - (ref[k] || 0) : 0;
      return {
        dia,
        temReferencia: !!ref,
        referencia: ref,
        producao: prod,
        retirada: ret,
        diferencas,
        bateProducaoComReferencia: !!ref && Object.values(diferencas).every(v => v === 0)
      };
    },

    // ============================================================
    // HTML
    // ============================================================
    _atualizarModal(resultado) {
      const title = document.getElementById("adminIntegridadeModalTitle");
      const body = document.getElementById("adminIntegridadeModalBody");
      if (!title || !body) return;
      title.textContent = resultado ? `Diagnóstico completo · ${resultado.semanaId}` : "Diagnóstico completo";
      body.innerHTML = resultado ? this._htmlDetalhes(resultado) : `<div class="integridade-empty">Nenhum diagnóstico carregado.</div>`;
    },

    _htmlDetalhes(r) {
      return `
        <div class="integridade-modal-actions">
          <button class="btn-secondary" onclick="AdminIntegridade.copiarDiagnostico()">Copiar JSON</button>
          <button class="btn-secondary" onclick="AdminIntegridade.baixarDiagnostico()">Baixar JSON completo</button>
          <button class="btn-secondary" onclick="AdminIntegridade.baixarPlanoCorrecao()">Baixar plano</button>
          <button class="btn-danger" onclick="AdminIntegridade.abrirCorrecaoAssistida()">Aplicar correção assistida</button>
        </div>
        ${this._htmlComparativo(r)}
        ${this._htmlPlano(r)}
        ${this._htmlIrregularidades(r)}
        ${this._htmlDetalhesProducao(r)}
      `;
    },

    _htmlComparativo(r) {
      const linhas = Object.values(r.comparativoPorDia || {}).map(c => {
        const ref = c.referencia;
        const p = c.producao || {};
        const ret = c.retirada || {};
        const difs = Object.entries(c.diferencas || {}).filter(([_, v]) => v !== 0).map(([k, v]) => `<span class="integridade-diff-pill">${this._esc(k)} ${v > 0 ? "+" : ""}${v}</span>`).join(" ") || `<span class="integridade-ok-pill">bate</span>`;
        return `<tr>
          <td><b>${this._esc(this._diaLabel(c.dia))}</b></td>
          <td>${ref ? this._htmlResumoInline(ref) : "—"}</td>
          <td>${this._htmlResumoInline(p)}</td>
          <td>${this._htmlResumoInline(ret)}</td>
          <td>${difs}</td>
        </tr>`;
      }).join("");
      return `<section class="integridade-section">
        <h3>📊 Comparativo por dia</h3>
        <p>Produção = pedidos que o sistema conta. Check-in = retiradas gravadas pela Cozinha. Referência = conferência manual informada pela Luana, quando cadastrada.</p>
        <div class="integridade-table-wrap"><table class="integridade-table"><thead><tr><th>Dia</th><th>Referência Luana</th><th>Produção calculada</th><th>Check-in</th><th>Diferença produção x Luana</th></tr></thead><tbody>${linhas}</tbody></table></div>
      </section>`;
    },

    _htmlPlano(r) {
      const plano = r.planoCorrecao || { acoes: [], totais: {} };
      const porDia = plano.porDia || {};
      const secoes = Object.entries(porDia).map(([dia, acoes]) => `
        <details class="integridade-details" open>
          <summary>${this._esc(this._diaLabel(dia))} · ${acoes.length} ação(ões)</summary>
          ${acoes.map(a => this._htmlAcaoPlano(a)).join("")}
        </details>`).join("");
      return `<section class="integridade-section integridade-plano">
        <h3>🛠 Plano de Correção sugerido</h3>
        <p>Prévia segura. A gravação só ocorre pelo botão de Correção Assistida, com confirmação e auditoria.</p>
        <div class="integridade-plano-totais">
          <span>Cancelar: <b>${this._num(plano.totais?.cancelar)}</b></span>
          <span>Reativar: <b>${this._num(plano.totais?.reativar)}</b></span>
          <span>Revisar: <b>${this._num(plano.totais?.revisar)}</b></span>
        </div>
        ${secoes || `<div class="integridade-empty ok">Nenhuma ação segura sugerida.</div>`}
        <h4>Resultado simulado após o plano</h4>
        ${this._htmlResultadoSimulado(r)}
      </section>`;
    },

    _htmlAcaoPlano(a) {
      const cls = a.acao === "cancelar" ? "critico" : a.acao === "reativar" ? "ok" : "atencao";
      const label = a.acao === "cancelar" ? "Cancelar" : a.acao === "reativar" ? "Reativar" : "Revisar";
      return `<article class="integridade-plan-row ${cls}">
        <div class="integridade-plan-top"><b>${this._esc(label)}</b><span>Pedido ID: ${this._esc(a.pedidoId || "—")}</span></div>
        <div><b>${this._esc(a.nome || "Registro")}</b> · ${this._esc(this._diaLabel(a.dia || ""))} · ${this._esc(a.opcao || "")}</div>
        <div class="integridade-issue-msg">${this._esc(a.justificativa || a.mensagem || "Revisar")}</div>
        ${a.camposSugeridos ? `<details class="integridade-json"><summary>Campos sugeridos</summary><pre>${this._esc(JSON.stringify(a.camposSugeridos, null, 2))}</pre></details>` : ""}
      </article>`;
    },

    _htmlResultadoSimulado(r) {
      const sim = r.planoCorrecao?.resultadoSimuladoPorDia || {};
      const linhas = Object.entries(sim).map(([dia, resumo]) => `<tr><td><b>${this._esc(this._diaLabel(dia))}</b></td><td>${this._htmlResumoInline(resumo)}</td></tr>`).join("");
      return `<div class="integridade-table-wrap"><table class="integridade-table"><thead><tr><th>Dia</th><th>Simulado</th></tr></thead><tbody>${linhas}</tbody></table></div>`;
    },

    _htmlIrregularidades(r) {
      const porTipo = new Map();
      for (const i of r.irregularidades || []) {
        if (!porTipo.has(i.tipo)) porTipo.set(i.tipo, []);
        porTipo.get(i.tipo).push(i);
      }
      if (!r.irregularidades?.length) return `<section class="integridade-section"><h3>⚠️ Irregularidades encontradas</h3><div class="integridade-empty ok">Nenhuma irregularidade encontrada.</div></section>`;
      const grupos = Array.from(porTipo.entries()).map(([tipo, itens]) => {
        const linhas = itens.slice(0, this._config.maxDetalhesPorGrupo).map(i => this._htmlIssue(i)).join("");
        const extra = itens.length > this._config.maxDetalhesPorGrupo ? `<div class="integridade-limite">+ ${itens.length - this._config.maxDetalhesPorGrupo} item(ns) ocultos. Baixe o JSON para ver tudo.</div>` : "";
        return `<details class="integridade-details"><summary>${this._esc(this._tituloTipo(tipo))} · ${itens.length}</summary>${linhas}${extra}</details>`;
      }).join("");
      return `<section class="integridade-section"><h3>⚠️ Irregularidades encontradas</h3>${grupos}</section>`;
    },

    _htmlDetalhesProducao(r) {
      const secoes = Object.values(r.detalhesProducaoPorDia || {}).map(d => `
        <details class="integridade-details">
          <summary>✅ ${this._esc(this._diaLabel(d.dia))} · incluídos ${d.incluidos.length} · excluídos ${d.excluidos.length}</summary>
          <h4>Incluídos na produção</h4>
          <div class="integridade-lista-compacta">${d.incluidos.map(p => this._htmlPedidoLinha(p)).join("")}</div>
          <h4>Excluídos da produção</h4>
          <div class="integridade-lista-compacta">${d.excluidos.map(p => this._htmlPedidoLinha(p)).join("")}</div>
        </details>`).join("");
      return `<section class="integridade-section"><h3>🔎 Detalhe da produção</h3>${secoes}</section>`;
    },

    _htmlPedidoLinha(p) {
      return `<div class="integridade-pedido-linha">
        <span><b>${this._esc(p.nome || "—")}</b></span>
        <span>ID ${this._esc(p.id || "—")}</span>
        <span>${this._esc(p.opcao || "—")}</span>
        <span>${this._esc(p.status || "—")}</span>
        <span>${this._esc(p.origem || "—")}</span>
      </div>`;
    },

    _htmlIssue(i) {
      const sev = i.severidade || "info";
      const pedido = i.pedidoId ? `<span>Pedido ID: <b>${this._esc(i.pedidoId)}</b></span>` : "";
      const ausencia = i.ausenciaId ? `<span>Ausência ID: <b>${this._esc(i.ausenciaId)}</b></span>` : "";
      const data = i.data ? `<span>Data: <b>${this._esc(this._br(i.data))}</b></span>` : "";
      const dia = i.dia ? `<span>Dia: <b>${this._esc(this._diaLabel(i.dia))}</b></span>` : "";
      const detalhes = i.detalhes ? `<details class="integridade-json"><summary>Ver detalhes técnicos</summary><pre>${this._esc(JSON.stringify(i.detalhes, null, 2))}</pre></details>` : "";
      return `<article class="integridade-issue ${this._esc(sev)}">
        <div class="integridade-issue-top">
          <div><div class="integridade-issue-nome">${this._esc(i.nome || "Registro")}</div><div class="integridade-issue-msg">${this._esc(i.mensagem || "Irregularidade encontrada.")}</div></div>
          <span class="integridade-sev ${this._esc(sev)}">${this._esc(this._severidadeLabel(sev))}</span>
        </div>
        <div class="integridade-meta">${[dia, data, pedido, ausencia].filter(Boolean).join("")}</div>
        <div class="integridade-acao"><b>Ação sugerida:</b> ${this._esc(i.acaoSugerida || "Revisar manualmente antes de corrigir.")}</div>${detalhes}
      </article>`;
    },

    _htmlResumoInline(r) {
      return `T:${this._num(r?.total)} · P:${this._num(r?.principal)} · L:${this._num(r?.light)} · C:${this._num(r?.carne)} · M:${this._num(r?.massa)}`;
    },

    // ============================================================
    // HELPERS DE REGRA
    // ============================================================
    _R() { return global.HomyRefeitorioRegras || null; },
    _pick(obj, ...keys) {
      const R = this._R();
      if (R?.pick) return R.pick(obj, ...keys);
      for (const k of keys) if (obj && obj[k] !== undefined && obj[k] !== null) return obj[k];
      return "";
    },
    _norm(v) {
      const R = this._R();
      if (R?.norm) return R.norm(v);
      return String(v ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    },
    _isTrue(v) {
      const R = this._R();
      if (R?.isTrue) return R.isTrue(v);
      if (v === true || v === 1) return true;
      return ["sim", "true", "yes", "1"].includes(this._norm(v));
    },
    _id(obj) { return String(this._pick(obj, "id", "ID") || "").trim(); },
    _nome(obj) { return this._pick(obj, "Colaborador_nome", "colaborador_nome", "Colaborador", "Nome", "Title", "nome") || ""; },
    _colabId(obj) { return String(this._pick(obj, "Colaborador_id", "colaborador_id", "ColaboradorId", "colaboradorId", "Matricula", "Matrícula", "id", "ID") || "").trim(); },
    _dia(obj) { return this._pick(obj, "Dia", "dia") || ""; },
    _opcao(obj) { return this._pick(obj, "Opcao", "opcao", "Opção") || ""; },
    _status(obj) { return this._pick(obj, "Status", "status") || ""; },
    _origem(obj) { return this._pick(obj, "Origem", "origem", "tipo", "Tipo") || ""; },
    _colaboradorAtivo(c) {
      const ativo = this._pick(c, "Ativo", "ativo");
      if (ativo === null || ativo === undefined || String(ativo).trim() === "") return true;
      return this._isTrue(ativo);
    },
    _colabKey(obj) {
      const R = this._R();
      if (R?.colaboradorKey) return R.colaboradorKey(obj);
      const id = this._colabId(obj);
      if (id) return `id:${id}`;
      const nome = this._norm(this._nome(obj));
      return nome ? `nome:${nome}` : "";
    },
    _statusBloqueiaProducao(status) {
      return ["cancelado", "bloqueado", "nao vai almocar", "nao_vai_almocar", "não vai almoçar", "ausente", "ferias", "férias", "afastado", "afastamento", "atestado", "licenca", "licença", "banco horas", "banco_horas", "homy office", "homy_office", "falta"].includes(this._norm(status));
    },
    _pedidoContaProducao(p) {
      if (this._isTravamentoAutomaticoConfirmado(p)) return true;
      const R = this._R();
      if (R?.pedidoConfirmadoProducao) return R.pedidoConfirmadoProducao(p);
      const status = this._norm(this._status(p));
      if (this._statusBloqueiaProducao(status)) return false;
      return ["confirmado", "aprovado", "extra"].includes(status) || this._isTrue(this._pick(p, "Confirmado", "confirmado"));
    },
    _pedidoAusencia(p) {
      const status = this._norm(this._status(p));
      const origem = this._norm(this._origem(p));
      return this._statusBloqueiaProducao(status) || origem.includes("ausencia") || origem.includes("ausência");
    },
    _isTravamentoAutomaticoConfirmado(p) {
      return this._norm(this._status(p)) === "travado" && this._norm(this._origem(p)).includes("travamento");
    },
    _isPedidoAdicionalColaborador(p) {
      const origem = this._norm(this._origem(p));
      const cid = this._norm(this._colabId(p));
      const obs = this._norm(this._pick(p, "Observacao", "Observação", "observacao", "Obs"));
      return origem.includes("segunda refeicao") || origem.includes("segunda refeição") || origem.includes("refeicao adicional") || origem.includes("refeição adicional") || obs.includes("adicionalid:") || obs.includes("colaboradorbaseid:") || cid.includes("-adicional-");
    },
    _isExtraPedido(p) {
      const origem = this._norm(this._origem(p));
      const nome = this._norm(this._nome(p));
      const cid = this._norm(this._colabId(p));
      return origem.includes("extra") || origem.includes("guarda") || origem.includes("investigador") || origem.includes("visitante") || origem.includes("prestador") || origem.includes("terceiro") || origem.includes("fornecedor") || origem.includes("representante") || origem.includes("motorista") || nome.includes("refeicao extra") || /^guarda\s*\d*$/.test(nome) || /^investigador(?:\s*\d+)?$/.test(nome) || cid.startsWith("extra-");
    },
    _compararPreferido(a, b) {
      const sa = this._scorePedido(a), sb = this._scorePedido(b);
      if (sa !== sb) return sb - sa;
      return this._timestamp(b) - this._timestamp(a);
    },
    _scorePedido(p) {
      if (this._isTravamentoAutomaticoConfirmado(p)) return 95;
      const status = this._norm(this._status(p));
      if (["cancelado", "bloqueado"].includes(status)) return -50;
      if (this._pedidoContaProducao(p)) return 90;
      if (this._pedidoAusencia(p)) return 60;
      return 10;
    },
    _timestamp(p) {
      const raw = this._pick(p, "Modified", "modified", "Data_Hora", "DataHora", "Created", "created", "Data") || "";
      const dt = raw ? new Date(raw) : null;
      if (dt && !isNaN(dt)) return dt.getTime();
      const id = Number(this._id(p) || 0);
      return Number.isFinite(id) ? id : 0;
    },
    _deduplicar(lista) {
      const mapa = new Map();
      for (const p of lista || []) {
        const key = this._isExtraPedido(p) || this._isPedidoAdicionalColaborador(p)
          ? `especial|${this._keyExtra(p)}|id:${this._id(p)}`
          : `colab|${this._norm(this._dia(p))}|${this._colabKey(p)}`;
        const atual = mapa.get(key);
        if (!atual || this._compararPreferido(p, atual) < 0) mapa.set(key, p);
      }
      return Array.from(mapa.values());
    },
    _statusCanceladoOuDuplicado(p) {
      const status = this._norm(this._status(p));
      const origem = this._norm(this._origem(p));
      return status === "cancelado" || status === "bloqueado" || origem.includes("duplicado inativado") || origem.includes("duplicidade");
    },
    _retornoAutomaticoRetroativo(p, dataISO) {
      if (!this._norm(this._origem(p)).includes("retorno automatico") && !this._norm(this._origem(p)).includes("retorno automático")) return false;
      const modified = String(this._pick(p, "Modified", "modified") || "").slice(0, 10);
      const hoje = this._hojeISO();
      return (!!modified && modified > dataISO) || dataISO < hoje;
    },
    _gruparExtrasProducao(pedidos) {
      const grupos = new Map();
      for (const p of pedidos || []) {
        if (!this._isExtraPedido(p) || !this._pedidoContaProducao(p)) continue;
        const key = this._keyExtra(p);
        if (!key) continue;
        if (!grupos.has(key)) grupos.set(key, []);
        grupos.get(key).push(p);
      }
      return grupos;
    },
    _keyExtra(p) {
      const semana = String(this._pick(p, "Semana_id", "Semana") || "");
      const dia = this._norm(this._dia(p));
      const nome = this._norm(this._nome(p));
      const origem = this._norm(this._origem(p));
      const opcao = this._norm(this._opcao(p) || "principal");
      return `${semana}|${dia}|${origem}|${nome}|${opcao}`;
    },
    _temOutroExtraIgualMaisAntigo(p, lista) {
      const key = this._keyExtra(p);
      const ts = this._timestamp(p);
      return (lista || []).some(x => x !== p && this._isExtraPedido(x) && this._keyExtra(x) === key && this._timestamp(x) <= ts);
    },
    _categoriaPedido(p) {
      if (this._isExtraPedido(p)) {
        const origem = this._norm(this._origem(p));
        if (origem.includes("guarda")) return "guarda";
        if (origem.includes("investigador")) return "investigador";
        if (origem.includes("prestador")) return "prestador";
        if (origem.includes("visitante")) return "visitante";
        return "extra";
      }
      return "colaborador";
    },
    _motivoPedidoProducao(p) {
      if (this._pedidoContaProducao(p)) return this._isTravamentoAutomaticoConfirmado(p) ? "Travamento automático confirmado." : `Status produtivo: ${this._status(p) || "Confirmado"}.`;
      if (this._statusBloqueiaProducao(this._status(p))) return `Status bloqueia produção: ${this._status(p)}.`;
      return "Pedido não confirmado para produção.";
    },
    _resumoPedido(p) {
      return { id: this._id(p), colaboradorId: this._colabId(p), nome: this._nome(p), dia: this._dia(p), opcao: this._opcao(p), status: this._status(p), confirmado: this._isTrue(this._pick(p, "Confirmado", "confirmado")), origem: this._origem(p), dataHora: this._pick(p, "Data_Hora", "DataHora"), modified: this._pick(p, "Modified", "modified"), contaProducao: this._pedidoContaProducao(p) };
    },
    _ausenciaAtiva(a) {
      const status = this._norm(this._pick(a, "Status", "status", "Status_Ausencia", "statusAusencia") || "");
      if (["inativo", "cancelado", "cancelada", "duplicado inativado", "duplicidade inativada", "false", "nao", "não", "0"].includes(status)) return false;
      const ativo = this._pick(a, "Ativo", "ativo");
      if (ativo === null || ativo === undefined || ativo === "") return true;
      return this._isTrue(ativo);
    },
    _ausenciaInicio(a) { return this._dataISO(this._pick(a, "Data_Inicio", "Inicio", "DataInicio", "Data")); },
    _ausenciaFim(a) { return this._dataISO(this._pick(a, "Data_Fim", "Fim", "DataFim", "Data")) || this._ausenciaInicio(a); },
    _ausenciaCobreData(a, dataISO) { const ini = this._ausenciaInicio(a), fim = this._ausenciaFim(a); return !!ini && !!fim && ini <= dataISO && fim >= dataISO; },
    _ausenciaVigenteParaKey(ausencias, key, dataISO) {
      return (ausencias || []).find(a => {
        const keyAus = this._colabKey({ Colaborador_id: this._pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId"), Colaborador_nome: this._pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title") });
        return keyAus === key && this._ausenciaAtiva(a) && this._ausenciaCobreData(a, dataISO);
      }) || null;
    },

    _issue(dados) { return { id: dados.id || `${dados.tipo}-${dados.pedidoId || dados.ausenciaId || dados.nome || Date.now()}-${dados.dia || ""}`, ...dados }; },
    _referenciaSemana(semanaId, options = {}) { return options.referenciaOperacional?.[semanaId] || this._config.referenciaOperacional?.[semanaId] || null; },
    _diasParaVerificar(semanaId, options = {}) {
      const nomes = ["segunda", "terca", "quarta", "quinta", "sexta"];
      const datas = typeof SP?.getWeekDates === "function" ? SP.getWeekDates(semanaId) : [];
      const hoje = this._hojeISO();
      return nomes.map((dia, idx) => ({ dia, data: this._dataISO(datas[idx]) })).filter(d => d.data && (!options.checarSomenteAteHoje || d.data <= hoje));
    },
    _semanaAtual() { return global.AdminState?.getSemanaId?.() || global.SP?.getCurrentWeekId?.() || this._semanaId(new Date()); },
    _semanaId(date) { if (global.SP?.getSemanaId) return SP.getSemanaId(date); return ""; },
    _hojeISO() { return new Date().toISOString().slice(0, 10); },
    _dataPassadaOuHoje(dataISO) { return !!dataISO && dataISO <= this._hojeISO(); },
    _dataISO(v) { if (!v) return ""; if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10); const s = String(v); if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10); const d = new Date(s); return d && !isNaN(d) ? d.toISOString().slice(0, 10) : ""; },
    _br(iso) { if (!iso) return "—"; const [y, m, d] = String(iso).slice(0, 10).split("-"); return d && m && y ? `${d}/${m}/${y}` : iso; },
    _diaLabel(dia) { const n = this._norm(dia); return { segunda: "Segunda", terca: "Terça", quarta: "Quarta", quinta: "Quinta", sexta: "Sexta" }[n] || String(dia || ""); },
    _fmtDataHora(iso) { try { return new Date(iso).toLocaleString("pt-BR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }); } catch (_) { return ""; } },
    _num(v) { return Number(v || 0); },
    _esc(v) { return String(v ?? "").replace(/[&<>'"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[c])); },
    _severidadeLabel(s) { return { critico: "CRÍTICO", alerta: "ALERTA", atencao: "ATENÇÃO", info: "INFO", ok: "OK" }[s] || String(s || "INFO").toUpperCase(); },
    _tituloTipo(t) { return ({ "diferenca-referencia-operacional": "Diferença contra referência Luana", "pedido-incompleto": "Pedidos incompletos", "pedido-bloqueado-confirmado": "Bloqueados confirmados", "duplicidade-pedido": "Pedidos duplicados", "duplicidade-extra": "Extras duplicados", "ausencia-vigente-com-pedido-confirmado": "Ausência com pedido confirmado", "ausencia-encerrada-sem-retorno": "Ausência encerrada sem retorno", "retorno-bloqueado-por-status-antigo": "Retorno bloqueado por status antigo", "checkin-sem-dados-dia": "Check-in sem dados" })[t] || t; },
    _formatarMotivoAusencia(motivo) { const n = this._norm(motivo); if (n.includes("ferias")) return "Férias"; if (n.includes("almocar")) return "Não vai almoçar"; return motivo || "Ausente"; },
    _toast(msg, tipo = "info") { if (global.AdminUtils?.toast) AdminUtils.toast(msg, tipo); else console.log(`[${tipo}] ${msg}`); },

    // ============================================================
    // MODAL + CSS
    // ============================================================
    _ensureModal() {
      if (document.getElementById("adminIntegridadeModalOverlay")) return;
      const overlay = document.createElement("div");
      overlay.id = "adminIntegridadeModalOverlay";
      overlay.className = "admin-integridade-modal-overlay";
      overlay.innerHTML = `
        <div class="admin-integridade-modal" role="dialog" aria-modal="true">
          <div class="admin-integridade-modal-head">
            <div><h2 id="adminIntegridadeModalTitle">Diagnóstico completo</h2><p>Somente leitura · popup com rolagem interna · nenhuma correção automática nesta fase</p></div>
            <button class="admin-integridade-close" onclick="AdminIntegridade.fecharDetalhes()">×</button>
          </div>
          <div id="adminIntegridadeModalScroller" class="admin-integridade-modal-scroll"><div id="adminIntegridadeModalBody"></div></div>
        </div>`;
      overlay.addEventListener("click", e => { if (e.target === overlay) this.fecharDetalhes(); });
      document.addEventListener("keydown", e => { if (e.key === "Escape" && this._modalAberto) this.fecharDetalhes(); });
      document.body.appendChild(overlay);
    },

    _ensureStyle() {
      if (document.getElementById("adminIntegridadeStyle")) return;
      const style = document.createElement("style");
      style.id = "adminIntegridadeStyle";
      style.textContent = `
        .admin-integridade-card{margin-top:14px;border:1px solid rgba(255,255,255,.12);border-radius:14px;padding:14px;background:rgba(255,255,255,.045)}
        .admin-integridade-card.ok{border-color:rgba(64,208,144,.28);background:rgba(64,208,144,.08)}
        .admin-integridade-card.atencao{border-color:rgba(255,190,70,.35);background:rgba(255,160,40,.075)}
        .admin-integridade-card.critico{border-color:rgba(255,90,90,.42);background:rgba(192,40,28,.12)}
        .admin-integridade-head{display:flex;justify-content:space-between;gap:12px;align-items:flex-start}.admin-integridade-head-actions{display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end}
        .admin-integridade-title{font-weight:800;color:#fff}.admin-integridade-sub{font-size:.75rem;color:rgba(160,185,220,.68);margin-top:2px}.admin-integridade-status{margin-top:10px;color:#e8f0ff}.integridade-dif-alert{margin-top:10px;border-left:3px solid #ffd45a;padding:7px 10px;color:#ffe08a;background:rgba(255,210,80,.06)}.integridade-plano-mini{margin-top:8px;color:#cfe0ff;font-size:.82rem}
        .admin-integridade-pills{display:flex;gap:7px;flex-wrap:wrap;margin-top:10px}.admin-integridade-pill{font-size:.72rem;font-weight:800;padding:5px 9px;border-radius:999px;border:1px solid rgba(255,255,255,.12)}.admin-integridade-pill.critico{color:#ff8080;background:rgba(255,80,80,.1)}.admin-integridade-pill.alerta{color:#ffb56a;background:rgba(255,145,40,.1)}.admin-integridade-pill.atencao{color:#ffe082;background:rgba(255,210,60,.1)}.admin-integridade-pill.info{color:#8cc0ff;background:rgba(80,145,255,.1)}
        .integridade-preview-list{margin-top:12px;display:flex;flex-direction:column;gap:7px}.integridade-preview-row{border:1px solid rgba(255,255,255,.1);border-radius:10px;padding:8px 10px}.integridade-preview-main{font-weight:800;color:#fff}.integridade-preview-sub{font-size:.74rem;color:rgba(210,225,255,.72);margin-top:2px}.integridade-preview-row.critico{border-color:rgba(255,90,90,.35)}.integridade-preview-row.alerta{border-color:rgba(255,180,70,.35)}
        body.admin-integridade-modal-open{overflow:hidden!important}.admin-integridade-modal-overlay{position:fixed;inset:0;z-index:99999;display:none;align-items:center;justify-content:center;padding:24px;background:rgba(2,8,18,.78);backdrop-filter:blur(8px)}.admin-integridade-modal-overlay.open{display:flex}.admin-integridade-modal{width:min(1180px,96vw);height:min(88vh,920px);background:#07162d;border:1px solid rgba(90,150,255,.3);border-radius:18px;box-shadow:0 28px 90px rgba(0,0,0,.65);display:flex;flex-direction:column;overflow:hidden}.admin-integridade-modal-head{display:flex;justify-content:space-between;align-items:flex-start;gap:14px;padding:18px 20px;border-bottom:1px solid rgba(255,255,255,.08)}.admin-integridade-modal-head h2{margin:0;color:#fff;font-family:"Barlow Condensed",sans-serif;text-transform:uppercase;letter-spacing:.04em}.admin-integridade-modal-head p{margin:4px 0 0;color:rgba(150,175,210,.72);font-size:.8rem}.admin-integridade-close{width:38px;height:38px;border-radius:10px;border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.07);color:#dce8ff;font-size:24px;font-weight:800;cursor:pointer}.admin-integridade-modal-scroll{flex:1;overflow-y:auto;padding:18px 20px 28px;min-height:0}.integridade-modal-actions{display:flex;justify-content:flex-end;gap:10px;flex-wrap:wrap;margin-bottom:16px}.integridade-section{border:1px solid rgba(255,255,255,.1);background:rgba(255,255,255,.035);border-radius:16px;padding:14px;margin-bottom:16px}.integridade-section h3{margin:0 0 12px;color:#fff}.integridade-section h4{margin:14px 0 8px;color:#dce8ff}.integridade-section p{color:rgba(170,195,230,.75);font-size:.82rem;margin:0 0 12px;line-height:1.45}.integridade-table-wrap{overflow:auto;border-radius:12px;border:1px solid rgba(255,255,255,.08)}.integridade-table{width:100%;border-collapse:collapse;min-width:780px}.integridade-table th,.integridade-table td{padding:12px 14px;text-align:left;border-bottom:1px solid rgba(255,255,255,.07);color:#dce8ff}.integridade-table th{font-size:.72rem;text-transform:uppercase;letter-spacing:.12em;color:rgba(145,170,205,.75);background:rgba(255,255,255,.045)}.integridade-diff-pill,.integridade-ok-pill{display:inline-block;border-radius:999px;padding:3px 8px;margin:2px;font-size:.72rem;font-weight:800}.integridade-diff-pill{border:1px solid rgba(255,210,80,.45);color:#ffe082;background:rgba(255,210,80,.08)}.integridade-ok-pill{border:1px solid rgba(64,208,144,.35);color:#73e3a8;background:rgba(64,208,144,.08)}.integridade-plano-totais{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px}.integridade-plano-totais span{border:1px solid rgba(255,255,255,.1);border-radius:999px;padding:6px 10px;background:rgba(255,255,255,.045)}.integridade-details{border:1px solid rgba(255,255,255,.1);border-radius:13px;margin:10px 0;overflow:hidden}.integridade-details summary{cursor:pointer;padding:12px 14px;background:rgba(255,255,255,.045);font-weight:800;color:#fff}.integridade-plan-row,.integridade-issue{margin:10px 12px;padding:12px;border-radius:12px;border:1px solid rgba(255,255,255,.1);background:rgba(255,255,255,.035)}.integridade-plan-row.critico{border-color:rgba(255,90,90,.35)}.integridade-plan-row.ok{border-color:rgba(64,208,144,.35)}.integridade-plan-row.atencao{border-color:rgba(255,210,80,.35)}.integridade-plan-top,.integridade-issue-top{display:flex;justify-content:space-between;gap:10px;align-items:flex-start}.integridade-issue-nome{font-weight:900;color:#fff}.integridade-issue-msg{font-size:.82rem;color:rgba(210,225,255,.75);margin-top:3px}.integridade-sev{font-size:.68rem;font-weight:900;border-radius:999px;padding:4px 8px;border:1px solid rgba(255,255,255,.15)}.integridade-meta{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px;font-size:.75rem;color:#a9c1e5}.integridade-acao{margin-top:8px;font-size:.82rem;color:#dce8ff;border-left:3px solid rgba(100,180,255,.5);padding-left:9px}.integridade-json pre{white-space:pre-wrap;overflow:auto;background:rgba(0,0,0,.25);padding:10px;border-radius:8px;color:#cfe0ff}.integridade-lista-compacta{display:flex;flex-direction:column;gap:6px;margin:0 12px 12px}.integridade-pedido-linha{display:grid;grid-template-columns:1fr 80px 90px 120px 160px;gap:8px;padding:8px;border-radius:8px;background:rgba(255,255,255,.035);font-size:.78rem;color:#dce8ff}.integridade-empty{padding:14px;color:#cfe0ff}.integridade-empty.ok{color:#73e3a8}.integridade-limite{padding:10px 14px;color:#ffe082}
        @media(max-width:800px){.admin-integridade-modal-overlay{padding:8px}.admin-integridade-modal{width:98vw;height:92vh}.integridade-pedido-linha{grid-template-columns:1fr}.admin-integridade-head{flex-direction:column}.admin-integridade-head-actions{justify-content:flex-start}}
      `;
      document.head.appendChild(style);
    }
  };

  setTimeout(() => AdminIntegridade.instalarNoDashboard(), 0);
  document.addEventListener("DOMContentLoaded", () => AdminIntegridade.instalarNoDashboard());
})(window);
