// ============================================================
// admin-integridade.js — Integridade dos Dados · Admin Homy
// v: integridade-dashboard-readonly-20260625-fase4
//
// Objetivo desta fase:
// - Ler dados reais do SharePoint via SP/sharepoint.js.
// - Mostrar divergências no Dashboard > Alertas.
// - Não gravar nada no SharePoint.
// - Não usar modal com scroll travado: detalhes abrem inline no Dashboard.
// - Exportar JSON completo com incluídos/excluídos para encontrar a causa real.
// ============================================================
(function (global) {
  "use strict";

  const DIA_LABEL = {
    segunda: "Segunda",
    terca: "Terça",
    terça: "Terça",
    quarta: "Quarta",
    quinta: "Quinta",
    sexta: "Sexta"
  };

  const ORDEM_DIAS = ["segunda", "terca", "quarta", "quinta", "sexta"];

  const REFERENCIAS_OPERACIONAIS = {
    // Referência operacional informada pela Luana para conferência da semana 2026-W26.
    // Esta referência NÃO altera pedido e NÃO corrige SharePoint; serve apenas para apontar diferença.
    "2026-W26": {
      segunda: { principal: 44, light: 4, carne: 17, massa: 0, lanche: 0, total: 65 },
      terca:   { principal: 49, light: 6, carne: 8,  massa: 6, lanche: 0, total: 69 },
      quarta:  { principal: 45, light: 3, carne: 17, massa: 5, lanche: 0, total: 70 },
      quinta:  { principal: 57, light: 5, carne: 4,  massa: 4, lanche: 0, total: 70 }
    }
  };

  const AdminIntegridade = global.AdminIntegridade = {
    _ultimoResultado: null,
    _loading: false,
    _instalado: false,
    _cache: new Map(),
    _config: {
      containerId: "dashAlertas",
      cacheMs: 12000,
      checarSomenteAteHoje: true,
      maxPreview: 5
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
        this._removerModalAntigo();

        const cacheKey = `${semana}|${cfg.checarSomenteAteHoje ? "ate-hoje" : "semana-inteira"}`;
        const cached = this._cache.get(cacheKey);
        if (cached && Date.now() - cached.ts < cfg.cacheMs) {
          this._ultimoResultado = cached.resultado;
          this.renderResumo(cached.resultado, cfg);
          return cached.resultado;
        }

        const resultado = await this.verificarSemana(semana, cfg);
        this._ultimoResultado = resultado;
        this._cache.set(cacheKey, { ts: Date.now(), resultado });
        this.renderResumo(resultado, cfg);
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
      const referencia = REFERENCIAS_OPERACIONAIS[semanaId] || null;

      const producao = {};
      const retiradas = {};
      const comparativo = {};
      const irregularidades = [];

      for (const diaInfo of dias) {
        const dia = diaInfo.dia;
        producao[dia] = this._calcularProducaoDia(dados.pedidos, dia);
        retiradas[dia] = this._calcularRetiradasDia(dados.checkins, dados.pedidos, dia);
        comparativo[dia] = this._compararDia({
          dia,
          producao: producao[dia].resumo,
          retirada: retiradas[dia].resumo,
          referencia: referencia?.[dia] || null
        });

        irregularidades.push(...this._irregularidadesComparativo(semanaId, diaInfo, comparativo[dia]));
      }

      irregularidades.push(...this._verificarPedidosQuebrados(dados, semanaId));
      irregularidades.push(...this._verificarBloqueadosConfirmados(dados, semanaId));
      irregularidades.push(...this._verificarDuplicidades(dados, semanaId));
      irregularidades.push(...this._verificarCheckinsInconsistentes(dados, semanaId, dias, retiradas));

      irregularidades.sort((a, b) => {
        const peso = { critico: 0, alerta: 1, atencao: 2, info: 3 };
        return (peso[a.severidade] ?? 9) - (peso[b.severidade] ?? 9) ||
          String(a.nome || a.tipo || "").localeCompare(String(b.nome || b.tipo || ""), "pt-BR", { sensitivity: "base" });
      });

      const porSeveridade = irregularidades.reduce((acc, i) => {
        acc[i.severidade] = (acc[i.severidade] || 0) + 1;
        return acc;
      }, {});

      return {
        semanaId,
        geradoEm: new Date().toISOString(),
        observacao: "Somente leitura. Nenhuma correção automática foi executada.",
        referenciaOperacional: referencia ? { fonte: "Luana", valores: referencia } : null,
        diasVerificados: dias,
        dadosResumo: {
          pedidos: dados.pedidos.length,
          ausencias: dados.ausencias.length,
          colaboradores: dados.colaboradores.length,
          extras: dados.extras.length,
          cardapio: dados.cardapio.length,
          checkins: dados.checkins.length
        },
        resumoProducaoPorDia: this._somenteResumo(producao),
        resumoRetiradaPorDia: this._somenteResumo(retiradas),
        comparativoPorDia: comparativo,
        detalhesProducaoPorDia: producao,
        detalhesRetiradaPorDia: retiradas,
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

      let painel = document.getElementById("adminIntegridadePainelDetalhes");
      if (!painel) {
        painel = document.createElement("div");
        painel.id = "adminIntegridadePainelDetalhes";
        container.appendChild(painel);
      }

      const total = resultado.irregularidades.length;
      const criticos = resultado.porSeveridade.critico || 0;
      const alertas = resultado.porSeveridade.alerta || 0;
      const atencoes = resultado.porSeveridade.atencao || 0;
      const infos = resultado.porSeveridade.info || 0;
      const classe = resultado.status === "ok" ? "ok" : resultado.status === "critico" ? "critico" : "atencao";
      const titulo = resultado.status === "ok"
        ? "Integridade dos dados sem irregularidades críticas."
        : `${total} irregularidade(s) encontradas. Abra o popup para ver a causa.`;

      const diffs = Object.entries(resultado.comparativoPorDia || {})
        .filter(([, c]) => c.temReferencia && c.diferencas?.total !== 0)
        .map(([dia, c]) => `${this._diaLabel(dia)} ${this._sinal(c.diferencas.total)}`);

      const preview = diffs.length
        ? `<div class="integridade-diff-preview">Diferença contra referência Luana: ${this._esc(diffs.join(" · "))}</div>`
        : `<div class="integridade-diff-preview ok">Sem diferença contra referência operacional cadastrada.</div>`;

      const amostra = resultado.irregularidades.slice(0, options.maxPreview || this._config.maxPreview);
      const lista = amostra.length ? `
        <div class="integridade-preview-list">
          ${amostra.map(i => `
            <div class="integridade-preview-row ${this._esc(i.severidade)}">
              <div class="integridade-preview-main">${this._esc(i.nome || i.titulo || this._tituloTipo(i.tipo))}</div>
              <div class="integridade-preview-sub">${this._esc(i.mensagem || i.tipo || "Irregularidade")}</div>
            </div>
          `).join("")}
        </div>` : "";

      box.innerHTML = `
        <div class="admin-integridade-card ${classe}">
          <div class="admin-integridade-head">
            <div>
              <div class="admin-integridade-title">🔎 Integridade dos Dados</div>
              <div class="admin-integridade-sub">Semana ${this._esc(resultado.semanaId)} · ${this._fmtDataHora(resultado.geradoEm)} · somente leitura</div>
            </div>
            <div class="admin-integridade-actions">
              <button class="btn-secondary admin-integridade-btn" onclick="AdminIntegridade.abrirDetalhes()">Abrir detalhes</button>
              <button class="btn-secondary admin-integridade-btn" onclick="AdminIntegridade.baixarDiagnostico()">Baixar JSON</button>
            </div>
          </div>

          <div class="admin-integridade-status">${this._esc(titulo)}</div>
          ${preview}

          <div class="admin-integridade-pills">
            <span class="admin-integridade-pill critico">🔴 ${criticos} crítico(s)</span>
            <span class="admin-integridade-pill alerta">🟠 ${alertas} alerta(s)</span>
            <span class="admin-integridade-pill atencao">🟡 ${atencoes} atenção</span>
            <span class="admin-integridade-pill info">🔵 ${infos} info</span>
          </div>

          ${lista}
        </div>`;

      if (painel.dataset.open === "1") painel.innerHTML = this._htmlDetalhes(resultado);
      else painel.innerHTML = "";
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
      const r = this._ultimoResultado;
      this.fecharDetalhes();

      const overlay = document.createElement("div");
      overlay.id = "adminIntegridadeModalOverlay";
      overlay.setAttribute("role", "presentation");
      overlay.innerHTML = `
        <div class="admin-integridade-modal-shell" role="dialog" aria-modal="true" aria-labelledby="adminIntegridadeModalTitulo">
          <div class="admin-integridade-modal-header">
            <div>
              <div id="adminIntegridadeModalTitulo" class="admin-integridade-modal-title">Diagnóstico completo · ${this._esc(r?.semanaId || this._semanaAtual())}</div>
              <div class="admin-integridade-modal-sub">Somente leitura · popup com rolagem interna · nenhuma correção automática nesta fase</div>
            </div>
            <button type="button" class="admin-integridade-modal-close" aria-label="Fechar" onclick="AdminIntegridade.fecharDetalhes()">×</button>
          </div>
          <div class="admin-integridade-modal-toolbar">
            <button class="btn-secondary" onclick="AdminIntegridade.copiarDiagnostico()">Copiar JSON</button>
            <button class="btn-secondary" onclick="AdminIntegridade.baixarDiagnostico()">Baixar JSON completo</button>
          </div>
          <div class="admin-integridade-modal-body" tabindex="0">
            ${r ? this._htmlDetalhes(r) : `<div class="integridade-empty">Nenhum diagnóstico carregado.</div>`}
          </div>
        </div>`;

      overlay.addEventListener("click", ev => {
        if (ev.target === overlay) this.fecharDetalhes();
      });
      overlay.addEventListener("wheel", ev => {
        const body = overlay.querySelector(".admin-integridade-modal-body");
        if (!body) return;
        if (!body.contains(ev.target)) {
          body.scrollTop += ev.deltaY;
          ev.preventDefault();
        }
      }, { passive: false });

      this._escHandler = ev => {
        if (ev.key === "Escape") this.fecharDetalhes();
      };
      document.addEventListener("keydown", this._escHandler);
      document.documentElement.classList.add("admin-integridade-modal-open");
      document.body.classList.add("admin-integridade-modal-open");
      document.body.appendChild(overlay);
      setTimeout(() => overlay.querySelector(".admin-integridade-modal-body")?.focus(), 30);
    },

    fecharDetalhes() {
      document.getElementById("adminIntegridadeModalOverlay")?.remove();
      document.documentElement.classList.remove("admin-integridade-modal-open");
      document.body.classList.remove("admin-integridade-modal-open");
      if (this._escHandler) {
        document.removeEventListener("keydown", this._escHandler);
        this._escHandler = null;
      }
      const painel = document.getElementById("adminIntegridadePainelDetalhes");
      if (painel) {
        painel.dataset.open = "0";
        painel.innerHTML = "";
      }
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
        await self.load(semanaId, { cacheMs: 0 }).catch(e => console.warn("[AdminIntegridade] render ignorado:", e));
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
      const pedidosPromise = this._safeList(() => SP.getPedidos?.(semanaId, { reparar: false, force: true }), "Pedidos");
      const ausenciasPromise = this._safeList(() => SP.getAusencias?.(false), "Ausencias do Refeitorio");
      const colaboradoresPromise = this._safeList(() => SP.getTodosColaboradores?.(), "Colaboradores");
      const extrasPromise = this._safeList(() => SP.getExtras?.(semanaId), "Extras");
      const cardapioPromise = this._safeList(() => SP.getCardapio?.(semanaId), "Cardapio");
      const checkinsPromise = this._safeList(async () => {
        if (typeof SP.getItems === "function") return SP.getItems("CheckIn");
        const todos = [];
        for (const dia of ORDEM_DIAS) {
          const itens = await SP.getCheckIn?.(semanaId, dia).catch(() => []);
          if (Array.isArray(itens)) todos.push(...itens);
        }
        return todos;
      }, "CheckIn");

      const [pedidosRaw, ausencias, colaboradores, extras, cardapio, checkinsRaw] = await Promise.all([
        pedidosPromise,
        ausenciasPromise,
        colaboradoresPromise,
        extrasPromise,
        cardapioPromise,
        checkinsPromise
      ]);

      const pedidos = (pedidosRaw || []).filter(p => String(this._pick(p, "Semana_id", "Semana") || "") === String(semanaId));
      const checkins = (checkinsRaw || []).filter(c => String(this._pick(c, "Semana_id", "Semana") || "") === String(semanaId));

      return {
        pedidos,
        ausencias: ausencias || [],
        colaboradores: (colaboradores || []).filter(c => this._colaboradorAtivo(c)),
        extras: extras || [],
        cardapio: cardapio || [],
        checkins
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
    // CÁLCULO
    // ============================================================
    _calcularProducaoDia(pedidos, dia) {
      const rawDia = (pedidos || []).filter(p => this._norm(this._dia(p)) === this._norm(dia));
      const { mantidos, descartadosDuplicidade, grupos } = this._deduplicarPedidos(rawDia);
      const incluidos = [];
      const excluidos = [];

      for (const p of mantidos) {
        const decisao = this._decidirProducao(p);
        const item = this._resumoPedido(p);
        item.motivo = decisao.motivo;
        item.categoria = this._categoriaPedido(p);
        if (decisao.conta) incluidos.push(item);
        else excluidos.push(item);
      }

      for (const p of descartadosDuplicidade) {
        const item = this._resumoPedido(p);
        item.motivo = "Descartado da produção por duplicidade; outro pedido foi escolhido para o mesmo colaborador/dia.";
        item.categoria = this._categoriaPedido(p);
        excluidos.push(item);
      }

      return {
        dia,
        resumo: this._resumoPorOpcao(incluidos),
        incluidos,
        excluidos,
        duplicidadesTecnicas: grupos.filter(g => g.todos.length > 1).map(g => ({
          chave: g.chave,
          preferido: this._resumoPedido(g.preferido),
          todos: g.todos.map(p => this._resumoPedido(p))
        }))
      };
    },

    _calcularRetiradasDia(checkins, pedidos, dia) {
      const lista = (checkins || []).filter(c => this._norm(this._pick(c, "Dia")) === this._norm(dia) && this._isTrue(this._pick(c, "Retirou")));
      const pedidosPorKey = new Map();
      (pedidos || []).forEach(p => {
        if (this._norm(this._dia(p)) !== this._norm(dia)) return;
        pedidosPorKey.set(this._keyPessoa(p), p);
      });

      const incluidos = lista.map(c => {
        const pedido = pedidosPorKey.get(this._keyPessoa(c));
        const opcao = this._pick(c, "Opcao") || (pedido ? this._opcao(pedido) : "");
        return {
          id: this._id(c),
          colaboradorId: this._colabId(c),
          nome: this._nome(c),
          dia: this._pick(c, "Dia"),
          opcao: opcao || "semOpcao",
          retirou: this._pick(c, "Retirou"),
          dataHoraRetirada: this._pick(c, "Data_Hora_Retirada", "DataHoraRetirada"),
          confirmadoPor: this._pick(c, "Confirmado_Por"),
          pedidoIdRelacionado: pedido ? this._id(pedido) : "",
          statusPedidoRelacionado: pedido ? this._status(pedido) : ""
        };
      });

      return { dia, resumo: this._resumoPorOpcao(incluidos), incluidos };
    },

    _compararDia({ dia, producao, retirada, referencia }) {
      const base = { dia, temReferencia: !!referencia, referencia: referencia || null, producao, retirada, diferencas: {} };
      if (!referencia) return base;

      for (const k of ["total", "principal", "light", "carne", "massa", "lanche"]) {
        base.diferencas[k] = this._num(producao?.[k]) - this._num(referencia?.[k]);
      }
      base.bateProducaoComReferencia = Object.values(base.diferencas).every(v => v === 0);
      return base;
    },

    _irregularidadesComparativo(semanaId, diaInfo, comparativoDia) {
      const out = [];
      if (!comparativoDia.temReferencia) return out;
      const dif = comparativoDia.diferencas || {};
      const relevantes = Object.entries(dif).filter(([, v]) => v !== 0);
      if (!relevantes.length) return out;

      out.push(this._issue({
        tipo: "diferenca-referencia-operacional",
        severidade: Math.abs(dif.total || 0) >= 2 ? "critico" : "alerta",
        nome: this._diaLabel(diaInfo.dia),
        semanaId,
        dia: diaInfo.dia,
        mensagem: `Produção calculada não bate com referência Luana. Diferenças: ${relevantes.map(([k, v]) => `${k} ${this._sinal(v)}`).join(", ")}.`,
        acaoSugerida: "Abrir os detalhes do dia e comparar a lista de incluídos na produção com a referência operacional.",
        detalhes: comparativoDia
      }));
      return out;
    },

    _resumoPorOpcao(lista) {
      const r = { total: 0, principal: 0, light: 0, carne: 0, massa: 0, lanche: 0, outros: 0, semOpcao: 0 };
      for (const item of lista || []) {
        r.total++;
        const op = this._norm(item.opcao || "");
        if (!op) r.semOpcao++;
        else if (Object.prototype.hasOwnProperty.call(r, op)) r[op]++;
        else r.outros++;
      }
      return r;
    },

    _somenteResumo(mapa) {
      return Object.fromEntries(Object.entries(mapa || {}).map(([dia, obj]) => [dia, obj.resumo || obj]));
    },

    // ============================================================
    // VERIFICAÇÕES
    // ============================================================
    _verificarPedidosQuebrados(dados, semanaId) {
      const out = [];
      for (const p of dados.pedidos || []) {
        const faltas = [];
        if (!this._pick(p, "Semana_id", "Semana")) faltas.push("Semana_id");
        if (!this._dia(p)) faltas.push("Dia");
        if (!this._nome(p) && !this._colabId(p)) faltas.push("Colaborador");
        if (!this._opcao(p) && !this._pedidoAusencia(p)) faltas.push("Opcao");
        if (!faltas.length) continue;

        out.push(this._issue({
          tipo: "pedido-incompleto",
          severidade: "atencao",
          nome: this._nome(p) || `Pedido ${this._id(p) || "sem ID"}`,
          semanaId,
          dia: this._dia(p),
          pedidoId: this._id(p),
          mensagem: `Pedido com campo(s) obrigatório(s) ausente(s): ${faltas.join(", ")}.`,
          acaoSugerida: "Revisar o registro antes de considerar na produção.",
          detalhes: this._resumoPedido(p)
        }));
      }
      return out;
    },

    _verificarBloqueadosConfirmados(dados, semanaId) {
      const out = [];
      for (const p of dados.pedidos || []) {
        if (this._isEspecialQueNaoDeduplica(p)) continue;
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
          pedidoId: this._id(p),
          mensagem: `Pedido está Confirmado=true, mas Status=${status || "—"} bloqueia produção.`,
          acaoSugerida: "Ajustar Confirmado=false ou reprocessar ausência/cancelamento.",
          detalhes: this._resumoPedido(p)
        }));
      }
      return out;
    },

    _verificarDuplicidades(dados, semanaId) {
      const out = [];
      const grupos = new Map();
      for (const p of dados.pedidos || []) {
        if (this._isEspecialQueNaoDeduplica(p)) continue;
        const key = `${this._pick(p, "Semana_id", "Semana") || semanaId}|${this._norm(this._dia(p))}|${this._keyPessoa(p)}`;
        if (!grupos.has(key)) grupos.set(key, []);
        grupos.get(key).push(p);
      }

      for (const [chave, lista] of grupos.entries()) {
        if (lista.length <= 1) continue;
        const preferido = this._escolherPreferido(lista);
        const todosResumo = lista.map(p => this._resumoPedido(p));
        const validosProducao = lista.filter(p => this._decidirProducao(p).conta);
        const somenteResolvidos = validosProducao.length <= 1 && lista.every(p => !this._isTrue(this._pick(p, "Confirmado")) || !this._statusBloqueiaProducao(this._status(p)));
        const severidade = validosProducao.length > 1 ? "critico" : somenteResolvidos ? "info" : "alerta";

        out.push(this._issue({
          tipo: "duplicidade-pedido",
          severidade,
          nome: this._nome(preferido) || "Duplicidade",
          semanaId,
          dia: this._dia(preferido),
          pedidoId: this._id(preferido),
          mensagem: `${lista.length} pedidos encontrados para o mesmo colaborador/dia. ${validosProducao.length > 1 ? "Mais de um ainda pode contar na produção." : "Duplicidade técnica encontrada."}`,
          acaoSugerida: validosProducao.length > 1
            ? "Regularizar para deixar apenas um pedido válido de produção."
            : "Não afeta produção se os obsoletos estão cancelados, mas vale limpar a base.",
          detalhes: { chave, pedidoPreferido: this._resumoPedido(preferido), todos: todosResumo }
        }));
      }
      return out;
    },

    _verificarCheckinsInconsistentes(dados, semanaId, dias, retiradas) {
      const out = [];
      for (const diaInfo of dias) {
        const total = this._num(retiradas[diaInfo.dia]?.resumo?.total);
        const producao = this._num(this._ultimoResultado?.resumoProducaoPorDia?.[diaInfo.dia]?.total);
        if (total === 0) {
          out.push(this._issue({
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
      return out;
    },

    // ============================================================
    // HTML DETALHADO INLINE
    // ============================================================
    _htmlDetalhes(r) {
      return `
        <div class="integridade-modal-content">
          ${this._htmlComparativo(r)}
          ${this._htmlIrregularidades(r)}
          ${this._htmlIncluidosPorDia(r)}
          ${this._htmlExcluidosPorDia(r)}
        </div>`;
    },

    _htmlComparativo(r) {
      const linhas = ORDEM_DIAS
        .filter(dia => r.comparativoPorDia?.[dia])
        .map(dia => {
          const c = r.comparativoPorDia[dia];
          const ref = c.referencia || {};
          const p = c.producao || {};
          const ck = c.retirada || {};
          const d = c.diferencas || {};
          return `
            <tr class="${c.temReferencia && !c.bateProducaoComReferencia ? "linha-diff" : ""}">
              <td><strong>${this._diaLabel(dia)}</strong></td>
              <td>${c.temReferencia ? this._resumoCurto(ref) : "—"}</td>
              <td>${this._resumoCurto(p)}</td>
              <td>${this._resumoCurto(ck)}</td>
              <td>${c.temReferencia ? this._htmlDiferencas(d) : "Sem referência"}</td>
            </tr>`;
        }).join("");

      return `
        <section class="integridade-section">
          <div class="integridade-section-title">📊 Comparativo por dia</div>
          <div class="integridade-nota">Produção = pedidos que o sistema está contando. Check-in = retiradas gravadas pela Cozinha. Referência = conferência manual informada pela Luana, quando cadastrada.</div>
          <div class="table-wrap integridade-table-wrap">
            <table class="table integridade-table">
              <thead><tr><th>Dia</th><th>Referência Luana</th><th>Produção calculada</th><th>Check-in</th><th>Diferença produção x Luana</th></tr></thead>
              <tbody>${linhas}</tbody>
            </table>
          </div>
        </section>`;
    },

    _htmlIrregularidades(r) {
      const itens = r.irregularidades || [];
      if (!itens.length) return `<section class="integridade-section"><div class="integridade-empty ok">✅ Nenhuma irregularidade encontrada.</div></section>`;
      const grupos = new Map();
      itens.forEach(i => {
        const tipo = i.tipo || "outros";
        if (!grupos.has(tipo)) grupos.set(tipo, []);
        grupos.get(tipo).push(i);
      });
      return `
        <section class="integridade-section">
          <div class="integridade-section-title">⚠️ Irregularidades encontradas</div>
          ${Array.from(grupos.entries()).map(([tipo, lista]) => `
            <details class="integridade-details" open>
              <summary>${this._esc(this._tituloTipo(tipo))} <span class="badge badge-blue">${lista.length}</span></summary>
              ${lista.map(i => this._htmlIssue(i)).join("")}
            </details>
          `).join("")}
        </section>`;
    },

    _htmlIncluidosPorDia(r) {
      return `
        <section class="integridade-section">
          <div class="integridade-section-title">✅ Quem está entrando na produção</div>
          <div class="integridade-nota">Esta é a lista que explica o número calculado. Se um dia não bate com a Luana, procure aqui quem entrou a mais ou com opção errada.</div>
          ${ORDEM_DIAS.filter(dia => r.detalhesProducaoPorDia?.[dia]).map(dia => {
            const det = r.detalhesProducaoPorDia[dia];
            return `
              <details class="integridade-details">
                <summary>${this._diaLabel(dia)} · ${det.resumo.total} incluído(s) · ${this._resumoCurto(det.resumo)}</summary>
                ${this._htmlListaPedidos(det.incluidos, true)}
              </details>`;
          }).join("")}
        </section>`;
    },

    _htmlExcluidosPorDia(r) {
      return `
        <section class="integridade-section">
          <div class="integridade-section-title">🚫 Quem foi excluído da produção</div>
          <div class="integridade-nota">Aqui aparecem férias, cancelados, duplicados descartados e ausências. Eles não entram no total calculado.</div>
          ${ORDEM_DIAS.filter(dia => r.detalhesProducaoPorDia?.[dia]).map(dia => {
            const det = r.detalhesProducaoPorDia[dia];
            return `
              <details class="integridade-details">
                <summary>${this._diaLabel(dia)} · ${det.excluidos.length} excluído(s)</summary>
                ${this._htmlListaPedidos(det.excluidos, false)}
              </details>`;
          }).join("")}
        </section>`;
    },

    _htmlListaPedidos(lista, incluido) {
      if (!lista?.length) return `<div class="integridade-empty">Nenhum item.</div>`;
      const grupos = new Map();
      lista.forEach(p => {
        const op = this._norm(p.opcao || "semOpcao") || "semOpcao";
        if (!grupos.has(op)) grupos.set(op, []);
        grupos.get(op).push(p);
      });
      const ordem = ["principal", "light", "carne", "massa", "lanche", "outros", "semOpcao"];
      return ordem.filter(op => grupos.has(op)).map(op => `
        <div class="integridade-opcao-bloco">
          <div class="integridade-opcao-title">${this._esc(op)} · ${grupos.get(op).length}</div>
          <div class="integridade-lista-pedidos">
            ${grupos.get(op).map(p => `
              <div class="integridade-pedido-row ${incluido ? "incluido" : "excluido"}">
                <div>
                  <strong>${this._esc(p.nome || "Sem nome")}</strong>
                  <div class="integridade-pedido-sub">ID ${this._esc(p.id || "—")} · ${this._esc(p.status || "—")} · ${this._esc(p.origem || "—")} · ${this._esc(p.categoria || "—")}</div>
                  <div class="integridade-pedido-motivo">${this._esc(p.motivo || "—")}</div>
                </div>
                <span class="badge ${incluido ? "badge-green" : "badge-red"}">${this._esc(p.opcao || "—")}</span>
              </div>`).join("")}
          </div>
        </div>
      `).join("");
    },

    _htmlIssue(i) {
      return `
        <div class="integridade-issue ${this._esc(i.severidade || "info")}">
          <div class="integridade-issue-top">
            <div>
              <div class="integridade-issue-nome">${this._esc(i.nome || i.titulo || "Registro")}</div>
              <div class="integridade-issue-msg">${this._esc(i.mensagem || "Irregularidade")}</div>
            </div>
            <span class="integridade-sev ${this._esc(i.severidade || "info")}">${this._esc(this._severidadeLabel(i.severidade))}</span>
          </div>
          <div class="integridade-meta">
            ${i.dia ? `<span>Dia: ${this._esc(this._diaLabel(i.dia))}</span>` : ""}
            ${i.pedidoId ? `<span>Pedido ID: ${this._esc(i.pedidoId)}</span>` : ""}
          </div>
          ${i.acaoSugerida ? `<div class="integridade-acao"><b>Ação sugerida:</b> ${this._esc(i.acaoSugerida)}</div>` : ""}
          ${i.detalhes ? `<details class="integridade-json"><summary>Ver detalhes técnicos</summary><pre>${this._esc(JSON.stringify(i.detalhes, null, 2))}</pre></details>` : ""}
        </div>`;
    },

    // ============================================================
    // REGRAS DE PEDIDOS
    // ============================================================
    _decidirProducao(p) {
      if (!p) return { conta: false, motivo: "Pedido vazio." };
      if (!this._dia(p)) return { conta: false, motivo: "Pedido sem dia." };
      if (!this._nome(p) && !this._colabId(p)) return { conta: false, motivo: "Pedido sem colaborador." };

      const status = this._status(p);
      if (this._statusBloqueiaProducao(status) && !this._isTravamentoAutomaticoConfirmado(p)) {
        return { conta: false, motivo: `Status bloqueia produção: ${status || "—"}.` };
      }

      if (this._isTravamentoAutomaticoConfirmado(p)) return { conta: true, motivo: "Travamento automático confirmado." };
      if (["confirmado", "extra", "aprovado"].includes(this._norm(status))) return { conta: true, motivo: `Status produtivo: ${status}.` };
      if (this._isTrue(this._pick(p, "Confirmado", "confirmado"))) return { conta: true, motivo: "Confirmado=true." };

      return { conta: false, motivo: `Status não produtivo: ${status || "—"}.` };
    },

    _deduplicarPedidos(lista) {
      const gruposMap = new Map();
      const especiais = [];
      for (const p of lista || []) {
        if (this._isEspecialQueNaoDeduplica(p)) {
          especiais.push(p);
          continue;
        }
        const chave = `${this._norm(this._dia(p))}|${this._keyPessoa(p)}`;
        if (!gruposMap.has(chave)) gruposMap.set(chave, []);
        gruposMap.get(chave).push(p);
      }

      const mantidos = [...especiais];
      const descartadosDuplicidade = [];
      const grupos = [];

      for (const [chave, todos] of gruposMap.entries()) {
        const preferido = this._escolherPreferido(todos);
        mantidos.push(preferido);
        todos.forEach(p => { if (p !== preferido) descartadosDuplicidade.push(p); });
        grupos.push({ chave, preferido, todos });
      }

      return { mantidos, descartadosDuplicidade, grupos };
    },

    _escolherPreferido(lista) {
      const arr = [...(lista || [])];
      arr.sort((a, b) => this._pontuacaoPedido(b) - this._pontuacaoPedido(a));
      return arr[0];
    },

    _pontuacaoPedido(p) {
      let score = 0;
      const decisao = this._decidirProducao(p);
      if (decisao.conta) score += 1000000000;
      if (this._isTravamentoAutomaticoConfirmado(p)) score += 500000000;
      if (this._pedidoAusencia(p)) score += 200000000;
      const t = Date.parse(this._pick(p, "Modified", "Data_Hora", "Created") || "") || 0;
      score += Math.floor(t / 1000);
      score += this._num(this._id(p)) / 1000000;
      return score;
    },

    _statusBloqueiaProducao(status) {
      return [
        "cancelado", "afastado", "atestado", "ferias", "férias", "licenca", "licença",
        "nao vai almocar", "não vai almoçar", "nao_vai_almocar", "bloqueado", "ausente"
      ].includes(this._norm(status));
    },

    _isTravamentoAutomaticoConfirmado(p) {
      const status = this._norm(this._status(p));
      const origem = this._norm(this._origem(p));
      return status === "travado" && origem.includes("travamento");
    },

    _pedidoAusencia(p) {
      const origem = this._norm(this._origem(p));
      const status = this._norm(this._status(p));
      const obs = this._norm(this._pick(p, "Observacao", "Observação"));
      return origem.includes("ausencia") || ["ferias", "férias", "afastado", "atestado", "ausente", "licenca", "licença"].includes(status) || obs.includes("ausenciaid");
    },

    _isEspecialQueNaoDeduplica(p) {
      const cat = this._categoriaPedido(p);
      return ["guarda", "investigador", "extra", "prestador", "visitante", "terceiro"].includes(cat);
    },

    _categoriaPedido(p) {
      const origem = this._norm(this._origem(p));
      const tipo = this._norm(this._pick(p, "tipo", "Tipo"));
      const nome = this._norm(this._nome(p));
      const base = tipo || origem;
      if (base.includes("investigador") || nome.includes("investigador")) return "investigador";
      if (base.includes("guarda") || nome.includes("guarda")) return "guarda";
      if (base.includes("prestador")) return "prestador";
      if (base.includes("visitante")) return "visitante";
      if (base.includes("terceiro")) return "terceiro";
      if (base.includes("extra") || nome.includes("refeicao extra")) return "extra";
      return "colaborador";
    },

    _resumoPedido(p) {
      return {
        id: this._id(p),
        colaboradorId: this._colabId(p),
        nome: this._nome(p),
        dia: this._dia(p),
        opcao: this._opcao(p) || "semOpcao",
        status: this._status(p),
        confirmado: this._pick(p, "Confirmado", "confirmado"),
        origem: this._origem(p),
        dataHora: this._pick(p, "Data_Hora", "DataHora"),
        modified: this._pick(p, "Modified", "modified"),
        contaProducao: this._decidirProducao(p).conta
      };
    },

    // ============================================================
    // HELPERS
    // ============================================================
    _diasParaVerificar(semanaId, options = {}) {
      const datas = this._datasSemana(semanaId);
      const hoje = this._dateISO(new Date());
      return ORDEM_DIAS.map((dia, idx) => ({ dia, data: datas[idx] || "" }))
        .filter(d => !options.checarSomenteAteHoje || !d.data || d.data <= hoje);
    },

    _datasSemana(semanaId) {
      try {
        if (SP.getWeekDates) return SP.getWeekDates(semanaId).map(d => this._dateISO(d));
      } catch (_) {}
      const [year, week] = String(semanaId).split("-W").map(Number);
      const jan4 = new Date(year, 0, 4);
      const start = new Date(jan4);
      start.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1 + (week - 1) * 7);
      return Array.from({ length: 5 }, (_, i) => {
        const d = new Date(start);
        d.setDate(start.getDate() + i);
        return this._dateISO(d);
      });
    },

    _semanaAtual() {
      return global.AdminState?.getSemanaId?.() || global.AdminUtils?.getSemanaId?.() || global.SP?.getSemanaId?.(new Date()) || "";
    },

    _colaboradorAtivo(c) {
      const ativo = this._pick(c, "Ativo", "ativo");
      if (ativo === "" || ativo === null || ativo === undefined) return true;
      return this._isTrue(ativo);
    },

    _keyPessoa(obj) {
      const id = this._colabId(obj);
      if (id) return `id:${String(id)}`;
      return `nome:${this._norm(this._nome(obj))}`;
    },

    _id(obj) { return this._pick(obj, "id", "ID") || ""; },
    _colabId(obj) { return this._pick(obj, "Colaborador_id", "ColaboradorId", "colaboradorId") || ""; },
    _nome(obj) { return this._pick(obj, "Colaborador_nome", "Colaborador", "Nome", "Title") || ""; },
    _dia(obj) { return this._pick(obj, "Dia", "dia") || ""; },
    _opcao(obj) { return this._pick(obj, "Opcao", "opcao") || ""; },
    _status(obj) { return this._pick(obj, "Status", "status") || ""; },
    _origem(obj) { return this._pick(obj, "Origem", "origem", "tipo", "Tipo") || ""; },

    _pick(obj, ...keys) {
      for (const k of keys) {
        let v;
        try { v = global.SP?.pick ? SP.pick(obj, k) : obj?.[k]; }
        catch (_) { v = obj?.[k]; }
        if (v !== undefined && v !== null && String(v) !== "") return v;
      }
      return "";
    },

    _isTrue(v) {
      if (global.SP?.isTrue) return SP.isTrue(v);
      if (v === true || v === 1) return true;
      const s = String(v ?? "").toLowerCase().trim();
      return ["sim", "true", "yes", "1"].includes(s);
    },

    _norm(v) {
      return String(v ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    },

    _num(v) {
      const n = Number(v || 0);
      return Number.isFinite(n) ? n : 0;
    },

    _dateISO(v) {
      if (!v) return "";
      if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
      const d = v instanceof Date ? v : new Date(v);
      if (isNaN(d)) return "";
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, "0");
      const dia = String(d.getDate()).padStart(2, "0");
      return `${y}-${m}-${dia}`;
    },

    _fmtDataHora(v) {
      try { return new Date(v).toLocaleString("pt-BR", { dateStyle: "short", timeStyle: "short" }); }
      catch (_) { return String(v || ""); }
    },

    _diaLabel(dia) { return DIA_LABEL[this._norm(dia)] || dia || "—"; },

    _resumoCurto(r) {
      return `T:${this._num(r.total)} · P:${this._num(r.principal)} · L:${this._num(r.light)} · C:${this._num(r.carne)} · M:${this._num(r.massa)}`;
    },

    _htmlDiferencas(d) {
      return ["total", "principal", "light", "carne", "massa"]
        .filter(k => this._num(d[k]) !== 0)
        .map(k => `<span class="diff-pill">${this._esc(k)} ${this._sinal(d[k])}</span>`)
        .join(" ") || `<span class="diff-ok">bate</span>`;
    },

    _sinal(v) {
      const n = this._num(v);
      return n > 0 ? `+${n}` : String(n);
    },

    _tituloTipo(tipo) {
      const mapa = {
        "diferenca-referencia-operacional": "Diferença contra referência Luana",
        "pedido-incompleto": "Pedidos incompletos",
        "pedido-bloqueado-confirmado": "Confirmado com status bloqueado",
        "duplicidade-pedido": "Pedidos duplicados",
        "checkin-sem-dados-dia": "Check-in sem dados no dia"
      };
      return mapa[tipo] || tipo || "Outras irregularidades";
    },

    _severidadeLabel(sev) {
      return { critico: "Crítico", alerta: "Alerta", atencao: "Atenção", info: "Info" }[sev] || sev || "Info";
    },

    _issue(base) {
      const raw = `${base.tipo || "issue"}-${base.pedidoId || base.nome || "item"}-${base.dia || ""}`;
      return { id: raw.replace(/[^a-zA-Z0-9_-]/g, "-"), ...base };
    },

    _esc(v) {
      if (global.AdminUtils?.esc) return AdminUtils.esc(v);
      return String(v ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
    },

    _toast(msg, tipo = "info") {
      if (global.AdminUtils?.toast) AdminUtils.toast(msg, tipo);
      else console.log(`[${tipo}] ${msg}`);
    },

    _removerModalAntigo() {
      document.getElementById("adminIntegridadeModalOverlay")?.remove();
    },

    // ============================================================
    // CSS
    // ============================================================
    _ensureStyle() {
      let style = document.getElementById("adminIntegridadeCSS");
      if (style) style.remove();
      style = document.createElement("style");
      style.id = "adminIntegridadeCSS";
      style.textContent = `
        .admin-integridade-card{margin-top:.8rem;border-radius:14px;border:1px solid rgba(255,255,255,.09);background:rgba(255,255,255,.035);padding:.9rem;display:flex;flex-direction:column;gap:.7rem}
        html.admin-integridade-modal-open,body.admin-integridade-modal-open{overflow:hidden!important;height:100%!important}
        #adminIntegridadeModalOverlay{position:fixed;inset:0;z-index:999999;background:rgba(2,8,20,.78);backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);display:flex;align-items:center;justify-content:center;padding:18px;overflow:hidden}
        .admin-integridade-modal-shell{width:min(1180px,calc(100vw - 28px));height:min(92vh,920px);max-height:calc(100vh - 28px);border-radius:22px;border:1px solid rgba(80,150,255,.25);background:#06142a;box-shadow:0 30px 90px rgba(0,0,0,.72),inset 0 1px 0 rgba(255,255,255,.06);display:flex;flex-direction:column;overflow:hidden}
        .admin-integridade-modal-header{flex:0 0 auto;display:flex;align-items:flex-start;justify-content:space-between;gap:1rem;padding:1.05rem 1.25rem;border-bottom:1px solid rgba(255,255,255,.08);background:rgba(7,20,42,.96)}
        .admin-integridade-modal-title{font-family:"Barlow Condensed",sans-serif;font-size:1.35rem;font-weight:900;text-transform:uppercase;color:#fff;letter-spacing:.04em;line-height:1.1}
        .admin-integridade-modal-sub{font-size:.72rem;color:rgba(143,170,210,.68);margin-top:.28rem;line-height:1.35}
        .admin-integridade-modal-close{width:38px;height:38px;border-radius:11px;border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.06);color:#d8e7ff;font-size:1.35rem;font-weight:900;cursor:pointer;line-height:1;display:flex;align-items:center;justify-content:center;flex:0 0 auto}
        .admin-integridade-modal-close:hover{background:rgba(192,40,28,.30);border-color:rgba(255,120,120,.26);color:#fff}
        .admin-integridade-modal-toolbar{flex:0 0 auto;display:flex;gap:.6rem;justify-content:flex-end;align-items:center;padding:.8rem 1.25rem;border-bottom:1px solid rgba(255,255,255,.07);background:rgba(8,24,48,.92)}
        .admin-integridade-modal-body{flex:1 1 auto;min-height:0;overflow-y:scroll;overflow-x:hidden;overscroll-behavior:contain;padding:1rem 1.25rem 1.3rem;outline:none;scrollbar-gutter:stable}
        .integridade-modal-content{display:flex;flex-direction:column;gap:1rem;min-height:min-content}
        .admin-integridade-card.ok{border-color:rgba(64,208,144,.24);background:rgba(64,208,144,.055)}
        .admin-integridade-card.atencao{border-color:rgba(255,190,60,.32);background:rgba(255,180,0,.055)}
        .admin-integridade-card.critico{border-color:rgba(255,90,90,.36);background:rgba(192,40,28,.105)}
        .admin-integridade-head{display:flex;align-items:flex-start;justify-content:space-between;gap:.8rem;flex-wrap:wrap}
        .admin-integridade-title{font-weight:900;color:#fff;font-size:.9rem;letter-spacing:.02em}
        .admin-integridade-sub{font-size:.68rem;color:rgba(143,170,210,.66);margin-top:.2rem}
        .admin-integridade-actions,.integridade-painel-actions{display:flex;gap:.5rem;flex-wrap:wrap;justify-content:flex-end}
        .admin-integridade-btn{white-space:nowrap;padding:.48rem .75rem;font-size:.74rem}
        .admin-integridade-status{font-size:.8rem;color:rgba(220,235,255,.9);line-height:1.4}
        .integridade-diff-preview{font-size:.78rem;color:#ffd67a;border-left:2px solid rgba(255,190,60,.38);padding-left:.55rem;line-height:1.35}
        .integridade-diff-preview.ok{color:#78e6b0;border-left-color:rgba(64,208,144,.35)}
        .admin-integridade-pills{display:flex;gap:.4rem;flex-wrap:wrap}
        .admin-integridade-pill{font-size:.64rem;font-weight:900;text-transform:uppercase;border-radius:999px;padding:.32rem .55rem;border:1px solid rgba(255,255,255,.09);background:rgba(255,255,255,.045)}
        .admin-integridade-pill.critico{color:#ff9a90;border-color:rgba(255,90,90,.22);background:rgba(192,40,28,.10)}
        .admin-integridade-pill.alerta{color:#ffc46d;border-color:rgba(255,170,40,.22);background:rgba(255,150,0,.08)}
        .admin-integridade-pill.atencao{color:#ffe08a;border-color:rgba(255,220,80,.20);background:rgba(255,220,80,.06)}
        .admin-integridade-pill.info{color:#9bd2ff;border-color:rgba(80,150,255,.20);background:rgba(80,150,255,.06)}
        .integridade-preview-list{display:flex;flex-direction:column;gap:.42rem}
        .integridade-preview-row{border-radius:10px;padding:.55rem .65rem;border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.035)}
        .integridade-preview-row.critico{border-color:rgba(255,90,90,.22)}
        .integridade-preview-row.alerta,.integridade-preview-row.atencao{border-color:rgba(255,190,60,.22)}
        .integridade-preview-main{font-size:.78rem;font-weight:900;color:#fff}
        .integridade-preview-sub{font-size:.7rem;color:rgba(200,220,255,.68);margin-top:.12rem;line-height:1.35}
        .admin-integridade-error{font-size:.72rem;color:#ffaaa0;line-height:1.45}
        #adminIntegridadePainelDetalhes{margin-top:.8rem}
        .integridade-painel-inline{border-radius:18px;border:1px solid rgba(80,150,255,.18);background:rgba(5,15,34,.96);padding:1rem;display:flex;flex-direction:column;gap:1rem;max-width:100%;overflow:visible}
        .integridade-painel-head{display:flex;align-items:flex-start;justify-content:space-between;gap:1rem;flex-wrap:wrap;border-bottom:1px solid rgba(255,255,255,.08);padding-bottom:.85rem}
        .integridade-painel-title{font-family:"Barlow Condensed",sans-serif;font-size:1.25rem;font-weight:900;text-transform:uppercase;color:#fff;letter-spacing:.04em}
        .integridade-painel-sub,.integridade-nota{font-size:.72rem;color:rgba(143,170,210,.68);line-height:1.45}
        .integridade-section{border:1px solid rgba(255,255,255,.08);border-radius:15px;background:rgba(255,255,255,.025);padding:.9rem;display:flex;flex-direction:column;gap:.75rem;overflow:visible}
        .integridade-section-title{font-weight:900;color:#fff;font-size:.95rem}
        .integridade-table-wrap{max-width:100%;overflow:auto}
        .integridade-table td,.integridade-table th{white-space:nowrap}
        .linha-diff td{background:rgba(255,180,0,.045)!important}
        .diff-pill{display:inline-block;margin:.12rem .18rem .12rem 0;border:1px solid rgba(255,190,60,.28);background:rgba(255,180,0,.08);color:#ffd67a;border-radius:999px;padding:.18rem .42rem;font-size:.68rem;font-weight:800}
        .diff-ok{color:#78e6b0;font-weight:900}
        .integridade-details{border:1px solid rgba(255,255,255,.08);border-radius:13px;background:rgba(255,255,255,.028);padding:.6rem;overflow:visible}
        .integridade-details summary{cursor:pointer;color:#fff;font-weight:900;display:flex;align-items:center;gap:.5rem;justify-content:space-between}
        .integridade-issue{margin:.65rem 0 0;border-radius:13px;border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.04);padding:.75rem;display:flex;flex-direction:column;gap:.5rem}
        .integridade-issue.critico{border-color:rgba(255,90,90,.28);background:rgba(192,40,28,.075)}
        .integridade-issue.alerta,.integridade-issue.atencao{border-color:rgba(255,190,60,.24);background:rgba(255,180,0,.055)}
        .integridade-issue-top{display:flex;align-items:flex-start;justify-content:space-between;gap:.8rem}
        .integridade-issue-nome{font-weight:900;color:#fff;font-size:.88rem}
        .integridade-issue-msg{font-size:.74rem;color:rgba(220,235,255,.78);line-height:1.42;margin-top:.18rem}
        .integridade-sev{font-size:.6rem;font-weight:900;text-transform:uppercase;border-radius:999px;padding:.25rem .45rem;border:1px solid rgba(255,255,255,.12);white-space:nowrap}
        .integridade-sev.critico{color:#ff9a90;border-color:rgba(255,90,90,.25);background:rgba(192,40,28,.12)}
        .integridade-sev.alerta,.integridade-sev.atencao{color:#ffd67a;border-color:rgba(255,190,60,.24);background:rgba(255,180,0,.08)}
        .integridade-sev.info{color:#9bd2ff;border-color:rgba(80,150,255,.24);background:rgba(80,150,255,.08)}
        .integridade-meta{display:flex;gap:.45rem;flex-wrap:wrap;font-size:.66rem;color:rgba(143,170,210,.72)}
        .integridade-meta span{border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.035);border-radius:999px;padding:.2rem .4rem}
        .integridade-acao{font-size:.72rem;color:rgba(220,235,255,.78);line-height:1.4;border-left:2px solid rgba(97,184,255,.35);padding-left:.55rem}
        .integridade-json summary{cursor:pointer;color:#9bd2ff;font-size:.72rem;font-weight:900;display:block}
        .integridade-json pre{margin-top:.5rem;white-space:pre-wrap;word-break:break-word;max-height:320px;overflow:auto;border-radius:12px;background:rgba(0,0,0,.25);border:1px solid rgba(255,255,255,.07);padding:.75rem;color:rgba(220,235,255,.75);font-size:.68rem}
        .integridade-opcao-bloco{margin-top:.65rem;border-top:1px solid rgba(255,255,255,.06);padding-top:.65rem}
        .integridade-opcao-title{font-size:.76rem;font-weight:900;color:#9bd2ff;text-transform:uppercase;margin-bottom:.45rem}
        .integridade-lista-pedidos{display:flex;flex-direction:column;gap:.4rem}
        .integridade-pedido-row{display:flex;align-items:flex-start;justify-content:space-between;gap:.7rem;border:1px solid rgba(255,255,255,.08);border-radius:11px;background:rgba(255,255,255,.035);padding:.55rem .65rem}
        .integridade-pedido-row.incluido{border-color:rgba(64,208,144,.13)}
        .integridade-pedido-row.excluido{border-color:rgba(255,90,90,.12);opacity:.92}
        .integridade-pedido-sub{font-size:.68rem;color:rgba(143,170,210,.72);margin-top:.12rem;line-height:1.35}
        .integridade-pedido-motivo{font-size:.67rem;color:rgba(220,235,255,.62);margin-top:.16rem;line-height:1.35}
        .integridade-empty{border-radius:13px;border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.035);padding:.85rem;text-align:center;color:rgba(220,235,255,.75)}
        .integridade-empty.ok{border-color:rgba(64,208,144,.24);background:rgba(64,208,144,.06);color:#78e6b0}
        @media(max-width:760px){.admin-integridade-head,.integridade-painel-head{flex-direction:column}.admin-integridade-actions,.integridade-painel-actions{justify-content:flex-start}.integridade-table td,.integridade-table th{font-size:.72rem}.integridade-pedido-row{flex-direction:column}}
      `;
      document.head.appendChild(style);
    }
  };

  function tentarInstalar() {
    if (AdminIntegridade.instalarNoDashboard()) return;
    setTimeout(tentarInstalar, 300);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", tentarInstalar);
  } else {
    tentarInstalar();
  }
})(window);
