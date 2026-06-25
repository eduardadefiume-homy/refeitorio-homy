// ============================================================
// admin-integridade.js — Integridade dos Dados · Admin Homy
// v: integridade-dashboard-readonly-20260625-fase2
//
// Objetivo:
// - Ler dados reais do SharePoint via SP/sharepoint.js.
// - Usar HomyRefeitorioRegras quando disponível.
// - Exibir irregularidades no Dashboard > Alertas.
// - NÃO grava no SharePoint nesta fase.
// - NÃO substitui os fluxos atuais do Admin/Cozinha.
// ============================================================
(function (global) {
  "use strict";

  const AdminIntegridade = global.AdminIntegridade = {
    _ultimoResultado: null,
    _loading: false,
    _instalado: false,
    _config: {
      containerId: "dashAlertas",
      checarSomenteAteHoje: true,
      maxItensResumo: 5,
      maxDetalhesPorGrupo: 60,
      cacheMs: 15000
    },
    _cache: new Map(),

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
      const irregularidades = [];

      irregularidades.push(...this._verificarPedidosQuebrados(dados, semanaId));
      irregularidades.push(...this._verificarPedidosBloqueadosConfirmados(dados, semanaId));
      irregularidades.push(...this._verificarDuplicidades(dados, semanaId));

      for (const diaInfo of dias) {
        irregularidades.push(...this._verificarAusenciasVigentesComPedidoConfirmado(dados, semanaId, diaInfo));
        irregularidades.push(...this._verificarRetornosEsperados(dados, semanaId, diaInfo));
      }

      const resumoProducaoPorDia = {};
      const resumoRetiradaPorDia = {};
      for (const diaInfo of dias) {
        resumoProducaoPorDia[diaInfo.dia] = this._calcularResumoDia(dados.pedidos, diaInfo.dia);
        resumoRetiradaPorDia[diaInfo.dia] = this._calcularResumoRetiradasDia(dados.pedidos, dados.checkins, diaInfo.dia);
      }

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
        diasVerificados: dias,
        geradoEm: new Date().toISOString(),
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
        irregularidades,
        porSeveridade,
        status: irregularidades.some(i => i.severidade === "critico") ? "critico" :
          irregularidades.some(i => i.severidade === "alerta" || i.severidade === "atencao") ? "atencao" : "ok"
      };
    },

    renderResumo(resultado, options = {}) {
      const container = document.getElementById(options.containerId || this._config.containerId);
      if (!container || !resultado) return;

      const id = "adminIntegridadeResumo";
      let box = document.getElementById(id);
      if (!box) {
        box = document.createElement("div");
        box.id = id;
        container.appendChild(box);
      }

      const total = resultado.irregularidades.length;
      const criticos = resultado.porSeveridade.critico || 0;
      const alertas = resultado.porSeveridade.alerta || 0;
      const atencoes = resultado.porSeveridade.atencao || 0;
      const infos = resultado.porSeveridade.info || 0;

      const classe = resultado.status === "ok" ? "ok" : resultado.status === "critico" ? "critico" : "atencao";
      const icone = resultado.status === "ok" ? "✅" : resultado.status === "critico" ? "🔴" : "⚠️";
      const titulo = resultado.status === "ok"
        ? "Integridade dos dados sem irregularidades críticas."
        : `${total} irregularidade(s) de integridade encontradas.`;

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
              <div class="admin-integridade-sub">Semana ${this._esc(resultado.semanaId)} · ${this._fmtDataHora(resultado.geradoEm)}</div>
            </div>
            <button class="btn-secondary admin-integridade-btn" onclick="AdminIntegridade.abrirDetalhes()">Ver detalhes</button>
          </div>

          <div class="admin-integridade-status">${this._esc(titulo)}</div>

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
      const body = document.getElementById("adminIntegridadeModalBody");
      const title = document.getElementById("adminIntegridadeModalTitle");
      if (!overlay || !body || !title) return;

      const r = this._ultimoResultado;
      if (!r) {
        body.innerHTML = `<div class="integridade-empty">Nenhum diagnóstico carregado.</div>`;
      } else {
        title.textContent = `Integridade dos Dados · ${r.semanaId}`;
        body.innerHTML = this._htmlDetalhes(r);
      }

      overlay.classList.add("open");
    },

    fecharDetalhes() {
      const overlay = document.getElementById("adminIntegridadeModalOverlay");
      if (overlay) overlay.classList.remove("open");
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
      const pedidosPromise = this._safeList(() => SP.getPedidos?.(semanaId, { reparar: false, force: true }), "Pedidos");
      const ausenciasPromise = this._safeList(() => SP.getAusencias?.(false), "Ausencias do Refeitorio");
      const colaboradoresPromise = this._safeList(() => SP.getTodosColaboradores?.(), "Colaboradores");
      const extrasPromise = this._safeList(() => SP.getExtras?.(semanaId), "Extras");
      const cardapioPromise = this._safeList(() => SP.getCardapio?.(semanaId), "Cardapio");
      const checkinsPromise = this._safeList(async () => {
        if (typeof SP.getItems === "function") return SP.getItems("CheckIn");
        if (typeof SP.getCheckIn === "function") {
          const dias = ["segunda", "terca", "quarta", "quinta", "sexta"];
          const listas = await Promise.all(dias.map(dia => SP.getCheckIn(semanaId, dia).catch(() => [])));
          return listas.flat();
        }
        return [];
      }, "CheckIn");

      const [pedidos, ausencias, colaboradores, extras, cardapio, checkins] = await Promise.all([
        pedidosPromise,
        ausenciasPromise,
        colaboradoresPromise,
        extrasPromise,
        cardapioPromise,
        checkinsPromise
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
          pedido: p,
          pedidoId: this._id(p),
          mensagem: `Pedido está Confirmado=true, mas Status=${status || "—"} bloqueia produção.`,
          acaoSugerida: "Ajustar Confirmado=false ou reprocessar a ausência/cancelamento para não afetar contagem."
        }));
      }
      return out;
    },

    _verificarDuplicidades(dados, semanaId) {
      const grupos = new Map();
      const out = [];

      for (const p of dados.pedidos || []) {
        if (this._isEspecialQueNaoDeduplica(p)) continue;
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
        const dia = this._dia(preferido);

        out.push(this._issue({
          tipo: "duplicidade-pedido",
          severidade: validos.length > 1 ? "critico" : "alerta",
          nome: this._nome(preferido),
          semanaId,
          dia,
          pedido: preferido,
          pedidoId: this._id(preferido),
          mensagem: `${grupo.length} pedidos encontrados para o mesmo colaborador/dia. Pedido preferido: ID ${this._id(preferido) || "sem ID"}.`,
          acaoSugerida: "Manter o pedido preferido e cancelar/regularizar os duplicados obsoletos.",
          detalhes: {
            chave: key,
            pedidoPreferido: this._resumoPedido(preferido),
            duplicados: ordenados.slice(1).map(p => this._resumoPedido(p)),
            todos: ordenados.map(p => this._resumoPedido(p))
          }
        }));
      }

      return out;
    },

    _verificarAusenciasVigentesComPedidoConfirmado(dados, semanaId, diaInfo) {
      const out = [];
      const dataRef = diaInfo.data;
      const pedidosDia = (dados.pedidos || []).filter(p => this._norm(this._dia(p)) === this._norm(diaInfo.dia));

      for (const a of dados.ausencias || []) {
        if (!this._ausenciaAtiva(a)) continue;
        if (!this._ausenciaCobreData(a, dataRef)) continue;

        const keyAus = this._colabKey({
          Colaborador_id: this._pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId"),
          Colaborador_nome: this._pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title")
        });
        if (!keyAus) continue;

        const pedidosColab = pedidosDia.filter(p => this._colabKey(p) === keyAus && !this._isEspecialQueNaoDeduplica(p));
        const confirmados = pedidosColab.filter(p => this._pedidoContaProducao(p));
        if (!confirmados.length) continue;

        const motivo = this._formatarMotivoAusencia(this._pick(a, "Motivo", "motivo", "Status") || "Ausente");
        for (const p of confirmados) {
          out.push(this._issue({
            tipo: "ausencia-vigente-com-pedido-confirmado",
            severidade: "critico",
            nome: this._nome(p) || this._nome(a),
            semanaId,
            dia: diaInfo.dia,
            data: dataRef,
            pedido: p,
            ausencia: a,
            pedidoId: this._id(p),
            ausenciaId: this._id(a),
            mensagem: `Ausência vigente (${motivo}) cobre ${this._br(dataRef)}, mas existe pedido confirmado na produção.`,
            acaoSugerida: "Reprocessar ausência para o dia ou alterar o pedido para status de ausência."
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
        if (this._isEspecialQueNaoDeduplica(p)) continue;
        const key = this._colabKey(p);
        if (!key) continue;
        const atual = pedidosPorColab.get(key);
        if (!atual || this._compararPreferido(p, atual) < 0) pedidosPorColab.set(key, p);
      }

      for (const c of dados.colaboradores || []) {
        const key = this._colabKey(c);
        if (!key) continue;

        const ausenciaHoje = this._ausenciaVigenteParaKey(dados.ausencias, key, dataRef);
        if (ausenciaHoje) continue;

        const ausenciasEncerradas = (dados.ausencias || []).filter(a => {
          const keyAus = this._colabKey({
            Colaborador_id: this._pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId"),
            Colaborador_nome: this._pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title")
          });
          if (keyAus !== key) return false;
          if (!this._ausenciaAtiva(a)) return false;
          const fim = this._ausenciaFim(a);
          return !!fim && fim < dataRef;
        });
        if (!ausenciasEncerradas.length) continue;

        ausenciasEncerradas.sort((a, b) => String(this._ausenciaFim(b)).localeCompare(String(this._ausenciaFim(a))));
        const ultimaAus = ausenciasEncerradas[0];
        const pedido = pedidosPorColab.get(key);

        if (!pedido) {
          out.push(this._issue({
            tipo: "ausencia-encerrada-sem-retorno",
            severidade: "critico",
            nome: this._nome(c),
            semanaId,
            dia: diaInfo.dia,
            data: dataRef,
            colaborador: c,
            ausencia: ultimaAus,
            ausenciaId: this._id(ultimaAus),
            mensagem: `Ausência encerrou em ${this._br(this._ausenciaFim(ultimaAus))}, mas não existe pedido para ${this._br(dataRef)}.`,
            acaoSugerida: "Gerar retorno automático como Principal para o dia."
          }));
          continue;
        }

        if (this._pedidoAusencia(pedido) && !this._pedidoContaProducao(pedido)) {
          out.push(this._issue({
            tipo: "retorno-bloqueado-por-status-antigo",
            severidade: "critico",
            nome: this._nome(c) || this._nome(pedido),
            semanaId,
            dia: diaInfo.dia,
            data: dataRef,
            pedido,
            ausencia: ultimaAus,
            pedidoId: this._id(pedido),
            ausenciaId: this._id(ultimaAus),
            mensagem: `Ausência encerrou em ${this._br(this._ausenciaFim(ultimaAus))}, mas o pedido do dia continua como ${this._status(pedido) || "ausência"}.`,
            acaoSugerida: "Atualizar o pedido para Confirmado/Principal ou reprocessar retorno automático."
          }));
        }
      }

      return out;
    },

    _calcularResumoDia(pedidos, dia) {
      const lista = (pedidos || []).filter(p => this._norm(this._dia(p)) === this._norm(dia));
      const dedup = this._deduplicar(lista);
      const validos = dedup.filter(p => this._pedidoContaProducao(p));
      const resumo = this._resumoVazio();
      validos.forEach(p => this._somarOpcaoResumo(resumo, this._opcao(p) || "principal"));
      return resumo;
    },

    _calcularResumoRetiradasDia(pedidos, checkins, dia) {
      const resumo = this._resumoVazio();
      const diaNorm = this._norm(dia);
      const mapaPedidos = this._mapaPedidosPreferidosPorColaboradorDia(pedidos, dia);
      const vistos = new Set();

      (checkins || []).forEach(ch => {
        if (this._norm(this._dia(ch)) !== diaNorm) return;
        if (!this._checkinRetirou(ch)) return;

        const key = this._colabKey(ch);
        const idCheck = this._id(ch) || key || `${diaNorm}-${vistos.size + 1}`;
        const dedupeKey = key ? `${diaNorm}|${key}` : `${diaNorm}|checkin:${idCheck}`;
        if (vistos.has(dedupeKey)) return;
        vistos.add(dedupeKey);

        const pedido = key ? mapaPedidos.get(key) : null;
        const opcao = this._opcao(ch) || (pedido ? this._opcao(pedido) : "");
        this._somarOpcaoResumo(resumo, opcao || "sem_opcao");
      });

      return resumo;
    },

    _resumoVazio() {
      return { total: 0, principal: 0, light: 0, carne: 0, massa: 0, lanche: 0, outros: 0, semOpcao: 0 };
    },

    _somarOpcaoResumo(resumo, opcao) {
      const op = this._norm(opcao || "");
      resumo.total++;
      if (op === "principal") resumo.principal++;
      else if (op === "light") resumo.light++;
      else if (op === "carne") resumo.carne++;
      else if (op === "massa") resumo.massa++;
      else if (op === "lanche") resumo.lanche++;
      else if (!op || op === "sem_opcao" || op === "sem opcao") resumo.semOpcao++;
      else resumo.outros++;
    },

    _mapaPedidosPreferidosPorColaboradorDia(pedidos, dia) {
      const mapa = new Map();
      const diaNorm = this._norm(dia);
      const lista = (pedidos || []).filter(p => this._norm(this._dia(p)) === diaNorm);
      const dedup = this._deduplicar(lista);
      dedup.forEach(p => {
        const key = this._colabKey(p);
        if (!key) return;
        const atual = mapa.get(key);
        if (!atual || this._compararPreferido(p, atual) < 0) mapa.set(key, p);
      });
      return mapa;
    },

    _checkinRetirou(ch) {
      const retirou = this._pick(ch, "Retirou", "retirou", "Retirada", "retirada");
      if (retirou === null || retirou === undefined || retirou === "") return true;
      return this._isTrue(retirou);
    },

    // ============================================================
    // HTML DETALHES
    // ============================================================
    _htmlDetalhes(resultado) {
      const irregularidades = resultado.irregularidades || [];
      const porTipo = new Map();
      for (const item of irregularidades) {
        const tipo = item.tipo || "outros";
        if (!porTipo.has(tipo)) porTipo.set(tipo, []);
        porTipo.get(tipo).push(item);
      }

      const resumoDias = this._htmlResumoDias(resultado);

      if (!irregularidades.length) {
        return `
          <div class="integridade-modal-actions">
            <button class="btn-secondary" onclick="AdminIntegridade.copiarDiagnostico()">Copiar JSON</button>
            <button class="btn-secondary" onclick="AdminIntegridade.baixarDiagnostico()">Baixar JSON</button>
          </div>
          <div class="integridade-dia-grid">${resumoDias}</div>
          <div class="integridade-empty ok">✅ Nenhuma irregularidade encontrada nos dias verificados.</div>`;
      }

      const grupos = Array.from(porTipo.entries()).map(([tipo, itens]) => {
        const titulo = this._tituloTipo(tipo);
        const linhas = itens.slice(0, this._config.maxDetalhesPorGrupo).map(i => this._htmlIssue(i)).join("");
        const extra = itens.length > this._config.maxDetalhesPorGrupo
          ? `<div class="integridade-limite">+ ${itens.length - this._config.maxDetalhesPorGrupo} item(ns) ocultos neste grupo. Use baixar JSON para ver tudo.</div>`
          : "";
        return `
          <section class="integridade-grupo">
            <div class="integridade-grupo-head">
              <span>${this._esc(titulo)}</span>
              <span class="badge badge-blue">${itens.length}</span>
            </div>
            ${linhas}${extra}
          </section>`;
      }).join("");

      return `
        <div class="integridade-modal-actions">
          <button class="btn-secondary" onclick="AdminIntegridade.copiarDiagnostico()">Copiar JSON</button>
          <button class="btn-secondary" onclick="AdminIntegridade.baixarDiagnostico()">Baixar JSON</button>
        </div>
        <div class="integridade-dia-grid">${resumoDias}</div>
        ${grupos}`;
    },

    _htmlResumoDias(resultado) {
      const producao = resultado.resumoProducaoPorDia || {};
      const retiradas = resultado.resumoRetiradaPorDia || {};
      const ordem = ["segunda", "terca", "quarta", "quinta", "sexta"];
      const listaDias = ordem.filter(d => producao[d] || retiradas[d]);

      return listaDias.map(dia => {
        const p = producao[dia] || this._resumoVazio();
        const r = retiradas[dia] || this._resumoVazio();
        const temRetirada = this._num(r.total) > 0;
        const principal = temRetirada ? r : p;
        const labelPrincipal = temRetirada ? "Entregues / Check-in" : "Produção calculada";
        const extras = [];
        if (this._num(principal.lanche)) extras.push(`La:${this._num(principal.lanche)}`);
        if (this._num(principal.outros)) extras.push(`Outros:${this._num(principal.outros)}`);
        if (this._num(principal.semOpcao)) extras.push(`Sem opção:${this._num(principal.semOpcao)}`);
        const extraTxt = extras.length ? ` · ${extras.join(" · ")}` : "";

        return `
          <div class="integridade-dia-card">
            <div class="integridade-dia-nome">${this._esc(this._diaLabel(dia))}</div>
            <div class="integridade-dia-total">${this._num(principal.total)}</div>
            <div class="integridade-dia-sub">${this._esc(labelPrincipal)}</div>
            <div class="integridade-dia-sub">P:${this._num(principal.principal)} · L:${this._num(principal.light)} · C:${this._num(principal.carne)} · M:${this._num(principal.massa)}${this._esc(extraTxt)}</div>
            <div class="integridade-dia-mini">Produção: ${this._num(p.total)} · Check-in: ${this._num(r.total)}</div>
          </div>`;
      }).join("");
    },

    _htmlIssue(i) {
      const sev = i.severidade || "info";
      const pedido = i.pedidoId ? `<span>Pedido ID: <b>${this._esc(i.pedidoId)}</b></span>` : "";
      const ausencia = i.ausenciaId ? `<span>Ausência ID: <b>${this._esc(i.ausenciaId)}</b></span>` : "";
      const data = i.data ? `<span>Data: <b>${this._esc(this._br(i.data))}</b></span>` : "";
      const dia = i.dia ? `<span>Dia: <b>${this._esc(this._diaLabel(i.dia))}</b></span>` : "";
      const detalhes = i.detalhes ? `<details class="integridade-json"><summary>Ver detalhes técnicos</summary><pre>${this._esc(JSON.stringify(i.detalhes, null, 2))}</pre></details>` : "";

      return `
        <article class="integridade-issue ${this._esc(sev)}">
          <div class="integridade-issue-top">
            <div>
              <div class="integridade-issue-nome">${this._esc(i.nome || "Registro")}</div>
              <div class="integridade-issue-msg">${this._esc(i.mensagem || "Irregularidade encontrada.")}</div>
            </div>
            <span class="integridade-sev ${this._esc(sev)}">${this._esc(this._severidadeLabel(sev))}</span>
          </div>
          <div class="integridade-meta">${[dia, data, pedido, ausencia].filter(Boolean).join("")}</div>
          <div class="integridade-acao"><b>Ação sugerida:</b> ${this._esc(i.acaoSugerida || "Revisar manualmente antes de corrigir.")}</div>
          ${detalhes}
        </article>`;
    },

    // ============================================================
    // HELPERS DE REGRA
    // ============================================================
    _R() {
      return global.HomyRefeitorioRegras || null;
    },

    _pick(obj, ...keys) {
      const R = this._R();
      if (R?.pick) return R.pick(obj, ...keys);
      for (const k of keys) {
        if (obj && obj[k] !== undefined && obj[k] !== null) return obj[k];
      }
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
      const n = this._norm(v);
      return ["sim", "true", "yes", "1"].includes(n);
    },

    _id(obj) { return String(this._pick(obj, "id", "ID") || "").trim(); },
    _nome(obj) { return this._pick(obj, "Colaborador_nome", "colaborador_nome", "Colaborador", "Nome", "Title", "nome") || ""; },
    _colabId(obj) { return String(this._pick(obj, "Colaborador_id", "colaborador_id", "ColaboradorId", "colaboradorId", "Matricula", "Matrícula", "id", "ID") || "").trim(); },
    _dia(obj) { return this._pick(obj, "Dia", "dia") || ""; },
    _opcao(obj) { return this._pick(obj, "Opcao", "opcao", "Opção") || ""; },
    _status(obj) { return this._pick(obj, "Status", "status") || ""; },
    _origem(obj) { return this._pick(obj, "Origem", "origem", "tipo", "Tipo") || ""; },

    _colabKey(obj) {
      const R = this._R();
      if (R?.colaboradorKey) return R.colaboradorKey(obj);
      const id = this._colabId(obj);
      if (id) return `id:${id}`;
      const nome = this._norm(this._nome(obj));
      return nome ? `nome:${nome}` : "";
    },

    _statusBloqueiaProducao(status) {
      const n = this._norm(status);
      return [
        "cancelado", "bloqueado", "nao vai almocar", "nao_vai_almocar", "não vai almoçar",
        "ausente", "ferias", "férias", "afastado", "afastamento", "atestado",
        "licenca", "licença", "banco horas", "banco_horas", "homy office", "homy_office", "falta"
      ].includes(n);
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
      const R = this._R();
      if (R?.pedidoAusente) return R.pedidoAusente(p);
      const status = this._norm(this._status(p));
      const origem = this._norm(this._origem(p));
      return ["nao vai almocar", "nao_vai_almocar", "não vai almoçar", "ausente", "ferias", "férias", "afastado", "atestado", "licenca", "licença"].includes(status) || origem.includes("ausencia") || origem.includes("ausência");
    },

    _isTravamentoAutomaticoConfirmado(p) {
      const status = this._norm(this._status(p));
      const origem = this._norm(this._origem(p));
      return status === "travado" && origem.includes("travamento");
    },

    _isEspecialQueNaoDeduplica(p) {
      const R = this._R();
      if (R?.isPedidoAdicionalColaborador?.(p)) return true;
      if (R?.isExtra?.(p)) return true;

      const origem = this._norm(this._origem(p));
      const nome = this._norm(this._nome(p));
      const cid = this._norm(this._colabId(p));
      const obs = this._norm(this._pick(p, "Observacao", "Observação", "observacao", "Obs"));

      const adicional = origem.includes("segunda refeicao") || origem.includes("segunda refeição") ||
        origem.includes("transferencia para hoje") || origem.includes("transferência para hoje") ||
        origem.includes("refeicao adicional") || origem.includes("refeição adicional") ||
        obs.includes("adicionalid:") || obs.includes("colaboradorbaseid:") || cid.includes("-adicional-");

      const extra = origem.includes("extra") || origem.includes("guarda") || origem.includes("investigador") ||
        origem.includes("visitante") || origem.includes("prestador") || origem.includes("terceiro") ||
        origem.includes("fornecedor") || origem.includes("representante") || origem.includes("motorista") ||
        nome.includes("refeicao extra") || /^guarda\s*\d*$/.test(nome) || /^investigador(?:\s*\d+)?$/.test(nome) ||
        cid.startsWith("extra-");

      return adicional || extra;
    },

    _compararPreferido(a, b) {
      const R = this._R();
      if (R?.compararPedidoPreferido) return R.compararPedidoPreferido(a, b);
      const sa = this._scorePedido(a);
      const sb = this._scorePedido(b);
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
        const key = this._isEspecialQueNaoDeduplica(p)
          ? `especial|${this._dia(p)}|${this._id(p) || this._colabId(p) || this._nome(p)}|${this._origem(p)}|${this._opcao(p)}`
          : `colab|${this._dia(p)}|${this._colabKey(p)}`;
        const atual = mapa.get(key);
        if (!atual || this._compararPreferido(p, atual) < 0) mapa.set(key, p);
      }
      return Array.from(mapa.values());
    },

    _ausenciaAtiva(a) {
      const R = this._R();
      if (R?.ausenciaAtivaPorCampo) return R.ausenciaAtivaPorCampo(a);
      const status = this._norm(this._pick(a, "Status", "status", "Status_Ausencia", "statusAusencia") || "");
      if (["inativo", "cancelado", "cancelada", "duplicado inativado", "duplicidade inativada", "false", "nao", "não", "0"].includes(status)) return false;
      const ativo = this._pick(a, "Ativo", "ativo");
      if (ativo === null || ativo === undefined || ativo === "") return true;
      return this._isTrue(ativo);
    },

    _ausenciaInicio(a) {
      const R = this._R();
      if (R?.ausenciaInicioISO) return R.ausenciaInicioISO(a);
      return this._dataISO(this._pick(a, "Data_Inicio", "Inicio", "DataInicio", "Data"));
    },

    _ausenciaFim(a) {
      const R = this._R();
      if (R?.ausenciaFimISO) return R.ausenciaFimISO(a);
      return this._dataISO(this._pick(a, "Data_Fim", "Fim", "DataFim", "Data")) || this._ausenciaInicio(a);
    },

    _ausenciaCobreData(a, dataISO) {
      const ini = this._ausenciaInicio(a);
      const fim = this._ausenciaFim(a);
      return !!ini && !!fim && ini <= dataISO && fim >= dataISO;
    },

    _ausenciaVigenteParaKey(ausencias, key, dataISO) {
      const candidatas = (ausencias || []).filter(a => {
        if (!this._ausenciaAtiva(a)) return false;
        const keyAus = this._colabKey({
          Colaborador_id: this._pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId"),
          Colaborador_nome: this._pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title")
        });
        return keyAus === key && this._ausenciaCobreData(a, dataISO);
      });
      if (!candidatas.length) return null;
      candidatas.sort((a, b) => String(this._ausenciaFim(b)).localeCompare(String(this._ausenciaFim(a))) || Number(this._id(b) || 0) - Number(this._id(a) || 0));
      return candidatas[0];
    },

    _formatarMotivoAusencia(motivo) {
      const n = this._norm(motivo);
      if (["nao_vai_almocar", "nao vai almocar", "não vai almoçar", "ausente"].includes(n)) return "Não vai almoçar";
      if (["ferias", "férias"].includes(n)) return "Férias";
      if (["afastado", "afastamento"].includes(n)) return "Afastado";
      if (n === "atestado") return "Atestado";
      if (["licenca", "licença"].includes(n)) return "Licença";
      if (["banco_horas", "banco horas"].includes(n)) return "Banco de horas";
      if (["homy_office", "homy office"].includes(n)) return "Homy Office";
      return String(motivo || "Ausente");
    },

    _colaboradorAtivo(c) {
      const ativo = this._pick(c, "Ativo", "ativo");
      if (ativo === null || ativo === undefined || ativo === "") return true;
      return this._isTrue(ativo);
    },

    // ============================================================
    // DATAS / SEMANA
    // ============================================================
    _semanaAtual() {
      try { return global.AdminState?.getSemanaId?.() || global.SP?.getCurrentWeekId?.() || global.SP?.getSemanaId?.(new Date()) || ""; }
      catch (_) { return ""; }
    },

    _diasParaVerificar(semanaId, options = {}) {
      const nomes = ["segunda", "terca", "quarta", "quinta", "sexta"];
      const hojeISO = this._dataISO(new Date());
      return nomes.map(dia => ({ dia, data: this._dataPorSemanaDia(semanaId, dia) }))
        .filter(x => x.data)
        .filter(x => !options.checarSomenteAteHoje || x.data <= hojeISO || this._semanaTerminou(semanaId));
    },

    _semanaTerminou(semanaId) {
      const sexta = this._dataPorSemanaDia(semanaId, "sexta");
      const hoje = this._dataISO(new Date());
      return !!sexta && sexta < hoje;
    },

    _dataPorSemanaDia(semanaId, dia) {
      const R = this._R();
      if (R?.dataPorSemanaDia) return R.dataPorSemanaDia(semanaId, dia, global.SP || {});
      try { if (SP?.getDataRefBySemanaDia) return this._dataISO(SP.getDataRefBySemanaDia(semanaId, dia)); } catch (_) {}
      const idx = { segunda: 0, terca: 1, terça: 1, quarta: 2, quinta: 3, sexta: 4 }[this._norm(dia)];
      const datas = this._getWeekDates(semanaId);
      return idx === undefined || !datas[idx] ? "" : this._dataISO(datas[idx]);
    },

    _getWeekDates(semanaId) {
      try { if (SP?.getWeekDates) return SP.getWeekDates(semanaId); } catch (_) {}
      const m = String(semanaId || "").match(/^(\d{4})-W(\d{2})$/);
      if (!m) return [];
      const year = Number(m[1]);
      const week = Number(m[2]);
      const jan4 = new Date(year, 0, 4);
      const start = new Date(jan4);
      start.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1 + (week - 1) * 7);
      return Array.from({ length: 5 }, (_, i) => {
        const d = new Date(start);
        d.setDate(start.getDate() + i);
        return d;
      });
    },

    _dataISO(v) {
      const R = this._R();
      if (R?.dataISO) return R.dataISO(v);
      if (!v) return "";
      if (v instanceof Date && !isNaN(v)) {
        const y = v.getFullYear();
        const m = String(v.getMonth() + 1).padStart(2, "0");
        const d = String(v.getDate()).padStart(2, "0");
        return `${y}-${m}-${d}`;
      }
      const s = String(v || "");
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
      const d = new Date(s);
      if (isNaN(d)) return "";
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    },

    _br(iso) {
      const d = this._dataISO(iso);
      if (!d) return "—";
      const [a, m, dia] = d.split("-");
      return `${dia}/${m}/${a}`;
    },

    _fmtDataHora(v) {
      const d = new Date(v);
      if (isNaN(d)) return "—";
      return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")} ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
    },

    _diaLabel(dia) {
      const mapa = { segunda: "Segunda", terca: "Terça", terça: "Terça", quarta: "Quarta", quinta: "Quinta", sexta: "Sexta" };
      return mapa[this._norm(dia)] || String(dia || "—");
    },

    // ============================================================
    // UTIL / ISSUE
    // ============================================================
    _issue(base) {
      return {
        id: `${base.tipo || "issue"}-${base.pedidoId || base.ausenciaId || base.nome || "item"}-${base.dia || ""}`.replace(/[^a-zA-Z0-9_-]/g, "-"),
        ...base
      };
    },

    _resumoPedido(p) {
      return {
        id: this._id(p),
        nome: this._nome(p),
        dia: this._dia(p),
        opcao: this._opcao(p),
        status: this._status(p),
        confirmado: this._pick(p, "Confirmado", "confirmado"),
        origem: this._origem(p),
        dataHora: this._pick(p, "Data_Hora", "DataHora"),
        modified: this._pick(p, "Modified", "modified"),
        contaProducao: this._pedidoContaProducao(p)
      };
    },

    _tituloTipo(tipo) {
      const mapa = {
        "pedido-incompleto": "Pedidos incompletos",
        "pedido-bloqueado-confirmado": "Confirmado com status bloqueado",
        "duplicidade-pedido": "Pedidos duplicados",
        "ausencia-vigente-com-pedido-confirmado": "Ausências vigentes com pedido confirmado",
        "ausencia-encerrada-sem-retorno": "Ausências encerradas sem retorno",
        "retorno-bloqueado-por-status-antigo": "Retornos bloqueados por status antigo"
      };
      return mapa[tipo] || tipo || "Outras irregularidades";
    },

    _severidadeLabel(sev) {
      return { critico: "Crítico", alerta: "Alerta", atencao: "Atenção", info: "Info" }[sev] || sev || "Info";
    },

    _esc(v) {
      if (global.AdminUtils?.esc) return AdminUtils.esc(v);
      return String(v ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
    },

    _num(v) { return Number(v || 0) || 0; },

    _toast(msg, tipo = "info") {
      if (global.AdminUtils?.toast) AdminUtils.toast(msg, tipo);
      else console.log(`[${tipo}] ${msg}`);
    },

    // ============================================================
    // UI: MODAL / CSS
    // ============================================================
    _ensureModal() {
      if (document.getElementById("adminIntegridadeModalOverlay")) return;
      const overlay = document.createElement("div");
      overlay.id = "adminIntegridadeModalOverlay";
      overlay.className = "admin-integridade-modal-overlay";
      overlay.innerHTML = `
        <div class="admin-integridade-modal">
          <div class="admin-integridade-modal-head">
            <div>
              <div class="admin-integridade-modal-title" id="adminIntegridadeModalTitle">Integridade dos Dados</div>
              <div class="admin-integridade-modal-sub">Somente leitura · nenhuma correção automática nesta fase</div>
            </div>
            <button class="admin-integridade-modal-close" onclick="AdminIntegridade.fecharDetalhes()">×</button>
          </div>
          <div class="admin-integridade-modal-body" id="adminIntegridadeModalBody"></div>
        </div>`;
      overlay.addEventListener("click", ev => {
        if (ev.target === overlay) this.fecharDetalhes();
      });
      document.body.appendChild(overlay);
    },

    _ensureStyle() {
      if (document.getElementById("adminIntegridadeCSS")) return;
      const style = document.createElement("style");
      style.id = "adminIntegridadeCSS";
      style.textContent = `
        .admin-integridade-card{margin-top:.8rem;border-radius:14px;border:1px solid rgba(255,255,255,.09);background:rgba(255,255,255,.035);padding:.9rem;display:flex;flex-direction:column;gap:.7rem}
        .admin-integridade-card.ok{border-color:rgba(64,208,144,.24);background:rgba(64,208,144,.055)}
        .admin-integridade-card.atencao{border-color:rgba(255,190,60,.32);background:rgba(255,180,0,.055)}
        .admin-integridade-card.critico{border-color:rgba(255,90,90,.36);background:rgba(192,40,28,.105)}
        .admin-integridade-head{display:flex;align-items:flex-start;justify-content:space-between;gap:.8rem}
        .admin-integridade-title{font-weight:800;color:#fff;font-size:.9rem;letter-spacing:.02em}
        .admin-integridade-sub{font-size:.68rem;color:rgba(143,170,210,.66);margin-top:.2rem}
        .admin-integridade-status{font-size:.8rem;color:rgba(220,235,255,.9);line-height:1.4}
        .admin-integridade-pills{display:flex;gap:.4rem;flex-wrap:wrap}
        .admin-integridade-pill{font-size:.64rem;font-weight:800;letter-spacing:.04em;text-transform:uppercase;border-radius:999px;padding:.32rem .55rem;border:1px solid rgba(255,255,255,.09);background:rgba(255,255,255,.045)}
        .admin-integridade-pill.critico{color:#ff9a90;border-color:rgba(255,90,90,.22);background:rgba(192,40,28,.10)}
        .admin-integridade-pill.alerta{color:#ffc46d;border-color:rgba(255,170,40,.22);background:rgba(255,150,0,.08)}
        .admin-integridade-pill.atencao{color:#ffe08a;border-color:rgba(255,220,80,.20);background:rgba(255,220,80,.06)}
        .admin-integridade-pill.info{color:#9bd2ff;border-color:rgba(80,150,255,.20);background:rgba(80,150,255,.06)}
        .admin-integridade-btn{white-space:nowrap;padding:.48rem .75rem;font-size:.74rem}
        .integridade-preview-list{display:flex;flex-direction:column;gap:.42rem}
        .integridade-preview-row{border-radius:10px;padding:.55rem .65rem;border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.035)}
        .integridade-preview-row.critico{border-color:rgba(255,90,90,.22)}
        .integridade-preview-row.alerta,.integridade-preview-row.atencao{border-color:rgba(255,190,60,.22)}
        .integridade-preview-main{font-size:.78rem;font-weight:800;color:#fff}
        .integridade-preview-sub{font-size:.7rem;color:rgba(200,220,255,.68);margin-top:.12rem;line-height:1.35}
        .admin-integridade-error{font-size:.72rem;color:#ffaaa0;line-height:1.45}
        .admin-integridade-modal-overlay{position:fixed;inset:0;z-index:9998;background:rgba(3,8,20,.82);backdrop-filter:blur(8px);display:none;align-items:flex-start;justify-content:center;padding:1.2rem;overflow-y:auto;overscroll-behavior:contain;-webkit-overflow-scrolling:touch}
        .admin-integridade-modal-overlay.open{display:flex}
        .admin-integridade-modal{width:min(980px,96vw);max-height:calc(100vh - 2.4rem);min-height:0;overflow:hidden;border-radius:22px;border:1px solid rgba(255,255,255,.12);background:#081426;box-shadow:0 24px 80px rgba(0,0,0,.62);display:flex;flex-direction:column;margin:auto 0}
        .admin-integridade-modal-head{display:flex;align-items:flex-start;justify-content:space-between;gap:1rem;padding:1.1rem 1.25rem;border-bottom:1px solid rgba(255,255,255,.08)}
        .admin-integridade-modal-title{font-family:"Barlow Condensed",sans-serif;font-weight:800;font-size:1.4rem;letter-spacing:.04em;text-transform:uppercase;color:#fff}
        .admin-integridade-modal-sub{font-size:.72rem;color:rgba(143,170,210,.62);margin-top:.2rem}
        .admin-integridade-modal-close{width:36px;height:36px;border-radius:10px;border:1px solid rgba(255,255,255,.10);background:rgba(255,255,255,.055);color:#dce8ff;font-size:1.4rem;cursor:pointer}
        .admin-integridade-modal-body{overflow-y:auto;overflow-x:hidden;max-height:calc(100vh - 150px);min-height:0;padding:1rem 1.2rem 1.25rem;display:flex;flex-direction:column;gap:1rem;-webkit-overflow-scrolling:touch}
        .integridade-modal-actions{display:flex;align-items:center;justify-content:flex-end;gap:.6rem;flex-wrap:wrap}
        .integridade-dia-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:.55rem}
        .integridade-dia-card{border-radius:13px;border:1px solid rgba(80,150,255,.16);background:rgba(80,150,255,.055);padding:.75rem}
        .integridade-dia-nome{font-size:.7rem;text-transform:uppercase;letter-spacing:.08em;color:rgba(143,170,210,.72);font-weight:800}
        .integridade-dia-total{font-family:"Barlow Condensed",sans-serif;font-size:1.7rem;font-weight:800;color:#fff;margin-top:.2rem}
        .integridade-dia-sub{font-size:.68rem;color:rgba(200,220,255,.58);line-height:1.35}
        .integridade-dia-mini{font-size:.62rem;color:rgba(143,170,210,.48);margin-top:.28rem}
        .integridade-grupo{border:1px solid rgba(255,255,255,.08);border-radius:16px;overflow:hidden;background:rgba(255,255,255,.025)}
        .integridade-grupo-head{display:flex;align-items:center;justify-content:space-between;gap:.8rem;padding:.8rem .9rem;border-bottom:1px solid rgba(255,255,255,.07);font-weight:800;color:#fff}
        .integridade-issue{margin:.75rem;border-radius:14px;border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.04);padding:.85rem;display:flex;flex-direction:column;gap:.55rem}
        .integridade-issue.critico{border-color:rgba(255,90,90,.28);background:rgba(192,40,28,.075)}
        .integridade-issue.alerta,.integridade-issue.atencao{border-color:rgba(255,190,60,.24);background:rgba(255,180,0,.055)}
        .integridade-issue-top{display:flex;align-items:flex-start;justify-content:space-between;gap:.8rem}
        .integridade-issue-nome{font-weight:900;color:#fff;font-size:.9rem}
        .integridade-issue-msg{font-size:.76rem;color:rgba(220,235,255,.78);line-height:1.42;margin-top:.18rem}
        .integridade-sev{font-size:.62rem;font-weight:900;text-transform:uppercase;letter-spacing:.06em;border-radius:999px;padding:.28rem .48rem;border:1px solid rgba(255,255,255,.12);white-space:nowrap}
        .integridade-sev.critico{color:#ff9a90;border-color:rgba(255,90,90,.25);background:rgba(192,40,28,.12)}
        .integridade-sev.alerta,.integridade-sev.atencao{color:#ffd67a;border-color:rgba(255,190,60,.24);background:rgba(255,180,0,.08)}
        .integridade-meta{display:flex;gap:.5rem;flex-wrap:wrap;font-size:.68rem;color:rgba(143,170,210,.72)}
        .integridade-meta span{border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.035);border-radius:999px;padding:.24rem .45rem}
        .integridade-acao{font-size:.74rem;color:rgba(220,235,255,.78);line-height:1.4;border-left:2px solid rgba(97,184,255,.35);padding-left:.55rem}
        .integridade-json summary{cursor:pointer;color:#9bd2ff;font-size:.72rem;font-weight:800}
        .integridade-json pre{margin-top:.5rem;white-space:pre-wrap;word-break:break-word;max-height:260px;overflow:auto;border-radius:12px;background:rgba(0,0,0,.25);border:1px solid rgba(255,255,255,.07);padding:.75rem;color:rgba(220,235,255,.75);font-size:.68rem}
        .integridade-empty{border-radius:14px;border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.035);padding:1rem;text-align:center;color:rgba(220,235,255,.75)}
        .integridade-empty.ok{border-color:rgba(64,208,144,.24);background:rgba(64,208,144,.06);color:#78e6b0}
        .integridade-limite{font-size:.72rem;color:rgba(143,170,210,.62);padding:0 .9rem .9rem}
        @media(max-width:720px){.admin-integridade-modal-overlay{padding:.6rem}.admin-integridade-modal{width:100%;max-height:calc(100vh - 1.2rem);border-radius:16px}.admin-integridade-modal-body{max-height:calc(100vh - 132px);padding:.85rem}.integridade-dia-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.integridade-issue-top{flex-direction:column}.integridade-modal-actions{justify-content:flex-start}}
      `;
      document.head.appendChild(style);
    }
  };

  // Instala automaticamente quando o Dashboard existir.
  function tentarInstalar() {
    if (AdminIntegridade.instalarNoDashboard()) return;
    setTimeout(tentarInstalar, 300);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", tentarInstalar);
  } else {
    tentarInstalar();
  }
})(typeof window !== "undefined" ? window : globalThis);
