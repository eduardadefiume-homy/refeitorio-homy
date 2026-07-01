// ============================================================
// admin-correcao-integridade.js — Correção Assistida · Admin Homy
// v: base-centralizada-correcao-v10-21-20260701
//
// Carregar depois de admin-fechamento.js e admin-operacao-dia.js.
// Não executa correção automática. Só aplica após confirmação explícita.
// ============================================================
(function (global) {
  "use strict";

  const AdminCorrecaoIntegridade = global.AdminCorrecaoIntegridade = {
    _modalId: "adminCorrecaoIntegridadeModal",
    _semanaId: "",
    _ultimoPlano: null,
    _patchAplicado: false,

    init() {
      this._ensureStyle();
      this._patchOperacao();
      this._ensureUI();
    },

    _patchOperacao() {
      if (this._patchAplicado || !global.AdminOperacao) return;
      this._patchAplicado = true;

      const originalLoad = global.AdminOperacao.load?.bind(global.AdminOperacao);
      if (typeof originalLoad === "function") {
        global.AdminOperacao.load = async (semanaId) => {
          this._semanaId = semanaId || this._semanaId;
          const r = await originalLoad(semanaId);
          this._ensureUI();
          return r;
        };
      }

      const originalCarregar = global.AdminOperacao._carregar?.bind(global.AdminOperacao);
      if (typeof originalCarregar === "function") {
        global.AdminOperacao._carregar = async (semanaId) => {
          this._semanaId = semanaId || this._semanaId;
          const r = await originalCarregar(semanaId);
          this._ensureUI();
          return r;
        };
      }
    },

    _semanaAtual(preferida = "") {
      return preferida ||
        this._ultimoPlano?.semanaId ||
        global.AdminIntegridade?._ultimoResultado?.semanaId ||
        this._semanaId ||
        global.AdminFechamento?._semanaAtual?.() ||
        global.AdminState?.semanaId ||
        global.AdminCore?.semanaId ||
        global.SP?.getCurrentWeekId?.() || "";
    },

    _diaAtual() {
      return global.AdminFechamento?._diaAtual?.() || document.getElementById("operacaoDia")?.value || global.AdminUtils?.DIA_HOJE?.() || "segunda";
    },

    _norm(v) {
      return String(v || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    },

    _diaLabel(dia) {
      const n = this._norm(dia);
      return { segunda:"Segunda", terca:"Terça", terça:"Terça", quarta:"Quarta", quinta:"Quinta", sexta:"Sexta" }[n] || dia || "Dia";
    },

    _esc(v) {
      const s = String(v ?? "");
      return s.replace(/[&<>'"]/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[ch]));
    },

    _toast(msg, tipo = "info") {
      if (global.AdminUtils?.toast) global.AdminUtils.toast(msg, tipo);
      else console.log(`[${tipo}] ${msg}`);
    },

    _ensureUI() {
      const reload = document.getElementById("btnRecarregarOperacao");
      if (!reload) return;

      // v10.3 — Correção Assistida tem um único dono visual: este módulo.
      // Remove botões criados por versões antigas do admin-fechamento.js e evita duplicidade por cache.
      const botoes = [...document.querySelectorAll("button")].filter(b =>
        String(b.textContent || "")
          .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
          .toLowerCase()
          .includes("correcao assistida")
      );

      let btn = document.getElementById("btnCorrecaoIntegridadeAssistida") || null;
      for (const b of botoes) {
        if (!btn && b.id !== "btnAbrirCorrecaoAssistidaFechamento") {
          btn = b;
          btn.id = "btnCorrecaoIntegridadeAssistida";
          continue;
        }
        if (b !== btn) b.remove();
      }

      if (!btn) {
        btn = document.createElement("button");
        btn.type = "button";
        btn.id = "btnCorrecaoIntegridadeAssistida";
        btn.className = "btn-secondary";
        btn.textContent = "🛠 Correção assistida";
        const fechamentoActions = document.querySelector(".fechamento-actions");
        if (fechamentoActions) fechamentoActions.appendChild(btn);
        else reload.insertAdjacentElement("beforebegin", btn);
      }

      if (!btn.dataset.correcaoIntegridadeBound) {
        btn.dataset.correcaoIntegridadeBound = "1";
        btn.addEventListener("click", () => this.abrir());
      }
    },

    async abrir(options = {}) {
      try {
        await SP.init();
        const semanaId = options.semanaId || this._semanaAtual();
        const dia = options.dia || this._diaAtual();
        let plano = await SP.gerarPlanoCorrecaoAssistida(semanaId, { dia: options.todos ? null : dia, force: true });
        plano = this._melhorPlano(plano, semanaId, options.todos ? "" : dia);
        this._ultimoPlano = plano;
        this._abrirModal(this._renderPlano(plano));
      } catch (e) {
        console.error("[Correção Assistida]", e);
        this._toast("Erro ao gerar plano de correção: " + (e.message || e), "error");
      }
    },

    async abrirSemana(semanaIdForcado = "") {
      try {
        await SP.init();
        const semanaId = this._semanaAtual(semanaIdForcado);
        if (!semanaId) throw new Error("Semana não identificada para correção assistida.");
        this._semanaId = semanaId;
        let plano = await SP.gerarPlanoCorrecaoAssistida(semanaId, { dia: null, force: true, incluirSemReferencia: true });
        plano = this._melhorPlano(plano, semanaId, "");
        this._ultimoPlano = plano;
        this._abrirModal(this._renderPlano(plano));
      } catch (e) {
        console.error("[Correção Assistida]", e);
        this._toast("Erro ao gerar plano da semana: " + (e.message || e), "error");
      }
    },

    async aplicarDia(dia) {
      const plano = this._ultimoPlano;
      const diaPlano = (plano?.dias || []).find(d => this._norm(d.dia) === this._norm(dia));
      if (!diaPlano) return this._toast("Plano do dia não encontrado.", "error");
      if (!diaPlano.acoesSeguras?.length) return this._toast("Não há ações seguras para aplicar neste dia.", "warning");

      const parcial = !diaPlano.fechaExato;
      const msg = parcial
        ? `Aplicar SOMENTE as ações seguras de ${this._diaLabel(dia)}?\n\nAtenção: ainda ficará pendência para revisão antes do fechamento oficial.\n\nAções: ${diaPlano.acoesSeguras.length}`
        : `Aplicar correção de ${this._diaLabel(dia)}?\n\nA simulação fecha exatamente com a referência.\n\nAções: ${diaPlano.acoesSeguras.length}`;
      if (!confirm(msg)) return;

      try {
        await SP.init();
        const r = await SP.aplicarPlanoCorrecaoAssistida(diaPlano, { aplicarParcial: true });
        this._toast(`Correção aplicada: ${r.total || 0} ação(ões).`, "success");
        await this._recarregarTelas();
        await this.abrirSemana(plano?.semanaId || "");
      } catch (e) {
        console.error("[Correção Assistida]", e);
        this._toast("Erro ao aplicar correção: " + (e.message || e), "error");
      }
    },

    async aplicarSemanaSegura() {
      const plano = this._ultimoPlano;
      const total = Number(plano?.totais?.acoesSeguras || 0);
      if (!total) return this._toast("Não há ações seguras para aplicar.", "warning");
      if (!confirm(`Aplicar todas as ações seguras da semana?

Somente cancelamentos seguros serão executados. Reativações e revisões não serão aplicadas automaticamente.\n\nTotal de ações: ${total}\n\nAs pendências de revisão não serão aplicadas.`)) return;

      try {
        await SP.init();
        const r = await SP.aplicarPlanoCorrecaoAssistida(plano, { aplicarParcial: true });
        this._toast(`Correções aplicadas: ${r.total || 0} ação(ões).`, "success");
        await this._recarregarTelas();
        await this.abrirSemana();
      } catch (e) {
        console.error("[Correção Assistida]", e);
        this._toast("Erro ao aplicar correções: " + (e.message || e), "error");
      }
    },

    baixarPlano() {
      const plano = this._ultimoPlano;
      if (!plano) return;
      const blob = new Blob([JSON.stringify(plano, null, 2)], { type: "application/json;charset=utf-8" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `plano-correcao-${plano.semanaId || "semana"}.json`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    },

    async _recarregarTelas() {
      const semanaId = this._semanaAtual();
      try { await global.AdminFechamento?.atualizarStatus?.(); } catch (_) {}
      try { if (global.AdminOperacao?._carregar) await global.AdminOperacao._carregar(semanaId); } catch (_) {}
    },

    _abrirModal(html) {
      this._ensureModal();
      const overlay = document.getElementById(this._modalId);
      const body = overlay?.querySelector(".correcao-modal-body");
      if (!overlay || !body) return;
      body.innerHTML = html;
      overlay.classList.add("open");
      document.body.classList.add("correcao-modal-open");
    },

    fecharModal() {
      document.getElementById(this._modalId)?.classList.remove("open");
      document.body.classList.remove("correcao-modal-open");
    },

    _ensureModal() {
      if (document.getElementById(this._modalId)) return;
      const div = document.createElement("div");
      div.id = this._modalId;
      div.className = "correcao-modal-overlay";
      div.innerHTML = `
        <div class="correcao-modal-card" role="dialog" aria-modal="true">
          <div class="correcao-modal-head">
            <div>
              <div class="correcao-modal-title">Correção Assistida</div>
              <div class="correcao-modal-sub">Limpeza assistida com regra central, simulação e auditoria. Nada é aplicado sem confirmação.</div>
            </div>
            <button class="correcao-modal-x" onclick="AdminCorrecaoIntegridade.fecharModal()">×</button>
          </div>
          <div class="correcao-modal-body"></div>
        </div>`;
      div.addEventListener("click", ev => { if (ev.target === div) this.fecharModal(); });
      document.body.appendChild(div);
      document.addEventListener("keydown", ev => { if (ev.key === "Escape") this.fecharModal(); });
    },

    _renderResumoLinha(label, r) {
      const esc = this._esc.bind(this);
      return `${esc(label)}: T:${esc(r?.total || 0)} · P:${esc(r?.principal || 0)} · L:${esc(r?.light || 0)} · C:${esc(r?.carne || 0)} · M:${esc(r?.massa || 0)}`;
    },

    _renderDelta(delta) {
      const partes = [];
      for (const k of ["total", "principal", "light", "carne", "massa", "lanche"]) {
        const v = Number(delta?.[k] || 0);
        if (v) partes.push(`<span class="correcao-pill ${v > 0 ? "warn" : "info"}">${this._esc(k)} ${v > 0 ? "+" : ""}${v}</span>`);
      }
      return partes.join(" ") || `<span class="correcao-pill ok">sem diferença</span>`;
    },

    _renderAcoes(acoes) {
      const esc = this._esc.bind(this);
      if (!acoes?.length) return `<div class="correcao-empty">Nenhuma ação segura.</div>`;
      return `<table class="correcao-table"><thead><tr><th>Ação</th><th>Pedido</th><th>Nome</th><th>Motivo</th><th>Resultado</th></tr></thead><tbody>${acoes.map(a => `
        <tr>
          <td><b>${esc(a.acao)}</b></td>
          <td>${esc(a.pedidoId || "—")}</td>
          <td>${esc(a.nome || "—")}</td>
          <td>${esc(a.justificativa || a.motivo || "—")}</td>
          <td>${esc(JSON.stringify(a.delta || {}))}</td>
        </tr>`).join("")}</tbody></table>`;
    },

    _renderRevisoes(revisoes) {
      const esc = this._esc.bind(this);
      if (!revisoes?.length) return "";
      return `<div class="correcao-review"><b>⚠️ Pendências para revisão</b>${revisoes.map(r => `
        <div class="correcao-review-item">${esc(r.justificativa || "Revisar")}</div>
      `).join("")}</div>`;
    },


    _normalizarAcaoDiagnostico(acao = {}) {
      const opcao = this._norm(acao.opcao || "principal");
      const op = ["principal", "light", "carne", "massa", "lanche"].includes(opcao) ? opcao : "principal";
      const delta = { total: -1 };
      delta[op] = -1;

      return {
        acao: "cancelar",
        autoAplicavel: true,
        motivo: acao.motivo || "extra-duplicado",
        pedidoId: String(acao.pedidoId || ""),
        nome: acao.nome || acao.pedido?.nome || "Pedido especial",
        dia: acao.dia || acao.pedido?.dia || "",
        opcao: op,
        statusAtual: acao.statusAtual || acao.pedido?.status || "",
        confirmadoAtual: acao.confirmadoAtual === true || acao.pedido?.confirmado === true,
        origemAtual: acao.origemAtual || acao.pedido?.origem || "",
        categoria: acao.pedido?.origem || acao.origemAtual || "especial",
        delta,
        justificativa: acao.justificativa || "Duplicidade produtiva especial.",
        camposSugeridos: acao.camposSugeridos || {
          Status: "Cancelado",
          Confirmado: false,
          Origem: "Duplicidade inativada",
          Observacao: "Correção assistida: duplicidade produtiva especial."
        },
        pedido: acao.pedido || acao
      };
    },

    _resumoAposAcoes(resumo = {}, acoes = []) {
      const r = {
        total: Number(resumo.total || 0),
        principal: Number(resumo.principal || 0),
        light: Number(resumo.light || 0),
        carne: Number(resumo.carne || 0),
        massa: Number(resumo.massa || 0),
        lanche: Number(resumo.lanche || 0),
        outros: Number(resumo.outros || 0),
        semOpcao: Number(resumo.semOpcao || 0)
      };
      for (const a of acoes || []) {
        const op = this._norm(a.opcao || "principal");
        r.total += Number(a.delta?.total ?? -1);
        if (Object.prototype.hasOwnProperty.call(r, op)) r[op] += Number(a.delta?.[op] ?? -1);
      }
      return r;
    },

    _deltaResumo(atual = {}, alvo = {}) {
      const out = {};
      for (const k of ["total", "principal", "light", "carne", "massa", "lanche"]) {
        out[k] = Number(atual?.[k] || 0) - Number(alvo?.[k] || 0);
      }
      return out;
    },

    _converterPlanoDiagnosticoParaAssistido(semanaId, diaFiltro = "") {
      const resultado = global.AdminIntegridade?._ultimoResultado;
      const planoDiag = resultado?.planoCorrecao;
      if (!resultado || !planoDiag?.acoes?.length) return null;
      if (String(resultado.semanaId || "") !== String(semanaId || "")) return null;

      const filtro = this._norm(diaFiltro || "");
      const acoes = (planoDiag.acoes || [])
        .filter(a => a?.acao === "cancelar" && a?.motivo === "extra-duplicado" && a?.pedidoId)
        .filter(a => !filtro || this._norm(a.dia) === filtro)
        .map(a => this._normalizarAcaoDiagnostico(a));

      if (!acoes.length) return null;

      const porDia = new Map();
      for (const acao of acoes) {
        const dia = this._norm(acao.dia || "segunda");
        if (!porDia.has(dia)) porDia.set(dia, []);
        porDia.get(dia).push(acao);
      }

      const dias = [];
      for (const [dia, acoesDia] of porDia.entries()) {
        const atual = resultado.resumoProducaoPorDia?.[dia] || {};
        const simulado = this._resumoAposAcoes(atual, acoesDia);
        const alvo = { ...simulado };
        dias.push({
          semanaId,
          dia,
          dataOperacao: (resultado.diasVerificados || []).find(d => this._norm(d.dia) === dia)?.data || "",
          referencia: {
            tipo: "limpeza-segura-sem-referencia",
            fonte: "Diagnóstico de Integridade",
            semReferencia: true,
            valores: alvo
          },
          atual,
          alvo,
          deltaInicial: this._deltaResumo(atual, alvo),
          candidatas: acoesDia,
          acoesSeguras: acoesDia,
          revisoes: [],
          simulado,
          deltaFinal: this._deltaResumo(simulado, alvo),
          jaBate: false,
          fechaExato: true,
          semReferencia: true,
          status: "corrigivel-sem-referencia",
          mensagem: "Duplicidade objetiva encontrada pelo Diagnóstico de Integridade. A limpeza é segura e limitada aos pedidos especiais produtivos."
        });
      }

      return {
        semanaId,
        geradoEm: new Date().toISOString(),
        status: "previa-correcao-assistida",
        escopo: filtro ? "dia" : "semana",
        origem: "diagnostico-integridade",
        dias,
        totais: {
          dias: dias.length,
          acoesSeguras: acoes.length,
          revisoes: 0,
          corrigiveis: dias.length,
          parciais: 0,
          semReferencia: dias.length
        },
        hashPlano: `diag-${String(planoDiag?.totais?.total || acoes.length)}-${String(resultado.geradoEm || "").replace(/\D/g, "").slice(-8)}`
      };
    },

    _melhorPlano(planoSP, semanaId, diaFiltro = "") {
      const totalSP = Number(planoSP?.totais?.acoesSeguras || 0);
      if (totalSP > 0) return planoSP;
      const diag = this._converterPlanoDiagnosticoParaAssistido(semanaId, diaFiltro);
      return diag || planoSP;
    },


    _renderPlano(plano) {
      const esc = this._esc.bind(this);
      const dias = plano?.dias || [];
      const header = `
        <div class="correcao-toolbar">
          <button class="btn-secondary" onclick="AdminCorrecaoIntegridade.abrirSemana('${esc(plano?.semanaId || "")}')">Ver semana inteira</button>
          <button class="btn-secondary" onclick="AdminCorrecaoIntegridade.baixarPlano()">Baixar plano JSON</button>
          <button class="btn-danger" onclick="AdminCorrecaoIntegridade.aplicarSemanaSegura()" ${Number(plano?.totais?.acoesSeguras || 0) ? "" : "disabled"}>Aplicar ações seguras da semana</button>
        </div>
        <div class="correcao-summary">
          <b>Semana:</b> ${esc(plano?.semanaId || "—")} ·
          <b>Ações seguras:</b> ${esc(plano?.totais?.acoesSeguras || 0)} ·
          <b>Pendências:</b> ${esc(plano?.totais?.revisoes || 0)} ·
          <b>Hash:</b> ${esc(plano?.hashPlano || "—")} ${plano?.origem ? "· <b>Origem:</b> " + esc(plano.origem) : ""}
        </div>`;

      if (!dias.length) {
        return header + `<div class="correcao-empty big">Não há fechamento oficial nem referência histórica para comparar. Para semanas futuras, a auditoria passa a usar o fechamento oficial depois que o dia for fechado.</div>`;
      }

      return header + dias.map(d => {
        const statusClass = d.jaBate ? "ok" : (d.fechaExato ? "good" : (d.acoesSeguras?.length ? "warn" : "bad"));
        return `
          <div class="correcao-card ${statusClass}">
            <div class="correcao-card-head">
              <div>
                <h3>${esc(this._diaLabel(d.dia))}</h3>
                <p>${esc(d.mensagem || "")}</p>
              </div>
              <button class="btn-danger" onclick="AdminCorrecaoIntegridade.aplicarDia('${esc(d.dia)}')" ${d.acoesSeguras?.length ? "" : "disabled"}>Aplicar ações seguras</button>
            </div>
            <div class="correcao-lines">
              <div>${this._renderResumoLinha("Atual", d.atual)}</div>
              <div>${this._renderResumoLinha(d.referencia?.fonte || "Referência", d.alvo)}</div>
              <div>${this._renderResumoLinha("Simulado", d.simulado)}</div>
              <div><b>Diferença inicial:</b> ${this._renderDelta(d.deltaInicial)}</div>
              <div><b>Diferença após simulação:</b> ${this._renderDelta(d.deltaFinal)}</div>
            </div>
            <div class="correcao-section"><h4>✅ Ações seguras (${esc(d.acoesSeguras?.length || 0)})</h4>${this._renderAcoes(d.acoesSeguras || [])}</div>
            ${this._renderRevisoes(d.revisoes || [])}
          </div>`;
      }).join("");
    },

    _ensureStyle() {
      if (document.getElementById("adminCorrecaoIntegridadeStyle")) return;
      const style = document.createElement("style");
      style.id = "adminCorrecaoIntegridadeStyle";
      style.textContent = `
        .correcao-modal-open{overflow:hidden!important}
        .correcao-modal-overlay{position:fixed;inset:0;z-index:100000;background:rgba(0,0,0,.68);backdrop-filter:blur(5px);display:none;align-items:center;justify-content:center;padding:22px}
        .correcao-modal-overlay.open{display:flex}
        .correcao-modal-card{width:min(1180px,97vw);height:min(860px,94vh);background:#07142a;border:1px solid rgba(255,190,80,.38);border-radius:18px;box-shadow:0 24px 80px rgba(0,0,0,.48);display:flex;flex-direction:column;overflow:hidden;color:#e8f0ff}
        .correcao-modal-head{padding:18px 20px;border-bottom:1px solid rgba(255,255,255,.08);display:flex;align-items:flex-start;justify-content:space-between;gap:1rem;background:rgba(8,22,46,.98)}
        .correcao-modal-title{font-family:"Barlow Condensed",sans-serif;font-size:1.35rem;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:#fff}
        .correcao-modal-sub{font-size:.78rem;color:rgba(143,170,210,.7);margin-top:.2rem}
        .correcao-modal-x{width:36px;height:36px;border-radius:10px;border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.07);color:#fff;font-size:1.4rem;cursor:pointer}
        .correcao-modal-body{padding:18px 20px;overflow:auto;flex:1}
        .correcao-toolbar{display:flex;gap:.6rem;flex-wrap:wrap;margin-bottom:.8rem;align-items:center}
        .correcao-summary{border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.04);padding:.75rem .9rem;border-radius:12px;font-size:.82rem;color:rgba(210,225,255,.88);margin-bottom:1rem}
        .correcao-card{border:1px solid rgba(255,255,255,.1);border-radius:14px;background:rgba(255,255,255,.035);margin:1rem 0;overflow:hidden}
        .correcao-card.ok{border-color:rgba(64,208,144,.25)}.correcao-card.good{border-color:rgba(64,208,144,.45)}.correcao-card.warn{border-color:rgba(255,193,84,.45)}.correcao-card.bad{border-color:rgba(255,80,100,.45)}
        .correcao-card-head{display:flex;gap:1rem;align-items:flex-start;justify-content:space-between;padding:1rem;border-bottom:1px solid rgba(255,255,255,.07)}
        .correcao-card h3{margin:0;color:#fff;font-size:1.05rem}.correcao-card p{margin:.25rem 0 0;color:rgba(190,210,240,.74);font-size:.8rem}
        .correcao-lines{padding:1rem;display:grid;gap:.4rem;font-size:.82rem;color:rgba(218,230,255,.9)}
        .correcao-pill{display:inline-flex;border-radius:999px;padding:.12rem .45rem;font-size:.72rem;font-weight:800;border:1px solid rgba(255,255,255,.12);margin:.1rem}.correcao-pill.warn{background:rgba(255,193,84,.13);color:#ffd36e;border-color:rgba(255,193,84,.35)}.correcao-pill.info{background:rgba(80,140,255,.13);color:#b9d6ff;border-color:rgba(80,140,255,.35)}.correcao-pill.ok{background:rgba(64,208,144,.13);color:#86efc1;border-color:rgba(64,208,144,.35)}
        .correcao-section{border-top:1px solid rgba(255,255,255,.06);padding:1rem}.correcao-section h4{margin:.1rem 0 .7rem;color:#fff}
        .correcao-table{width:100%;border-collapse:collapse;font-size:.78rem}.correcao-table th,.correcao-table td{padding:.55rem .65rem;border-bottom:1px solid rgba(255,255,255,.06);text-align:left;vertical-align:top}.correcao-table th{font-size:.65rem;text-transform:uppercase;letter-spacing:.08em;color:rgba(143,170,210,.72);background:rgba(255,255,255,.03)}
        .correcao-review{margin:0 1rem 1rem;border:1px solid rgba(255,193,84,.28);background:rgba(255,193,84,.08);border-radius:12px;padding:.8rem;color:#ffe1a0}.correcao-review-item{font-size:.8rem;margin:.35rem 0;color:#ffefc4}
        .correcao-empty{padding:.8rem;color:rgba(190,210,240,.75);font-size:.85rem}.correcao-empty.big{border:1px solid rgba(255,255,255,.08);border-radius:12px;background:rgba(255,255,255,.04)}
      `;
      document.head.appendChild(style);
    }
  };

  function boot() {
    AdminCorrecaoIntegridade.init();
    let tentativas = 0;
    const t = setInterval(() => {
      AdminCorrecaoIntegridade.init();
      if (++tentativas > 25 || global.AdminOperacao) clearInterval(t);
    }, 300);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();
})(window);
