// ============================================================
// admin-fechamento.js — Fechamento Oficial do Dia · Admin Homy
// v: fechamento-operacional-v9-correcao-assistida-20260626
//
// Carregar depois de admin-operacao-dia.js.
// Não substitui a Operação do Dia; apenas injeta os controles de fechamento.
// ============================================================
(function (global) {
  "use strict";

  const AdminFechamento = global.AdminFechamento = {
    _semanaId: "",
    _modalId: "adminFechamentoModal",
    _ultimoPreview: null,
    _patchAplicado: false,

    init() {
      this._ensureStyle();
      this._patchAdminOperacao();
      this._ensureUI();
    },

    _patchAdminOperacao() {
      if (this._patchAplicado || !global.AdminOperacao) return;
      this._patchAplicado = true;

      const originalLoad = global.AdminOperacao.load?.bind(global.AdminOperacao);
      if (typeof originalLoad === "function") {
        global.AdminOperacao.load = async (semanaId) => {
          this._semanaId = semanaId || this._semanaId;
          const r = await originalLoad(semanaId);
          this._ensureUI();
          this.atualizarStatus().catch(e => console.warn("[Fechamento] status ignorado:", e));
          return r;
        };
      }

      const originalCarregar = global.AdminOperacao._carregar?.bind(global.AdminOperacao);
      if (typeof originalCarregar === "function") {
        global.AdminOperacao._carregar = async (semanaId) => {
          this._semanaId = semanaId || this._semanaId;
          const r = await originalCarregar(semanaId);
          this._ensureUI();
          this.atualizarStatus().catch(e => console.warn("[Fechamento] status ignorado:", e));
          return r;
        };
      }
    },

    _semanaAtual() {
      return this._semanaId || global.AdminState?.semanaId || global.AdminCore?.semanaId || global.SP?.getCurrentWeekId?.() || "";
    },

    _diaAtual() {
      return document.getElementById("operacaoDia")?.value || global.AdminUtils?.DIA_HOJE?.() || "segunda";
    },

    _diaLabel(dia) {
      const n = this._norm(dia);
      return { segunda:"Segunda", terca:"Terça", terça:"Terça", quarta:"Quarta", quinta:"Quinta", sexta:"Sexta" }[n] || dia || "Dia";
    },

    _norm(v) {
      return String(v || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
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
      if (!reload || document.getElementById("btnPreviaFechamentoDia")) return;

      const wrap = document.createElement("div");
      wrap.className = "fechamento-actions";
      wrap.innerHTML = `
        <button class="btn-secondary" id="btnPreviaFechamentoDia" type="button">📋 Prévia fechamento</button>
        <button class="btn-success" id="btnFecharOperacaoDia" type="button">🔒 Fechar dia</button>
        <button class="btn-secondary" id="btnReabrirOperacaoDia" type="button" style="display:none">🔓 Reabrir</button>
        <button class="btn-secondary" id="btnAbrirCorrecaoAssistidaFechamento" type="button">🛠 Correção assistida</button>
      `;

      reload.insertAdjacentElement("beforebegin", wrap);
      document.getElementById("btnPreviaFechamentoDia")?.addEventListener("click", () => this.abrirPrevia());
      document.getElementById("btnFecharOperacaoDia")?.addEventListener("click", () => this.fecharDia());
      document.getElementById("btnReabrirOperacaoDia")?.addEventListener("click", () => this.reabrirDia());
      document.getElementById("btnAbrirCorrecaoAssistidaFechamento")?.addEventListener("click", () => global.AdminCorrecaoIntegridade?.abrir?.({ dia: this._diaAtual() }));

      const stats = document.querySelector("#mod-operacao .stats-grid");
      if (stats && !document.getElementById("fechamentoStatusBox")) {
        stats.insertAdjacentHTML("afterend", `
          <div id="fechamentoStatusBox" class="fechamento-status-box">
            <span class="fechamento-status-muted">Fechamento:</span>
            <span id="fechamentoStatusTexto">verificando...</span>
          </div>
        `);
      }

      const diaEl = document.getElementById("operacaoDia");
      if (diaEl && !diaEl.dataset.boundFechamento) {
        diaEl.dataset.boundFechamento = "1";
        diaEl.addEventListener("change", () => this.atualizarStatus());
      }
    },

    async atualizarStatus() {
      this._ensureUI();
      const box = document.getElementById("fechamentoStatusBox");
      const txt = document.getElementById("fechamentoStatusTexto");
      const btnFechar = document.getElementById("btnFecharOperacaoDia");
      const btnReabrir = document.getElementById("btnReabrirOperacaoDia");
      if (!txt || !global.SP?.getFechamentoDia) return;

      const semanaId = this._semanaAtual();
      const dia = this._diaAtual();
      if (!semanaId || !dia) return;

      await SP.init();
      const fechamento = await SP.getFechamentoDia(semanaId, dia).catch(() => null);
      const status = this._norm(SP.pick?.(fechamento, "Status_Fechamento", "Status") || "");

      if (fechamento?.id && status !== "reaberto" && status !== "cancelado") {
        const total = SP.pick(fechamento, "Total") ?? "—";
        const fechadoPor = SP.pick(fechamento, "Fechado_Por") || "—";
        const fechadoEm = SP.pick(fechamento, "Fechado_Em") || "";
        txt.innerHTML = `🔒 <b>${this._diaLabel(dia)} fechado</b> · Total oficial: <b>${this._esc(total)}</b> · Por: ${this._esc(fechadoPor)} ${fechadoEm ? `· ${this._esc(String(fechadoEm).slice(0,16).replace("T"," "))}` : ""}`;
        box?.classList.add("is-fechado");
        if (btnFechar) btnFechar.style.display = "none";
        if (btnReabrir) btnReabrir.style.display = "";
      } else {
        txt.innerHTML = `🔓 <b>${this._diaLabel(dia)} aberto</b> · ainda não há fechamento oficial.`;
        box?.classList.remove("is-fechado");
        if (btnFechar) btnFechar.style.display = "";
        if (btnReabrir) btnReabrir.style.display = "none";
      }
    },

    async abrirPrevia() {
      try {
        await SP.init();
        const semanaId = this._semanaAtual();
        const dia = this._diaAtual();
        const previa = await SP.gerarPreviaFechamentoDia(semanaId, dia);
        previa.validacaoAuditoria = await SP.validarPreviaFechamentoContraReferencia?.(semanaId, dia, previa).catch(() => null);
        this._ultimoPreview = previa;
        this._abrirModal(previewHtml(previewModel(previa), false));
      } catch (e) {
        console.error("[Fechamento]", e);
        this._toast("Erro ao gerar prévia: " + (e.message || e), "error");
      }
    },

    async fecharDia() {
      try {
        await SP.init();
        const semanaId = this._semanaAtual();
        const dia = this._diaAtual();
        const previa = await SP.gerarPreviaFechamentoDia(semanaId, dia);
        previa.validacaoAuditoria = await SP.validarPreviaFechamentoContraReferencia?.(semanaId, dia, previa).catch(() => null);
        this._ultimoPreview = previa;

        if (previa.validacaoAuditoria?.bloquear) {
          this._abrirModal(previewHtml(previewModel(previa), false));
          this._toast("Fechamento bloqueado: a prévia diverge da referência/fechamento. Use Correção Assistida antes de fechar.", "error");
          return;
        }

        const msg = `Fechar ${this._diaLabel(dia)} como oficial?\n\nTotal: ${previa.totais.total}\nPrincipal: ${previa.totais.principal}\nLight: ${previa.totais.light}\nCarne: ${previa.totais.carne}\nMassa: ${previa.totais.massa}\n\nApós fechar, alterações antigas devem exigir reabertura/auditoria.`;
        if (!confirm(msg)) return;

        const observacao = prompt("Observação do fechamento (opcional):", "Fechamento conferido pela Operação do Dia.") || "Fechamento conferido pela Operação do Dia.";
        await SP.salvarFechamentoDia(semanaId, dia, { previa, observacao });
        this._toast("Fechamento oficial salvo.", "success");
        await this.atualizarStatus();
        if (global.AdminOperacao?._carregar) await global.AdminOperacao._carregar(semanaId);
      } catch (e) {
        console.error("[Fechamento]", e);
        this._toast("Erro ao fechar dia: " + (e.message || e), "error");
      }
    },

    async reabrirDia() {
      try {
        await SP.init();
        const semanaId = this._semanaAtual();
        const dia = this._diaAtual();
        const motivo = prompt(`Motivo obrigatório para reabrir ${this._diaLabel(dia)}:`);
        if (!motivo || !motivo.trim()) return;
        await SP.reabrirFechamentoDia(semanaId, dia, motivo.trim());
        this._toast("Dia reaberto com auditoria registrada.", "success");
        await this.atualizarStatus();
      } catch (e) {
        console.error("[Fechamento]", e);
        this._toast("Erro ao reabrir dia: " + (e.message || e), "error");
      }
    },

    _abrirModal(html) {
      this._ensureModal();
      const overlay = document.getElementById(this._modalId);
      const body = overlay?.querySelector(".fechamento-modal-body");
      if (!overlay || !body) return;
      body.innerHTML = html;
      overlay.classList.add("open");
      document.body.classList.add("fechamento-modal-open");
    },

    fecharModal() {
      document.getElementById(this._modalId)?.classList.remove("open");
      document.body.classList.remove("fechamento-modal-open");
    },

    _ensureModal() {
      if (document.getElementById(this._modalId)) return;
      const div = document.createElement("div");
      div.id = this._modalId;
      div.className = "fechamento-modal-overlay";
      div.innerHTML = `
        <div class="fechamento-modal-card" role="dialog" aria-modal="true">
          <div class="fechamento-modal-head">
            <div>
              <div class="fechamento-modal-title">Fechamento Oficial do Dia</div>
              <div class="fechamento-modal-sub">Prévia somente leitura antes de salvar o fechamento.</div>
            </div>
            <button class="fechamento-modal-x" onclick="AdminFechamento.fecharModal()">×</button>
          </div>
          <div class="fechamento-modal-body"></div>
        </div>`;
      div.addEventListener("click", ev => { if (ev.target === div) this.fecharModal(); });
      document.body.appendChild(div);
      document.addEventListener("keydown", ev => { if (ev.key === "Escape") this.fecharModal(); });
    },

    _ensureStyle() {
      if (document.getElementById("adminFechamentoStyle")) return;
      const style = document.createElement("style");
      style.id = "adminFechamentoStyle";
      style.textContent = `
        .fechamento-actions{display:inline-flex;gap:.5rem;margin-right:.6rem;flex-wrap:wrap}
        .fechamento-status-box{margin:-.7rem 0 1rem;border:1px solid rgba(80,140,255,.18);background:rgba(30,80,150,.10);border-radius:12px;padding:.75rem .9rem;color:rgba(210,225,255,.88);font-size:.82rem}
        .fechamento-status-box.is-fechado{border-color:rgba(64,208,144,.35);background:rgba(64,208,144,.10)}
        .fechamento-status-muted{color:rgba(143,170,210,.65);margin-right:.35rem}
        .fechamento-modal-open{overflow:hidden!important}
        .fechamento-modal-overlay{position:fixed;inset:0;z-index:99990;background:rgba(0,0,0,.64);backdrop-filter:blur(5px);display:none;align-items:center;justify-content:center;padding:24px}
        .fechamento-modal-overlay.open{display:flex}
        .fechamento-modal-card{width:min(1120px,96vw);height:min(820px,92vh);background:#07142a;border:1px solid rgba(80,140,255,.35);border-radius:18px;box-shadow:0 24px 80px rgba(0,0,0,.45);display:flex;flex-direction:column;overflow:hidden;color:#e8f0ff}
        .fechamento-modal-head{padding:18px 20px;border-bottom:1px solid rgba(255,255,255,.08);display:flex;align-items:flex-start;justify-content:space-between;gap:1rem;background:rgba(8,22,46,.98)}
        .fechamento-modal-title{font-family:"Barlow Condensed",sans-serif;font-size:1.35rem;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:#fff}
        .fechamento-modal-sub{font-size:.78rem;color:rgba(143,170,210,.7);margin-top:.2rem}
        .fechamento-modal-x{width:36px;height:36px;border-radius:10px;border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.07);color:#fff;font-size:1.4rem;cursor:pointer}
        .fechamento-modal-body{padding:18px 20px;overflow:auto;flex:1}
        .fechamento-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(145px,1fr));gap:.75rem;margin-bottom:1rem}
        .fechamento-kpi{border:1px solid rgba(255,255,255,.09);background:rgba(255,255,255,.045);border-radius:14px;padding:.9rem}
        .fechamento-kpi strong{font-family:"Barlow Condensed",sans-serif;font-size:1.7rem;color:#fff;display:block;line-height:1}
        .fechamento-kpi span{display:block;margin-top:.35rem;font-size:.68rem;color:rgba(143,170,210,.7);text-transform:uppercase;letter-spacing:.09em}
        .fechamento-section{margin-top:1rem;border:1px solid rgba(255,255,255,.08);border-radius:14px;background:rgba(255,255,255,.035);overflow:hidden}
        .fechamento-section h4{margin:0;padding:.8rem 1rem;border-bottom:1px solid rgba(255,255,255,.07);font-size:.95rem;color:#fff}
        .fechamento-table{width:100%;border-collapse:collapse;font-size:.78rem}
        .fechamento-table th,.fechamento-table td{padding:.55rem .7rem;border-bottom:1px solid rgba(255,255,255,.06);text-align:left;vertical-align:top}
        .fechamento-table th{color:rgba(143,170,210,.68);font-size:.65rem;text-transform:uppercase;letter-spacing:.08em;background:rgba(255,255,255,.03)}
        .fechamento-badge{display:inline-flex;border-radius:999px;padding:.15rem .45rem;border:1px solid rgba(80,140,255,.35);background:rgba(80,140,255,.14);font-weight:700;color:#b9d6ff}
        .fechamento-alerta-bloqueio{border:1px solid rgba(255,90,110,.45);background:rgba(255,90,110,.10);color:#ffd7dd;border-radius:14px;padding:.9rem 1rem;margin-bottom:1rem;font-size:.86rem;line-height:1.45}
        .fechamento-alerta-bloqueio button{margin-top:.6rem}
        .fechamento-alerta-ok{border:1px solid rgba(64,208,144,.35);background:rgba(64,208,144,.10);color:#b8f7d9;border-radius:14px;padding:.75rem 1rem;margin-bottom:1rem;font-size:.86rem}
        @media(max-width:760px){.fechamento-actions{width:100%;margin:0}.fechamento-actions button{flex:1}.fechamento-modal-overlay{padding:8px}.fechamento-modal-card{width:100vw;height:96vh}}
      `;
      document.head.appendChild(style);
    }
  };

  function previewModel(previa) {
    const t = previa?.totais || {};
    return {
      semanaId: previa.semanaId,
      dia: previa.dia,
      data: previa.dataOperacao,
      hash: previa.hashResumo,
      totais: t,
      incluidos: previa.incluidos || [],
      excluidos: previa.excluidos || [],
      existente: previa.existente || null,
      validacao: previa.validacaoAuditoria || null
    };
  }

  function previewHtml(m) {
    const esc = AdminFechamento._esc.bind(AdminFechamento);
    const rows = (m.incluidos || []).slice(0, 120).map(i => `
      <tr><td>${esc(i.colaboradorNome)}</td><td><span class="fechamento-badge">${esc(i.opcao || "—")}</span></td><td>${esc(i.categoria || "—")}</td><td>${esc(i.origem || "—")}</td><td>${esc(i.pedidoId || "—")}</td></tr>
    `).join("");
    const excl = (m.excluidos || []).slice(0, 80).map(i => `
      <tr><td>${esc(i.colaboradorNome)}</td><td>${esc(i.status || "—")}</td><td>${esc(i.motivo || "—")}</td><td>${esc(i.pedidoId || "—")}</td></tr>
    `).join("");
    const validacaoHtml = m.validacao?.bloquear ? `
      <div class="fechamento-alerta-bloqueio">
        <b>⚠️ Fechamento bloqueado por divergência.</b><br>
        ${esc(m.validacao.mensagem || "A prévia não bate com a referência.")}<br>
        <button class="btn-danger" type="button" onclick="AdminCorrecaoIntegridade?.abrir?.({dia:'${esc(m.dia)}'})">🛠 Abrir Correção Assistida</button>
      </div>` : (m.validacao?.bate ? `<div class="fechamento-alerta-ok">✅ Prévia validada contra ${esc(m.validacao.fonte || "referência")}.</div>` : "");
    return `
      ${validacaoHtml}
      <div class="fechamento-grid">
        <div class="fechamento-kpi"><strong>${esc(m.totais.total || 0)}</strong><span>Total</span></div>
        <div class="fechamento-kpi"><strong>${esc(m.totais.principal || 0)}</strong><span>Principal</span></div>
        <div class="fechamento-kpi"><strong>${esc(m.totais.light || 0)}</strong><span>Light</span></div>
        <div class="fechamento-kpi"><strong>${esc(m.totais.carne || 0)}</strong><span>Carne</span></div>
        <div class="fechamento-kpi"><strong>${esc(m.totais.massa || 0)}</strong><span>Massa</span></div>
        <div class="fechamento-kpi"><strong>${esc(m.totais.extras || 0)}</strong><span>Extras incluídos</span></div>
        <div class="fechamento-kpi"><strong>${esc(m.totais.duplicidadesIgnoradas || 0)}</strong><span>Duplicidades ignoradas</span></div>
        <div class="fechamento-kpi"><strong>${esc(m.totais.cancelados || 0)}</strong><span>Cancelados</span></div>
      </div>
      <div style="font-size:.82rem;color:rgba(190,210,240,.8);margin-bottom:.8rem">
        <b>Semana:</b> ${esc(m.semanaId)} · <b>Dia:</b> ${esc(AdminFechamento._diaLabel(m.dia))} · <b>Data:</b> ${esc(m.data)} · <b>Hash:</b> ${esc(m.hash || "—")}
      </div>
      <div class="fechamento-section">
        <h4>✅ Incluídos no fechamento (${esc(m.incluidos.length)})</h4>
        <table class="fechamento-table"><thead><tr><th>Nome</th><th>Opção</th><th>Categoria</th><th>Origem</th><th>Pedido</th></tr></thead><tbody>${rows || `<tr><td colspan="5">Nenhum incluído.</td></tr>`}</tbody></table>
      </div>
      <div class="fechamento-section">
        <h4>🚫 Excluídos / ignorados (${esc(m.excluidos.length)})</h4>
        <table class="fechamento-table"><thead><tr><th>Nome</th><th>Status</th><th>Motivo</th><th>Pedido</th></tr></thead><tbody>${excl || `<tr><td colspan="4">Nenhum excluído.</td></tr>`}</tbody></table>
      </div>
    `;
  }

  function boot() {
    AdminFechamento.init();
    let tentativas = 0;
    const t = setInterval(() => {
      AdminFechamento.init();
      if (++tentativas > 20 || global.AdminOperacao) clearInterval(t);
    }, 300);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();
})(window);
