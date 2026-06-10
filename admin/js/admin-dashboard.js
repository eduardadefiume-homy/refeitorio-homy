// admin-dashboard.js — Dashboard do Admin Homy
// Versão corrigida: fix setoresHoje object/array + render defensivo
// Data: 2026-06-10

const AdminDashboard = window.AdminDashboard = {

  _prazoTimer: null,

  async load(semanaId) {
    try {
      await SP.init();

      const resumo = await SP.getDashboardResumo(semanaId);
      const r = resumo || {};

      this._renderCards(r);

      await this._safeRender("toggle-status",     () => this._renderToggleStatus(r));
      await this._safeRender("extras-hoje",       () => this._renderExtrasHoje(r, semanaId));
      await this._safeRender("semana-table",      () => this._renderSemanaTable(r));
      await this._safeRender("ausencias",         () => this._renderAusencias(r, semanaId));
      await this._safeRender("operacao-dia",      () => this._renderOperacaoDia(r, semanaId));
      await this._safeRender("gerencial",         () => this._renderGerencial(r));
      await this._safeRender("setores",           () => this._renderSetores(r));
      await this._safeRender("proximos-dias",     () => this._renderProximosDias(r));
      await this._safeRender("alertas",           () => this._renderAlertas(r));
      await this._safeRender("prazo",             () => this._carregarPrazo(semanaId, r));

      AdminUtils.setTxt("semanaLabel", AdminState.getSemanaLabel());
    } catch (e) {
      console.error("[Dashboard]", e);
      AdminUtils.toast("Erro ao carregar dashboard: " + (e.message || e), "error");
    }
  },

  async _safeRender(nome, fn) {
    try {
      return await fn();
    } catch (e) {
      console.error(`[Dashboard/${nome}]`, e);
      AdminUtils.toast(`Erro no bloco ${nome}: ` + (e.message || e), "error");
      return null;
    }
  },

  _num(v, fallback = 0) {
    if (v === null || v === undefined || v === "") return fallback;
    if (typeof v === "number") return Number.isFinite(v) ? v : fallback;

    let s = String(v).trim();
    if (!s) return fallback;

    if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
    s = s.replace(/[^0-9.-]/g, "");

    const n = Number(s);
    return Number.isFinite(n) ? n : fallback;
  },

  _fmtQtd(v) {
    const n = this._num(v, 0);
    return Number.isInteger(n) ? String(n) : n.toLocaleString("pt-BR", { maximumFractionDigits: 2 });
  },

  _diaHoje(r) {
    return r?.diaHoje || (typeof AdminUtils.DIA_HOJE === "function" ? AdminUtils.DIA_HOJE() : "segunda");
  },

  _inputDateLocal(date) {
    if (!(date instanceof Date) || isNaN(date)) return "";
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  },

  _inputTimeLocal(date) {
    if (!(date instanceof Date) || isNaN(date)) return "";
    const h = String(date.getHours()).padStart(2, "0");
    const m = String(date.getMinutes()).padStart(2, "0");
    return `${h}:${m}`;
  },

  _renderCards(r) {
    const setoresNormalizados = this._normalizarSetores(r?.setoresHoje);

    AdminUtils.setTxt("stat-colab",        this._num(r?.colaboradoresAtivos));
    AdminUtils.setTxt("stat-confirmados",  this._num(r?.pedidosConfirmadosColaboradores));
    AdminUtils.setTxt("stat-pendentes",    this._num(r?.pendentesColaboradores));
    AdminUtils.setTxt("stat-checkin",      this._num(r?.checkinsHoje));
    AdminUtils.setTxt("dashExtrasAtivos",  this._num(r?.extrasAtivos));
    AdminUtils.setTxt("dashTotalHoje",     this._num(r?.totalPedidosHoje));
    AdminUtils.setTxt("dashAusenciasHoje", this._num(r?.ausenciasHoje));
    AdminUtils.setTxt("dashTotalSemana",   this._num(r?.totalPedidosSemana));
    AdminUtils.setTxt("dashPrincipalHoje", this._num(r?.principalHoje));
    AdminUtils.setTxt("dashLightHoje",     this._num(r?.lightHoje));
    AdminUtils.setTxt("dashOutrasHoje",    this._num(r?.outrasHoje));
    AdminUtils.setTxt("dashSetoresHoje",   setoresNormalizados.length);
  },

  async _renderToggleStatus() {
    const liberado  = await SP.isCardapioLiberado().catch(() => false);
    const cardapio  = await SP.getConfig("cardapio_visivel").catch(() => null);
    const cardapioV = SP.isTrue(cardapio);

    const tMarcacao = document.getElementById("toggleMarcacao");
    const tCardapio = document.getElementById("toggleCardapio");

    // Atualiza o estado visual sempre que o dashboard recarregar.
    if (tMarcacao) tMarcacao.checked = !!liberado;
    if (tCardapio) tCardapio.checked = !!cardapioV;

    if (tMarcacao && !tMarcacao.dataset.bound) {
      tMarcacao.dataset.bound = "1";
      tMarcacao.addEventListener("change", async function () {
        try {
          await SP.setMarcacaoLiberada(this.checked);
          AdminUtils.toast(this.checked ? "Marcação liberada." : "Marcação bloqueada.", "success");
          AdminUtils.setTxt("dashMarcacaoStatus", this.checked ? "Aberta" : "Fechada");
        } catch (e) {
          AdminUtils.toast("Erro ao alterar marcação: " + (e.message || e), "error");
          this.checked = !this.checked;
        }
      });
    }

    if (tCardapio && !tCardapio.dataset.bound) {
      tCardapio.dataset.bound = "1";
      tCardapio.addEventListener("change", async function () {
        try {
          await SP.setCardapioVisivel(this.checked);
          AdminUtils.toast(this.checked ? "Cardápio visível." : "Cardápio ocultado.", "success");
          AdminUtils.setTxt("dashCardapioStatus", this.checked ? "Liberado" : "Bloqueado");
        } catch (e) {
          AdminUtils.toast("Erro ao alterar cardápio: " + (e.message || e), "error");
          this.checked = !this.checked;
        }
      });
    }

    AdminUtils.setTxt("dashMarcacaoStatus", liberado  ? "Aberta"   : "Fechada");
    AdminUtils.setTxt("dashCardapioStatus", cardapioV ? "Liberado" : "Bloqueado");
  },

  // ── Prazo + Painel de Travamento ──────────────────────────────
  async _carregarPrazo(semanaId, resumo) {
    const prazo = await SP.getPrazoMarcacao().catch(() => null);

    if (prazo) {
      const dt = new Date(prazo);
      AdminUtils.setVal("prazoData", this._inputDateLocal(dt));
      AdminUtils.setVal("prazoHora", this._inputTimeLocal(dt));
    }

    this._atualizarPainelPrazo(prazo, this._num(resumo?.pendentesColaboradores));

    const btnSalvar = document.getElementById("btnSalvarPrazo");
    if (btnSalvar && !btnSalvar.dataset.bound) {
      btnSalvar.dataset.bound = "1";
      btnSalvar.addEventListener("click", async () => {
        const data = AdminUtils.getVal("prazoData");
        const hora = AdminUtils.getVal("prazoHora") || "18:00";

        if (!data) {
          AdminUtils.toast("Informe a data limite.", "error");
          return;
        }

        const valor = `${data}T${hora}:00`;

        try {
          await SP.setPrazoMarcacao(valor);
          AdminUtils.toast("✅ Prazo salvo.", "success");
          const r = await SP.getDashboardResumo(AdminState.getSemanaId()).catch(() => ({}));
          this._atualizarPainelPrazo(valor, this._num(r?.pendentesColaboradores));
        } catch (e) {
          AdminUtils.toast("Erro ao salvar prazo: " + (e.message || e), "error");
        }
      });
    }
  },

  _atualizarPainelPrazo(prazoISO, pendentes) {
    const el = document.getElementById("dashPainelPrazo");
    if (!el) return;

    const agora   = new Date();
    const dt      = prazoISO ? new Date(prazoISO) : null;
    const vencido = dt && !isNaN(dt) && agora > dt;
    const qtdPend = this._num(pendentes);

    const fmtDt = dt && !isNaN(dt)
      ? dt.toLocaleDateString("pt-BR", { weekday: "long", day: "2-digit", month: "2-digit" }) +
        " às " + dt.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })
      : null;

    let alertClass = "alert-info";
    let icone      = "⏱️";
    let statusTxt  = "";

    if (!dt || isNaN(dt)) {
      statusTxt = "Nenhum prazo definido para esta semana.";
      alertClass = "alert-warning";
      icone = "⚠️";
    } else if (vencido) {
      alertClass = "alert-red";
      icone = "🔒";
      statusTxt = qtdPend > 0
        ? `Prazo encerrado em ${fmtDt}. ${qtdPend} colaborador(es) não marcaram.`
        : `Prazo encerrado em ${fmtDt}. Todos marcaram.`;
    } else {
      const msRestante = dt - agora;
      const horas = Math.floor(msRestante / 3600000);
      const minutos = Math.floor((msRestante % 3600000) / 60000);
      const tempoTxt = horas > 0 ? `${horas}h ${minutos}min restantes` : `${minutos} min restantes`;
      statusTxt = `Prazo: ${fmtDt} (${tempoTxt}). ${qtdPend} colaborador(es) ainda não marcaram.`;
    }

    const btnHtml = qtdPend > 0
      ? `<button id="btnTravarPendentes" class="btn-primary" style="flex-shrink:0;font-size:.82rem;padding:.5rem 1rem;margin-top:.5rem">
           🔒 Travar ${qtdPend} pendente(s) como Principal
         </button>`
      : "";

    el.innerHTML = `
      <div class="alert ${alertClass}" style="display:flex;flex-direction:column;gap:.5rem">
        <span>${icone} ${AdminUtils.esc(statusTxt)}</span>
        ${btnHtml}
      </div>
    `;

    const btnTravar = document.getElementById("btnTravarPendentes");
    if (btnTravar && !btnTravar.dataset.bound) {
      btnTravar.dataset.bound = "1";
      btnTravar.addEventListener("click", () =>
        this._confirmarTravamento(AdminState.getSemanaId(), qtdPend)
      );
    }
  },

  async _confirmarTravamento(semanaId, qtd) {
    if (typeof SP.travarPendentesComoPrincipal !== "function") {
      AdminUtils.toast("Função SP.travarPendentesComoPrincipal não encontrada no sharepoint.js.", "error");
      return;
    }

    const resumo = await SP.getDashboardResumo(semanaId).catch(() => ({}));
    const total  = this._num(qtd ?? resumo?.pendentesColaboradores);

    if (total <= 0) {
      AdminUtils.toast("✅ Todos os colaboradores já marcaram.", "success");
      return;
    }

    const confirmar = window.confirm(
      `Confirmar travamento?\n\n` +
      `${total} colaborador(es) não marcaram no prazo.\n` +
      `Todos serão marcados como PRINCIPAL em todos os dias desta semana.\n\n` +
      `A observação "Marcado automaticamente — prazo encerrado" será registrada para auditoria.\n\n` +
      `Esta ação não pode ser desfeita.`
    );
    if (!confirmar) return;

    const btn = document.getElementById("btnTravarPendentes");
    if (btn) {
      btn.disabled = true;
      btn.textContent = "⏳ Travando...";
    }

    try {
      const resultado = await SP.travarPendentesComoPrincipal(semanaId);
      const travados = this._num(resultado?.travados, total);

      AdminUtils.toast(`✅ ${travados} colaboradores travados como Principal.`, "success");

      await SP.setMarcacaoLiberada(false).catch(e => {
        console.warn("[Dashboard] não foi possível bloquear marcação após travamento:", e);
        AdminUtils.toast("Travamento concluído, mas não consegui bloquear a marcação automaticamente.", "error");
      });

      await AdminDashboard.load(semanaId);
    } catch (e) {
      AdminUtils.toast("Erro ao travar: " + (e.message || e), "error");
      if (btn) {
        btn.disabled = false;
        btn.textContent = "🔒 Travar pendentes como Principal";
      }
    }
  },

  // ── Extras do Dia ─────────────────────────────────────────────
  async _renderExtrasHoje(r, semanaId) {
    AdminUtils.setTxt("dashExtrasSolicitados", this._num(r?.extrasAtivos));
    AdminUtils.setTxt("dashExtrasConfirmados", this._num(r?.extrasConfirmados));
    AdminUtils.setTxt("dashExtrasPendentes",   this._num(r?.extrasPendentes));

    const tbody = document.getElementById("dashExtrasTable");
    if (!tbody) return;

    const diaHoje = this._diaHoje(r);
    const extras = await SP.getExtras(semanaId, diaHoje).catch(() => []);

    if (!extras.length) {
      tbody.innerHTML = `<tr><td colspan="4" class="empty-cell">Nenhum extra para hoje.</td></tr>`;
      return;
    }

    tbody.innerHTML = extras.slice(0, 12).map(e => `
      <tr>
        <td>${AdminUtils.esc(SP.pick(e, "Nome", "Title") || "—")}</td>
        <td>${AdminUtils.esc(SP.pick(e, "tipo", "Tipo") || "Extra")}</td>
        <td>${AdminUtils.esc(SP.pick(e, "Opcao") || "—")}</td>
        <td>${AdminUtils.badge(SP.pick(e, "Status") || "Confirmado")}</td>
      </tr>`).join("");
  },

  _renderSemanaTable(r) {
    const tbody = document.getElementById("dashSemanaTable");
    if (!tbody) return;

    tbody.innerHTML = AdminUtils.DIAS.map(dia => {
      const d = r?.porDia?.[dia] || {};
      return `<tr>
        <td>${AdminUtils.DIA_LABEL[dia]}</td>
        <td>${this._num(d.total)}</td>
        <td>${this._num(d.principal)}</td>
        <td>${this._num(d.light)}</td>
        <td>${this._num(d.pendentes)}</td>
      </tr>`;
    }).join("");
  },

  // ── Ausências ────────────────────────────────────────────────
  async _renderAusencias(r, semanaId) {
    AdminUtils.setTxt("dashNaoMarcaram", this._num(r?.pendentesColaboradores));
    AdminUtils.setTxt("dashCancelados",  this._num(r?.ausenciasHoje));
    AdminUtils.setTxt("dashTravados",    this._num(r?.travadosHoje ?? r?.travadosSemana ?? 0));

    const tbody = document.getElementById("dashAusenciasTable");
    if (!tbody) return;

    const diaHoje = this._diaHoje(r);
    const pedidos = await SP.getPedidos(semanaId).catch(() => []);
    const norm    = v => AdminUtils.norm(v);

    const ausentes = pedidos.filter(p =>
      norm(SP.pick(p, "Dia")) === norm(diaHoje) &&
      ["cancelado", "afastado", "ferias", "nao vai almocar", "bloqueado", "travado"]
        .includes(norm(SP.pick(p, "Status") || ""))
    );

    if (!ausentes.length) {
      tbody.innerHTML = `<tr><td colspan="3" class="empty-cell">Nenhuma ausência hoje.</td></tr>`;
      return;
    }

    tbody.innerHTML = ausentes.slice(0, 14).map(p => `
      <tr>
        <td>${AdminUtils.esc(SP.pick(p, "Colaborador_nome", "Title", "Nome") || "—")}</td>
        <td>${AdminUtils.esc(SP.pick(p, "Centro_Custo", "Setor", "Departamento") || "—")}</td>
        <td>${AdminUtils.badge(SP.pick(p, "Status") || "Pendente")}</td>
      </tr>`).join("");
  },

  // ── Operação do Dia ──────────────────────────────────────────
  async _renderOperacaoDia(r, semanaId) {
    const diaHoje = this._diaHoje(r);
    const sel = document.getElementById("dashOperacaoDia");

    if (sel) {
      if (!sel.value) sel.value = diaHoje;

      if (!sel.dataset.bound) {
        sel.dataset.bound = "1";
        sel.addEventListener("change", () =>
          this._carregarOperacaoTabela(AdminState.getSemanaId(), sel.value)
        );
      }
    }

    const btn = document.getElementById("btnAtualizarDashboard");
    if (btn && !btn.dataset.bound) {
      btn.dataset.bound = "1";
      btn.addEventListener("click", () => AdminDashboard.load(AdminState.getSemanaId()));
    }

    await this._carregarOperacaoTabela(semanaId, sel?.value || diaHoje);
  },

  async _carregarOperacaoTabela(semanaId, dia) {
    const tbody = document.getElementById("dashOperacaoTable");
    if (!tbody) return;

    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    const pedidos = await SP.getPedidos(semanaId).catch(() => []);
    const norm = v => AdminUtils.norm(v);
    const diaNorm = norm(dia);
    const lista = pedidos.filter(p => norm(SP.pick(p, "Dia")) === diaNorm);

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhum pedido para este dia.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.slice(0, 60).map(p => {
      const id     = AdminUtils.esc(p.id || "");
      const nome   = AdminUtils.esc(SP.pick(p, "Colaborador_nome", "Title", "Nome") || "—");
      const setor  = AdminUtils.esc(this._formatarCC(SP.pick(p, "Centro_Custo", "Setor", "Departamento") || "—"));
      const opcao  = AdminUtils.esc(SP.pick(p, "Opcao") || "—");
      const prato  = AdminUtils.esc(SP.pick(p, "Nome_Prato", "Descricao") || "—");
      const status = SP.pick(p, "Status") || "Pendente";

      return `<tr>
        <td>${nome}</td>
        <td>${setor}</td>
        <td><span class="badge badge-blue">${opcao}</span></td>
        <td>${prato}</td>
        <td>${AdminUtils.badge(status)}</td>
        <td><div class="table-actions">
          <button class="btn-icon" title="Confirmar" onclick="AdminOperacao.alterarStatus('${id}','Confirmado')">✅</button>
          <button class="btn-icon danger" title="Cancelar" onclick="AdminOperacao.alterarStatus('${id}','Cancelado')">❌</button>
          <button class="btn-icon" title="Não vai almoçar" onclick="AdminOperacao.alterarStatus('${id}','Não vai almoçar')">🚫</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  // ── Resumo Gerencial ─────────────────────────────────────────
  _renderGerencial(r) {
    const el = document.getElementById("dashGerencialList");
    if (!el) return;

    el.innerHTML = [
      ["Consumo da semana",    `${this._num(r?.totalPedidosSemana)} refeições`],
      ["Pendências da semana", `${this._num(r?.pendentesColaboradores)} registros`],
      ["Extras confirmados",   `${this._num(r?.extrasConfirmados)} extras`],
      ["Ausências hoje",       `${this._num(r?.ausenciasHoje)} registros`]
    ].map(([a, b]) => `
      <div class="dashboard-list-item">
        <div>
          <div class="dashboard-list-main">${AdminUtils.esc(a)}</div>
          <div class="dashboard-list-sub">${AdminUtils.esc(b)}</div>
        </div>
      </div>`).join("");
  },

  // Mapa código → nome de centro de custo.
  _CC_MAPA: {
    "110101":"DIRETORIA PRESIDENCIAL","110201":"DIRETORIA ADMINISTRATIVA",
    "110202":"DIRETORIA DE PRODUTOS","120101":"ADM GERAL","120102":"CUSTOS",
    "120103":"LEGALIZAÇÃO","120201":"CONTABILIDADE","120202":"FISCAL",
    "120301":"FINANCEIRO","120401":"RECURSOS HUMANOS","120402":"DEPARTAMENTO PESSOAL",
    "120501":"TI","120601":"RECEPÇÃO","120602":"PORTARIA",
    "120603":"ASSEIO E CONSERVAÇÃO","120604":"JARDINAGEM","150101":"SUPRIMENTOS",
    "160101":"CONTROLADORIA E COMPLIANCE","160102":"ADM CONTRATOS","170101":"SGI",
    "180101":"P&D","190101":"PATIO EXTERNO","220101":"ADM VENDAS",
    "220201":"COML INTERNO - SUPORTE","220202":"COML INTERNO - ATIVO",
    "220301":"COML EXTERNO - CLT","220302":"COML EXTERNO - REPRESENTANTE",
    "230101":"SUPORTE TECNICO INDUSTRIAL","230102":"SUPORTE TECNICO OBRAS/INFRA",
    "240101":"MARKETING","250101":"FATURAMENTO","250102":"LOGISTICA",
    "250103":"EXPEDIÇÃO","320101":"PRODUÇÃO","320201":"ENVASE MANUAL",
    "320202":"ENVASE AUTOMATICO","320301":"LABORATORIO E CONTROLE QUALIDADE",
    "360101":"APOIO A PRODUÇÃO","360102":"PCP","360201":"MANUTENÇÃO",
    "360301":"ALMOXARIFADO DE INSUMOS"
  },

  _formatarCC(valor) {
    if (!valor) return "Sem setor";

    const v = String(valor).trim();
    if (!v || v === "—") return "Sem setor";

    // Já está formatado.
    if (v.includes(" - ")) return v;

    // É só o código numérico.
    if (/^\d+$/.test(v)) {
      const nome = this._CC_MAPA[v];
      return nome ? `${v} - ${nome}` : v;
    }

    return v;
  },

  _extrairSetorTotal(item) {
    if (Array.isArray(item)) {
      return {
        setor: item[0] ?? "Sem setor",
        total: this._num(item[1])
      };
    }

    if (item && typeof item === "object") {
      const entries = Object.entries(item);

      // Exemplo aceito: { "120501": 3 }
      if (entries.length === 1 && typeof entries[0][1] !== "object") {
        return {
          setor: entries[0][0] || "Sem setor",
          total: this._num(entries[0][1])
        };
      }

      const setor = SP.pick(
        item,
        "setor", "Setor",
        "centroCusto", "CentroCusto", "Centro_Custo",
        "departamento", "Departamento",
        "nome", "Nome", "label", "Label", "key", "Key"
      );

      const total = SP.pick(
        item,
        "total", "Total",
        "quantidade", "Quantidade",
        "qtd", "Qtd",
        "count", "Count",
        "valor", "Valor"
      );

      return {
        setor: setor || "Sem setor",
        total: this._num(total)
      };
    }

    return { setor: "Sem setor", total: 0 };
  },

  _normalizarSetores(raw) {
    let base = [];

    if (Array.isArray(raw)) {
      base = raw;
    } else if (raw instanceof Map) {
      base = Array.from(raw.entries());
    } else if (raw && typeof raw === "object") {
      base = Object.entries(raw);
    }

    const acumulado = new Map();

    base.forEach(item => {
      const { setor, total } = this._extrairSetorTotal(item);
      const setorLimpo = String(setor || "Sem setor").trim() || "Sem setor";
      const qtd = this._num(total);

      if (qtd <= 0) return;

      acumulado.set(setorLimpo, (acumulado.get(setorLimpo) || 0) + qtd);
    });

    return Array.from(acumulado.entries())
      .map(([setor, total]) => ({ setor, total }))
      .sort((a, b) => b.total - a.total || String(a.setor).localeCompare(String(b.setor)));
  },

  _renderSetores(r) {
    const el = document.getElementById("dashSetoresList");
    if (!el) return;

    const setores = this._normalizarSetores(r?.setoresHoje).slice(0, 8);

    el.innerHTML = setores.length
      ? setores.map(({ setor, total }) => `
          <div class="dashboard-list-item">
            <div>
              <div class="dashboard-list-main">${AdminUtils.esc(this._formatarCC(setor))}</div>
              <div class="dashboard-list-sub">${this._fmtQtd(total)} refeições hoje</div>
            </div>
            <span class="badge badge-blue">${this._fmtQtd(total)}</span>
          </div>`).join("")
      : `<div class="dashboard-list-item"><div class="dashboard-list-main">Sem dados hoje</div></div>`;
  },

  _renderProximosDias(r) {
    const el = document.getElementById("dashProximosDias");
    if (!el) return;

    el.innerHTML = AdminUtils.DIAS.map(dia => {
      const d = r?.porDia?.[dia] || {};
      return `<div class="dashboard-list-item">
        <div>
          <div class="dashboard-list-main">${AdminUtils.esc(AdminUtils.DIA_LABEL[dia] || dia)}</div>
          <div class="dashboard-list-sub">${this._num(d.total)} conf. · ${this._num(d.pendentes)} pend.</div>
        </div>
      </div>`;
    }).join("");
  },

  _renderAlertas(r) {
    const el = document.getElementById("dashAlertas");
    if (!el) return;

    const itens = [];
    const pendentes = this._num(r?.pendentesColaboradores);
    const extrasPendentes = this._num(r?.extrasPendentes);

    if (pendentes > 0) {
      itens.push(`⚠️ ${pendentes} colaboradores ainda não marcaram.`);
    }

    if (extrasPendentes > 0) {
      itens.push(`⚠️ ${extrasPendentes} extras aguardando confirmação.`);
    }

    if (!itens.length) {
      itens.push("✅ Operação sem alertas críticos no momento.");
    }

    el.innerHTML = itens.map(x =>
      `<div class="dashboard-alert-item ${x.startsWith("✅") ? "ok" : ""}">${AdminUtils.esc(x)}</div>`
    ).join("");

    AdminUtils.setTxt("dashAlteracoesManuais", this._num(r?.alteracoesManuais ?? 0));
  }
};
