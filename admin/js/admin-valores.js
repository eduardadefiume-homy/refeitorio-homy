// admin-valores.js — Valores de Refeição + NF Vascon + Reconciliação
//
// NÃO usa nomes de coluna hardcoded.
// Usa SP._resolveColunasValores() para descobrir dinamicamente
// os nomes internos das colunas do SharePoint, evitando que
// qualquer renomeação quebre o código.

const AdminValores = window.AdminValores = {

  _lista:      [],
  _cols:       null,   // cache do mapa de colunas
  _editandoId: null,

  // ── Resolve colunas dinamicamente (cache por sessão) ──────────
  async _getCols() {
    if (this._cols) return this._cols;
    this._cols = await SP._resolveColunasValores();
    console.log("[AdminValores] colunas:", this._cols);
    return this._cols;
  },

  // Lê um campo usando o mapa dinâmico com fallback para o nome canônico
  _ler(item, chave) {
    if (!item || !this._cols) return null;
    const nomeReal = this._cols[chave];
    // Tenta o nome real detectado, depois nomes canônicos conhecidos
    const fallbacks = {
      titulo:   ["Title", "Titulo"],
      inicio:   ["Data_Inicio", "DataInicio"],
      fim:      ["Data_Fim",    "DataFim"],
      vascon:   ["Valor_Vascon", "ValorVascon"],
      desconto: ["Valor_Desconto_Funcionário", "Valor_Desconto_Funcionario", "ValorDescontoFuncionario"],
      obs:      ["Observacao", "Obs"],
      ativo:    ["Ativo"]
    };
    const tentativas = nomeReal
      ? [nomeReal, ...(fallbacks[chave] || [])]
      : (fallbacks[chave] || []);

    for (const k of tentativas) {
      if (item[k] !== undefined && item[k] !== null) return item[k];
    }
    return null;
  },

  // ── Carregamento ──────────────────────────────────────────────
  async load() {
    this._cols = null; // reseta cache para reler colunas
    this._bindBotoes();
    this._bindNF();
    await this._carregar();
  },

  async _carregar() {
    const tbody = document.getElementById("valoresTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Carregando...</td></tr>`;
    try {
      await SP.init();
      await this._getCols();
      this._lista = await SP.getValoresRefeicao(false);
      this._render();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
    }
  },

  _render() {
    const tbody = document.getElementById("valoresTable");
    if (!tbody) return;
    if (!this._lista.length) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Nenhum valor cadastrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = this._lista.map(v => {
      const id     = AdminUtils.esc(v.id || "");
      const titulo = AdminUtils.esc(this._ler(v, "titulo") || "—");
      const inicio = this._fmtData(this._ler(v, "inicio"));
      const fim    = this._fmtData(this._ler(v, "fim"));
      const vascon = Number(this._ler(v, "vascon")   || 0).toFixed(2);
      const desc   = Number(this._ler(v, "desconto") || 0).toFixed(2);
      // Ativo: se o campo vier null/undefined, assume ativo (true)
      const ativoRaw = this._ler(v, "ativo");
      const ativo  = ativoRaw === null || ativoRaw === undefined
        ? true
        : SP.isTrue(ativoRaw);

      return `<tr>
        <td>${titulo}</td>
        <td>${inicio || "—"}</td>
        <td>${fim    || "—"}</td>
        <td>R$ ${vascon}</td>
        <td>R$ ${desc}</td>
        <td><span class="badge ${ativo ? "badge-green" : "badge-red"}">${ativo ? "ATIVO" : "INATIVO"}</span></td>
        <td><div class="table-actions">
          <button class="btn-icon" title="Editar"  onclick="AdminValores.abrirEdicao('${id}')">✏️</button>
          <button class="btn-icon danger" title="Excluir" onclick="AdminValores.excluir('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  _fmtData(v) {
    if (!v) return "";
    const d = new Date(v);
    if (isNaN(d)) return String(v).slice(0, 10);
    return d.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit", year: "numeric" });
  },

  // ── Modal valor ───────────────────────────────────────────────
  abrirNovo() {
    this._editandoId = null;
    this._limparModal();
    const t = document.querySelector("#modalValorRefeicao .modal-title");
    if (t) t.textContent = "Novo valor de refeição";
    AdminUtils.openModal("modalValorRefeicao");
  },

  abrirEdicao(id) {
    this._editandoId = id;
    const v = this._lista.find(x => String(x.id) === String(id));
    if (!v) { AdminUtils.toast("Valor não encontrado.", "error"); return; }

    AdminUtils.setVal("valorTitulo",     this._ler(v, "titulo")   || "");
    AdminUtils.setVal("valorDataInicio", (this._ler(v, "inicio")  || "").toString().slice(0, 10));
    AdminUtils.setVal("valorDataFim",    (this._ler(v, "fim")     || "").toString().slice(0, 10));
    AdminUtils.setVal("valorVascon",     this._ler(v, "vascon")   ?? "");
    AdminUtils.setVal("valorDesconto",   this._ler(v, "desconto") ?? "");
    AdminUtils.setVal("valorObs",        this._ler(v, "obs")      ?? "");

    const ativoRaw = this._ler(v, "ativo");
    const ativo = ativoRaw === null || ativoRaw === undefined ? true : SP.isTrue(ativoRaw);
    AdminUtils.setVal("valorAtivo", ativo ? "sim" : "nao");

    const t = document.querySelector("#modalValorRefeicao .modal-title");
    if (t) t.textContent = "Editar valor de refeição";
    AdminUtils.openModal("modalValorRefeicao");
  },

  _limparModal() {
    ["valorTitulo", "valorDataInicio", "valorDataFim", "valorVascon", "valorDesconto", "valorObs"]
      .forEach(id => AdminUtils.setVal(id, ""));
    AdminUtils.setVal("valorAtivo", "sim");
  },

  async salvar() {
    const titulo   = AdminUtils.getVal("valorTitulo")     || "Valor refeição";
    const inicio   = AdminUtils.getVal("valorDataInicio");
    const fim      = AdminUtils.getVal("valorDataFim");
    const vascon   = AdminUtils.moeda(AdminUtils.getVal("valorVascon"));
    const desconto = AdminUtils.moeda(AdminUtils.getVal("valorDesconto"));
    const obs      = AdminUtils.getVal("valorObs");
    const ativo    = AdminUtils.getVal("valorAtivo") !== "nao";

    if (!inicio || !fim)   { AdminUtils.toast("Informe data início e fim.", "error"); return; }
    if (vascon === null)    { AdminUtils.toast("Informe o valor Vascon.",    "error"); return; }
    if (desconto === null)  { AdminUtils.toast("Informe o desconto.",        "error"); return; }

    try {
      await SP.init();
      const cols = await this._getCols();

      // Monta fields usando os nomes reais das colunas
      const fields = { Title: titulo };
      if (cols.inicio)   fields[cols.inicio]   = inicio;
      if (cols.fim)      fields[cols.fim]       = fim;
      if (cols.vascon)   fields[cols.vascon]    = vascon;
      if (cols.desconto) fields[cols.desconto]  = desconto;
      if (cols.obs)      fields[cols.obs]       = obs;
      if (cols.ativo)    fields[cols.ativo]     = ativo;
      else               fields["Ativo"]        = ativo; // fallback

      if (this._editandoId) {
        await SP.updateItem("Valores de Refeição", this._editandoId, fields);
        AdminUtils.toast("✅ Valor atualizado.", "success");
      } else {
        await SP.createItem("Valores de Refeição", fields);
        AdminUtils.toast("✅ Valor criado.", "success");
      }
      AdminUtils.closeModal("modalValorRefeicao");
      this._editandoId = null;
      this._cols = null; // força re-resolve na próxima carga
      await this._carregar();
    } catch (e) {
      AdminUtils.toast("Erro ao salvar: " + e.message, "error");
    }
  },

  async excluir(id) {
    if (!confirm("Excluir este valor?")) return;
    try {
      await SP.init();
      await SP.deleteItem("Valores de Refeição", id);
      this._lista = this._lista.filter(v => String(v.id) !== String(id));
      this._render();
      AdminUtils.toast("Valor excluído.", "success");
    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + e.message, "error");
    }
  },

  // ── NF Vascon ─────────────────────────────────────────────────
  _bindNF() {
    const bind = (id, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundNF) { el.dataset.boundNF = "1"; el.addEventListener("click", fn); }
    };
    bind("btnUploadNF",    () => document.getElementById("inputNFPDF")?.click());
    bind("btnReconciliar", () => this._executarReconciliacao());

    const inp = document.getElementById("inputNFPDF");
    if (inp && !inp.dataset.boundNF) {
      inp.dataset.boundNF = "1";
      inp.addEventListener("change", () => this._uploadNF(inp.files[0]));
    }
  },

  async _uploadNF(file) {
    if (!file) return;
    const valorId = AdminUtils.getVal("nfValorId");
    if (!valorId) { AdminUtils.toast("Selecione o período da NF antes de fazer upload.", "error"); return; }
    try {
      await SP.init();
      const token  = await SP.getToken();
      const siteId = await SP.getSiteId();
      const url    = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/NF-Vascon/${file.name}:/content`;
      AdminUtils.toast("Fazendo upload do PDF...", "info");
      const res = await fetch(url, {
        method: "PUT",
        headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/pdf" },
        body: file
      });
      if (!res.ok) throw new Error(`Upload falhou: ${res.status}`);
      const data = await res.json();
      const cols  = await this._getCols();
      if (cols.obs) {
        await SP.updateItem("Valores de Refeição", valorId, {
          [cols.obs]: `NF: ${file.name} | ${data.webUrl || ""}`
        });
      }
      AdminUtils.toast("✅ NF enviada.", "success");
    } catch (e) {
      AdminUtils.toast("Erro no upload: " + e.message, "error");
    }
  },

  async reconciliar(id) {
    const v = this._lista.find(x => String(x.id) === String(id));
    if (!v) return;
    AdminUtils.setVal("nfValorId", id);
    AdminUtils.setVal("nfInicio",  (this._ler(v, "inicio") || "").toString().slice(0, 10));
    AdminUtils.setVal("nfFim",     (this._ler(v, "fim")    || "").toString().slice(0, 10));
    const vu = Number(this._ler(v, "vascon") || 0);
    AdminUtils.setTxt("nfVasconUnit", `R$ ${vu.toFixed(2)} / refeição`);
    AdminUtils.openModal("modalNFVascon");
    await this._executarReconciliacao();
  },

  async _executarReconciliacao() {
    const inicio  = AdminUtils.getVal("nfInicio");
    const fim     = AdminUtils.getVal("nfFim");
    const nfTotal = AdminUtils.moeda(AdminUtils.getVal("nfTotalDigitado"));
    const valorId = AdminUtils.getVal("nfValorId");
    const wrap    = document.getElementById("nfResultado");
    if (!wrap) return;
    if (!inicio || !fim) { wrap.innerHTML = `<div class="alert alert-warning">Informe o período.</div>`; return; }
    wrap.innerHTML = `<div class="alert alert-info">Calculando...</div>`;
    try {
      await SP.init();
      const v = this._lista.find(x => String(x.id) === String(valorId));
      const valorUnit = v ? Number(this._ler(v, "vascon") || 0) : 0;
      const todos = await SP.getItems("Pedidos");
      const norm  = s => AdminUtils.norm(s);
      const noIntervalo = todos.filter(p => {
        const data = (SP.pick(p, "Data_Hora") || "").slice(0, 10);
        return data >= inicio && data <= fim;
      });
      const isConf = p => {
        const s = norm(SP.pick(p, "Status") || "");
        return s === "confirmado" || s === "extra" || SP.isTrue(SP.pick(p, "Confirmado"));
      };
      const confirmados   = noIntervalo.filter(isConf);
      const totalSistema  = confirmados.length * valorUnit;
      const porDia = {};
      confirmados.forEach(p => {
        const k = `${SP.pick(p,"Semana_id")||"—"}|${SP.pick(p,"Dia")||"—"}`;
        porDia[k] = (porDia[k]||0)+1;
      });
      const linhas = Object.entries(porDia).sort(([a],[b])=>a.localeCompare(b))
        .map(([k,qt])=>{const[s,d]=k.split("|");return `<tr><td>${AdminUtils.esc(s)}</td><td>${AdminUtils.esc(d)}</td><td>${qt}</td><td>R$ ${(qt*valorUnit).toFixed(2)}</td></tr>`;})
        .join("");
      const diff    = nfTotal !== null ? nfTotal - totalSistema : null;
      const ok      = diff !== null && Math.abs(diff) < 0.02;
      const diffCls = ok ? "badge-green" : "badge-red";
      const diffLbl = ok ? "✅ Sem divergência" :
        diff > 0 ? `⚠️ NF R$ ${Math.abs(diff).toFixed(2)} a mais` : `⚠️ Sistema R$ ${Math.abs(diff).toFixed(2)} a mais`;
      wrap.innerHTML = `
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:.7rem;margin-bottom:1rem">
          <div class="dashboard-mini-card"><div class="dashboard-mini-value">${confirmados.length}</div><div class="dashboard-mini-label">Refeições no sistema</div></div>
          <div class="dashboard-mini-card"><div class="dashboard-mini-value">R$ ${totalSistema.toFixed(2)}</div><div class="dashboard-mini-label">Total calculado</div></div>
          <div class="dashboard-mini-card"><div class="dashboard-mini-value">${nfTotal!==null?`R$ ${nfTotal.toFixed(2)}`:"—"}</div><div class="dashboard-mini-label">Total NF Vascon</div></div>
        </div>
        ${diff!==null?`<div style="margin-bottom:.8rem"><span class="badge ${diffCls}" style="font-size:.8rem;padding:6px 12px">${diffLbl}</span></div>`:""}
        <div style="font-size:.8rem;font-weight:600;color:rgba(143,170,210,.6);text-transform:uppercase;letter-spacing:.08em;margin:.8rem 0 .4rem">Detalhe por dia</div>
        <div class="table-wrap"><table class="table"><thead><tr><th>Semana</th><th>Dia</th><th>Qtd</th><th>Valor</th></tr></thead>
        <tbody>${linhas||`<tr><td colspan="4" class="empty-cell">Nenhum pedido confirmado neste período.</td></tr>`}</tbody></table></div>`;
    } catch (e) {
      wrap.innerHTML = `<div class="alert" style="background:rgba(220,50,50,.1);color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</div>`;
    }
  },

  _bindBotoes() {
    const bind = (id, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundVal) { el.dataset.boundVal = "1"; el.addEventListener("click", fn); }
    };
    bind("btnNovoValor",          () => this.abrirNovo());
    bind("salvarValorRefeicao",   () => this.salvar());
    bind("cancelarValorRefeicao", () => AdminUtils.closeModal("modalValorRefeicao"));
    bind("fecharModalNF",         () => AdminUtils.closeModal("modalNFVascon"));
    bind("btnReconciliar",        () => this._executarReconciliacao());
  }
};
