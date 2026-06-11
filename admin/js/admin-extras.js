// admin-extras.js — Extras / Visitantes do Admin Homy
// Versão limpa: preserva o fluxo original, adiciona Centro de Custo e cria/atualiza pedido espelho corretamente.

const AdminExtras = window.AdminExtras = {
  _lista: [],

  PORTARIA_CC: "120602 - PORTARIA",

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

  async load(semanaId) {
    this._garantirCampoCentroCusto();
    this._bindFiltros(semanaId);
    this._bindBotoes(semanaId);
    await this._carregar(semanaId);
  },

  _norm(v) {
    return String(v || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim();
  },

  _pick(obj, ...keys) {
    for (const k of keys) {
      const v = SP.pick ? SP.pick(obj, k) : obj?.[k];
      if (v !== undefined && v !== null && String(v).trim() !== "") return v;
    }
    return "";
  },

  _esc(v) {
    return AdminUtils.esc ? AdminUtils.esc(v) : String(v ?? "");
  },

  _formatarCC(valor) {
    if (!valor) return this.PORTARIA_CC;
    const v = String(valor).replace(/[–—]/g, " - ").trim();
    const match = v.match(/(\d{6})/);
    const codigo = match ? match[1] : "";
    if (codigo) return `${codigo} - ${this._CC_MAPA[codigo] || this._nomeNoValor(v) || "SETOR"}`;
    return v;
  },

  _nomeNoValor(valor) {
    const partes = String(valor || "").split(" - ").map(x => x.trim()).filter(Boolean);
    return partes.find(x => !/^\d{6}$/.test(x)) || "";
  },

  _centrosOptions() {
    return Object.entries(this._CC_MAPA)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([cod, nome]) => `<option value="${cod} - ${this._esc(nome)}">${cod} - ${this._esc(nome)}</option>`)
      .join("");
  },

  _garantirCampoCentroCusto() {
    if (document.getElementById("extraCentroCusto")) return;

    const tipo = document.getElementById("extraTipo");
    const tipoGroup = tipo?.closest(".form-group");
    if (!tipoGroup) return;

    const grupo = document.createElement("div");
    grupo.className = "form-group";
    grupo.id = "extraCentroCustoGroup";
    grupo.innerHTML = `
      <label class="form-label">Centro de Custo</label>
      <select class="form-select" id="extraCentroCusto">
        ${this._centrosOptions()}
      </select>
    `;
    tipoGroup.insertAdjacentElement("afterend", grupo);

    const select = document.getElementById("extraCentroCusto");
    if (select) select.value = this.PORTARIA_CC;

    tipo.addEventListener("change", () => {
      const t = this._norm(tipo.value);
      const cc = document.getElementById("extraCentroCusto");
      if (!cc) return;
      // Guarda, visitante, motorista, marmita e extras comuns entram por padrão na Portaria.
      // Prestador fica livre para escolher outro centro de custo.
      if (["guarda", "visitante", "motorista", "marmita"].includes(t)) cc.value = this.PORTARIA_CC;
    });
  },

  _setGrupoVisivel(id, visivel) {
    const el = document.getElementById(id);
    const grupo = el?.closest(".form-group");
    if (grupo) grupo.style.display = visivel ? "" : "none";
  },

  _setCentroCusto(valor, travar = false) {
    const cc = document.getElementById("extraCentroCusto");
    if (!cc) return;
    cc.value = this._formatarCC(valor || this.PORTARIA_CC);
    cc.disabled = !!travar;
  },

  _resetModal() {
    ["extraNome", "extraObs"].forEach(id => AdminUtils.setVal(id, ""));
    AdminUtils.setVal("extraTipo", "visitante");
    AdminUtils.setVal("extraDia", "segunda");
    AdminUtils.setVal("extraOpcao", "principal");
    this._setCentroCusto(this.PORTARIA_CC, false);
    ["extraNome", "extraTipo", "extraDia", "extraOpcao", "extraObs", "extraCentroCusto"].forEach(id => this._setGrupoVisivel(id, true));
  },

  async _carregar(semanaId) {
    const tbody = document.getElementById("extrasTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const todos = typeof SP.getExtras === "function" ? await SP.getExtras(semanaId) : await SP.getItems("Extras");

      // Remove duplicatas apenas da refeição extra automática do sistema.
      const seenAuto = new Set();
      this._lista = (todos || []).filter(e => {
        const nome = this._norm(this._pick(e, "Nome", "Title") || "");
        const tipo = this._norm(this._pick(e, "tipo", "Tipo") || "");
        const dia = this._norm(this._pick(e, "Dia") || "");
        const isAuto = nome.includes("refeicao extra") || nome.includes("refeicaoextra") || tipo.includes("extra automatica");
        if (isAuto) {
          const k = `auto-${dia}`;
          if (seenAuto.has(k)) return false;
          seenAuto.add(k);
        }
        return true;
      });

      const diaOrd = { segunda: 1, terca: 2, quarta: 3, quinta: 4, sexta: 5 };
      this._lista.sort((a, b) => {
        const da = diaOrd[this._norm(this._pick(a, "Dia"))] || 9;
        const db = diaOrd[this._norm(this._pick(b, "Dia"))] || 9;
        if (da !== db) return da - db;
        return this._norm(this._pick(a, "tipo", "Tipo")).localeCompare(this._norm(this._pick(b, "tipo", "Tipo")), "pt-BR");
      });

      this._render();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell" style="color:#ff8080">Erro: ${this._esc(e.message || e)}</td></tr>`;
    }
  },

  _render() {
    const tbody = document.getElementById("extrasTable");
    if (!tbody) return;

    const txt  = this._norm(AdminUtils.getVal("fExtraTexto"));
    const dia  = this._norm(AdminUtils.getVal("fExtraDia"));
    const tipo = this._norm(AdminUtils.getVal("fExtraTipo"));

    const lista = this._lista.filter(e => {
      const all = this._norm([
        this._pick(e, "Nome", "Title"), this._pick(e, "Observacao"),
        this._pick(e, "tipo", "Tipo"), this._pick(e, "Dia")
      ].join(" "));
      return (!txt || all.includes(txt)) &&
             (!dia || this._norm(this._pick(e, "Dia")) === dia) &&
             (!tipo || this._norm(this._pick(e, "tipo", "Tipo")).includes(tipo));
    });

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhum extra encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(e => {
      const id = this._esc(e.id || "");
      const nome = this._esc(this._pick(e, "Nome", "Title") || "Extra");
      const tipo = this._esc(this._pick(e, "tipo", "Tipo") || "—");
      const dia = this._esc(this._pick(e, "Dia") || "—");
      const opc = this._esc(this._pick(e, "Opcao") || "principal");
      const obs = this._esc(this._pick(e, "Observacao") || "—");
      return `<tr>
        <td>${nome}</td>
        <td><span class="badge badge-yellow">${tipo}</span></td>
        <td>${dia}</td>
        <td><span class="badge badge-blue">${opc}</span></td>
        <td style="font-size:.78rem;color:rgba(143,170,210,.7)">${obs}</td>
        <td><div class="table-actions">
          <button class="btn-icon danger" title="Excluir" onclick="AdminExtras.excluir('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  abrirModal(predefinido = "") {
    this._garantirCampoCentroCusto();
    const modal = document.getElementById("modalExtra");
    if (!modal) return;
    modal.dataset.predefinido = predefinido;
    this._resetModal();

    const title = modal.querySelector(".modal-title");

    if (predefinido === "investigador") {
      if (title) title.textContent = "Adicionar Investigador";
      AdminUtils.setVal("extraNome", "Investigador");
      AdminUtils.setVal("extraTipo", "investigador");
      AdminUtils.setVal("extraOpcao", "principal");
      AdminUtils.setVal("extraObs", "Investigador");
      this._setCentroCusto(this.PORTARIA_CC, true);
      // Mantém Dia e Opção visíveis; oculta só o que é fixo.
      this._setGrupoVisivel("extraNome", false);
      this._setGrupoVisivel("extraTipo", false);
      this._setGrupoVisivel("extraObs", false);
    } else if (predefinido === "guarda") {
      if (title) title.textContent = "Adicionar Guarda";
      AdminUtils.setVal("extraNome", "Guarda");
      AdminUtils.setVal("extraTipo", "guarda");
      AdminUtils.setVal("extraOpcao", "principal");
      AdminUtils.setVal("extraObs", `Guarda — Centro de custo ${this.PORTARIA_CC}`);
      this._setCentroCusto(this.PORTARIA_CC, true);
      // Mantém Dia e Opção visíveis; oculta só o que é fixo.
      this._setGrupoVisivel("extraNome", false);
      this._setGrupoVisivel("extraTipo", false);
      this._setGrupoVisivel("extraObs", false);
    } else {
      if (title) title.textContent = "Adicionar Extra";
      this._setCentroCusto(this.PORTARIA_CC, false);
    }

    AdminUtils.openModal("modalExtra");
  },

  async _pratoPorOpcao(semanaId, dia, opcao) {
    try {
      if (typeof SP.getCardapio !== "function") return "Cardápio do Dia";
      const cardapio = await SP.getCardapio(semanaId);
      const item = (cardapio || []).find(c =>
        this._norm(this._pick(c, "Dia")) === this._norm(dia) &&
        this._norm(this._pick(c, "Opcao", "Opção")) === this._norm(opcao)
      );
      return this._pick(item, "Nome_Prato", "Descricao", "Descrição", "Title") || "Cardápio do Dia";
    } catch (_) {
      return "Cardápio do Dia";
    }
  },

  _dataHoraDoDia(semanaId, dia) {
    try {
      if (typeof SP.getDataRefBySemanaDia === "function") {
        return `${SP.getDataRefBySemanaDia(semanaId, dia)}T12:00:00`;
      }
    } catch (_) {}
    return new Date().toISOString();
  },

  async _criarExtra(semanaId, dia, nome, tipo, opcao, obs) {
    if (typeof SP.addExtra === "function") {
      return await SP.addExtra(semanaId, dia, nome, tipo, opcao, obs, SP.getUserName ? SP.getUserName() : "Admin");
    }
    return await SP.createItem("Extras", {
      Title: `${semanaId}-${dia}-${nome}`,
      Semana_id: semanaId,
      Dia: dia,
      Nome: nome,
      tipo: tipo,
      Opcao: opcao || "principal",
      Observacao: obs || "",
      Adicionado_Por: SP.getUserName ? SP.getUserName() : "Admin"
    });
  },

  _pedidoExtraValido(p, dia, nome, tipo) {
    const pDia = this._norm(this._pick(p, "Dia"));
    const pNome = this._norm(this._pick(p, "Colaborador_nome", "Nome", "Title"));
    const pOrigem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
    if (pDia !== this._norm(dia)) return false;
    if (!pNome) return false;
    if (pNome === this._norm(nome)) return true;
    if (pNome.includes(this._norm(nome))) return true;
    return pOrigem.includes(this._norm(tipo)) && pNome.includes(this._norm(tipo));
  },

  async _salvarPedidoEspelho(semanaId, dia, nome, tipo, opcao, obs, centroCusto, extraId = "") {
    const pedidos = typeof SP.getPedidos === "function" ? await SP.getPedidos(semanaId).catch(() => []) : [];
    const existente = pedidos.find(p => this._pedidoExtraValido(p, dia, nome, tipo));
    const nomePrato = await this._pratoPorOpcao(semanaId, dia, opcao);
    const cc = this._formatarCC(centroCusto || this.PORTARIA_CC);
    const dataHora = this._dataHoraDoDia(semanaId, dia);
    const colabId = extraId ? `extra-${extraId}` : `extra-${this._norm(tipo)}-${this._norm(nome)}-${this._norm(dia)}`;
    const observacao = [obs || "", extraId ? `ExtraID:${extraId}` : ""].filter(Boolean).join(" | ");

    const dados = {
      Semana_id: semanaId,
      semanaId,
      Colaborador_id: colabId,
      colaboradorId: colabId,
      Colaborador_nome: nome,
      colaboradorNome: nome,
      Dia: dia,
      dia,
      Opcao: opcao || "principal",
      opcao: opcao || "principal",
      Nome_Prato: nomePrato,
      nomePrato,
      Confirmado: true,
      confirmado: true,
      Data_Hora: dataHora,
      dataHora,
      Centro_Custo: cc,
      centroCusto: cc,
      Status: "Confirmado",
      status: "Confirmado",
      Observacao: observacao,
      observacao,
      Origem: tipo || "Extra",
      origem: tipo || "Extra",
      Alterado_Por: SP.getUserName ? SP.getUserName() : "Admin",
      alteradoPor: SP.getUserName ? SP.getUserName() : "Admin"
    };

    if (existente?.id && typeof SP.updatePedido === "function") {
      return await SP.updatePedido(existente.id, dados);
    }

    if (typeof SP.savePedido === "function") {
      return await SP.savePedido(semanaId, colabId, nome, dia, opcao || "principal", nomePrato, {
        confirmado: true,
        status: "Confirmado",
        dataHora,
        centroCusto: cc,
        origem: tipo || "Extra",
        observacao,
        alteradoPor: SP.getUserName ? SP.getUserName() : "Admin"
      });
    }

    return await SP.createItem("Pedidos", {
      Title: `${semanaId}-${colabId}-${dia}`,
      Semana_id: semanaId,
      Colaborador_id: colabId,
      Colaborador_nome: nome,
      Dia: dia,
      Opcao: opcao || "principal",
      Nome_Prato: nomePrato,
      Confirmado: true,
      Data_Hora: dataHora,
      Centro_Custo: cc,
      Status: "Confirmado",
      Observacao: observacao,
      Origem: tipo || "Extra",
      Alterado_Por: SP.getUserName ? SP.getUserName() : "Admin"
    });
  },

  async salvar() {
    const modal = document.getElementById("modalExtra");
    const predefinido = modal?.dataset.predefinido || "";
    const semanaId = AdminState.getSemanaId();
    const dia = AdminUtils.getVal("extraDia") || "segunda";
    const centroCusto = this._formatarCC(AdminUtils.getVal("extraCentroCusto") || this.PORTARIA_CC);

    let nome, tipo, opcao, obs;

    if (predefinido === "investigador") {
      const pedidos = await SP.getPedidos(semanaId).catch(() => []);
      const qtd = pedidos.filter(p =>
        this._norm(this._pick(p, "Dia")) === this._norm(dia) &&
        this._norm(this._pick(p, "Origem", "tipo", "Tipo", "Colaborador_nome")).includes("investigador")
      ).length;
      nome = `Investigador ${qtd + 1}`;
      tipo = "investigador";
      opcao = AdminUtils.getVal("extraOpcao") || "principal";
      obs = `Investigador — Centro de custo ${centroCusto}`;
    } else if (predefinido === "guarda") {
      const pedidos = await SP.getPedidos(semanaId).catch(() => []);
      const qtd = pedidos.filter(p =>
        this._norm(this._pick(p, "Dia")) === this._norm(dia) &&
        this._norm(this._pick(p, "Origem", "tipo", "Tipo", "Colaborador_nome")).includes("guarda")
      ).length;
      nome = qtd > 0 ? `Guarda ${qtd + 1}` : "Guarda";
      tipo = "guarda";
      opcao = AdminUtils.getVal("extraOpcao") || "principal";
      obs = `Guarda — Centro de custo ${this.PORTARIA_CC}`;
    } else {
      nome = AdminUtils.getVal("extraNome");
      tipo = AdminUtils.getVal("extraTipo") || "visitante";
      opcao = AdminUtils.getVal("extraOpcao") || "principal";
      obs = AdminUtils.getVal("extraObs");
      if (!obs) obs = `Centro de custo ${centroCusto}`;
    }

    if (!nome) { AdminUtils.toast("Informe o nome do extra.", "error"); return; }

    const btn = document.getElementById("salvarExtra");
    const old = btn?.textContent;
    if (btn) { btn.disabled = true; btn.textContent = "⏳ Salvando..."; }

    try {
      await SP.init();
      const extraCriado = await this._criarExtra(semanaId, dia, nome, tipo, opcao, obs);
      const extraId = extraCriado?.id || extraCriado?.fields?.id || "";
      await this._salvarPedidoEspelho(semanaId, dia, nome, tipo, opcao, obs, centroCusto, extraId);

      AdminUtils.closeModal("modalExtra");
      AdminUtils.toast("Extra adicionado e refletido em Pedidos/Operação.", "success");
      await this._carregar(semanaId);
    } catch (e) {
      console.error("[Extras] salvar", e);
      AdminUtils.toast("Erro ao salvar extra: " + (e.message || e), "error");
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = old || "💾 Adicionar"; }
    }
  },

  async excluir(id) {
    if (!confirm("Excluir este extra?")) return;
    try {
      await SP.init();
      const item = this._lista.find(e => String(e.id) === String(id));
      if (typeof SP.deleteExtra === "function") await SP.deleteExtra(id);
      else await SP.deleteItem("Extras", id);

      // Remove pedido espelho correspondente quando possível.
      if (item && typeof SP.getPedidos === "function" && typeof SP.deletePedido === "function") {
        const semanaId = AdminState.getSemanaId();
        const nome = this._norm(this._pick(item, "Nome", "Title"));
        const dia = this._norm(this._pick(item, "Dia"));
        const pedidos = await SP.getPedidos(semanaId).catch(() => []);
        const espelhos = pedidos.filter(p =>
          this._norm(this._pick(p, "Dia")) === dia &&
          this._norm(this._pick(p, "Colaborador_nome", "Nome", "Title")) === nome &&
          this._norm(this._pick(p, "Origem", "tipo", "Tipo")).includes(this._norm(this._pick(item, "tipo", "Tipo")))
        );
        for (const p of espelhos) await SP.deletePedido(p.id).catch(() => null);
      }

      this._lista = this._lista.filter(e => String(e.id) !== String(id));
      this._render();
      AdminUtils.toast("Extra excluído.", "success");
    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + (e.message || e), "error");
    }
  },

  _bindFiltros(semanaId) {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundExt) { el.dataset.boundExt = "1"; el.addEventListener(ev, fn); }
    };
    bind("fExtraTexto", "input", () => this._render());
    bind("fExtraDia", "change", () => this._render());
    bind("fExtraTipo", "change", () => this._render());
    bind("btnFiltrarExtras", "click", () => this._render());
    bind("btnLimparExtras", "click", () => {
      ["fExtraTexto", "fExtraDia", "fExtraTipo"].forEach(id => AdminUtils.setVal(id, ""));
      this._render();
    });
  },

  _bindBotoes(semanaId) {
    const bind = (id, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundExtBtn) { el.dataset.boundExtBtn = "1"; el.addEventListener("click", fn); }
    };
    bind("btnAdicionarExtra", () => this.abrirModal());
    bind("btnExtraInvestigador", () => this.abrirModal("investigador"));
    bind("btnExtraGuarda", () => this.abrirModal("guarda"));
    bind("salvarExtra", () => this.salvar());
    bind("cancelModalExtra", () => AdminUtils.closeModal("modalExtra"));
    bind("closeModalExtra", () => AdminUtils.closeModal("modalExtra"));
  }
};
