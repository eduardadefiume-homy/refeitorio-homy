// admin-operacao-dia.js — Operação do Dia do Admin Homy
// Correções: centro de custo no padrão "120501 - TI" + extras/guarda aparecem no dia correto.

const AdminOperacao = window.AdminOperacao = {

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
    this._bindControles(semanaId);
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
      if (v !== undefined && v !== null && String(v) !== "") return v;
    }
    return "";
  },

  _esc(v) {
    return AdminUtils.esc ? AdminUtils.esc(v) : String(v ?? "");
  },

  _formatarCC(valor) {
    if (!valor) return "Sem setor";
    const v = String(valor).replace(/[–—]/g, " - ").trim();
    if (!v || v === "—") return "Sem setor";

    const match = v.match(/(\d{6})/);
    const codigo = match ? match[1] : "";
    if (codigo) {
      const nomeMapa = this._CC_MAPA[codigo];
      if (nomeMapa) return `${codigo} - ${nomeMapa}`;

      const partes = v.split(" - ").map(x => x.trim()).filter(Boolean);
      const nome = partes.find(x => !/^\d{6}$/.test(x));
      return nome ? `${codigo} - ${nome}` : codigo;
    }

    return v;
  },

  _centroCustoPedido(p) {
    const raw = this._pick(p, "Centro_Custo", "CentroCusto", "Setor", "Departamento");
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
    const nome = this._norm(this._pick(p, "Colaborador_nome", "Title", "Nome"));
    const extraPortaria = origem.includes("guarda") || nome.includes("guarda") || origem.includes("visitante") || origem.includes("motorista") || origem.includes("prestador");
    if (extraPortaria) return this.PORTARIA_CC;
    if ((!raw || this._norm(raw).includes("sem setor")) && (origem.includes("extra") || nome.includes("refeicao extra"))) return this.PORTARIA_CC;
    return this._formatarCC(raw || "Sem setor");
  },

  _isExtraPedido(p) {
    if (SP.isExtraPedido) return SP.isExtraPedido(p);
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
    const nome = this._norm(this._pick(p, "Colaborador_nome", "Title", "Nome"));
    return origem.includes("extra") || origem.includes("guarda") || origem.includes("investigador") || origem.includes("visitante") || nome.includes("refeicao extra") || nome.includes("guarda");
  },

  _isExtraAutomatico(p) {
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
    const nome = this._norm(this._pick(p, "Colaborador_nome", "Title", "Nome"));
    const obs = this._norm(this._pick(p, "Observacao", "Observação"));
    return nome.includes("refeicao extra") || origem.includes("extra automatica") || obs.includes("extra automatica");
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

  _extraJaTemPedido(extra, pedidos, dia) {
    const nome = this._norm(this._pick(extra, "Nome", "Title"));
    const tipo = this._norm(this._pick(extra, "tipo", "Tipo", "Origem"));
    return pedidos.some(p => {
      const pDia = this._norm(this._pick(p, "Dia"));
      const pNome = this._norm(this._pick(p, "Colaborador_nome", "Title", "Nome"));
      const pOrigem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
      if (pDia !== this._norm(dia)) return false;
      if (nome && pNome === nome) return true;
      if (nome && pNome.includes(nome)) return true;
      if (tipo && pOrigem.includes(tipo) && pNome.includes(tipo)) return true;
      return false;
    });
  },

  async _criarPedidoDoExtra(semanaId, dia, extra) {
    const nome = this._pick(extra, "Nome", "Title") || "Refeição Extra";
    const tipo = this._pick(extra, "tipo", "Tipo") || "Extra";
    const opcao = this._pick(extra, "Opcao", "Opção") || "principal";
    const obs = this._pick(extra, "Observacao", "Observação") || "Criado automaticamente a partir do módulo Extras";
    const centroCusto = this.PORTARIA_CC;
    const nomePrato = await this._pratoPorOpcao(semanaId, dia, opcao);
    const colabId = `extra-${this._norm(tipo || nome)}-${extra.id || Date.now()}`;

    if (typeof SP.savePedido === "function") {
      return SP.savePedido(semanaId, colabId, nome, dia, opcao, nomePrato, {
        confirmado: true,
        status: "Confirmado",
        centroCusto,
        origem: tipo,
        observacao: obs,
        alteradoPor: SP.getUserName ? SP.getUserName() : "Admin"
      });
    }

    return SP.createItem("Pedidos", {
      Title: `${semanaId}-${colabId}-${dia}`,
      Semana_id: semanaId,
      Colaborador_id: colabId,
      Colaborador_nome: nome,
      Dia: dia,
      Opcao: opcao,
      Nome_Prato: nomePrato,
      Confirmado: true,
      Data_Hora: new Date().toISOString(),
      Centro_Custo: centroCusto,
      Status: "Confirmado",
      Observacao: obs,
      Origem: tipo,
      Alterado_Por: SP.getUserName ? SP.getUserName() : "Admin"
    });
  },

  _pedidoVirtualExtra(semanaId, dia, extra) {
    const nome = this._pick(extra, "Nome", "Title") || "Refeição Extra";
    const tipo = this._pick(extra, "tipo", "Tipo") || "Extra";
    const opcao = this._pick(extra, "Opcao", "Opção") || "principal";
    return {
      id: "",
      _virtualExtra: true,
      Semana_id: semanaId,
      Colaborador_nome: nome,
      Dia: dia,
      Opcao: opcao,
      Nome_Prato: "Cardápio do Dia",
      Confirmado: true,
      Centro_Custo: this.PORTARIA_CC,
      Status: "Confirmado",
      Origem: tipo,
      Observacao: this._pick(extra, "Observacao", "Observação") || "Extra ainda sem pedido espelho"
    };
  },

  async _sincronizarExtrasParaPedidos(semanaId, dia, pedidos) {
    if (typeof SP.getExtras !== "function") return pedidos;

    const extras = await SP.getExtras(semanaId, dia).catch(() => []);
    if (!extras.length) return pedidos;

    const pendentes = extras.filter(e => !this._extraJaTemPedido(e, pedidos, dia));
    if (!pendentes.length) return pedidos;

    const virtuais = [];
    for (const extra of pendentes) {
      try {
        await this._criarPedidoDoExtra(semanaId, dia, extra);
      } catch (e) {
        console.warn("[Operação] Não foi possível criar pedido espelho do extra", extra, e);
        virtuais.push(this._pedidoVirtualExtra(semanaId, dia, extra));
      }
    }

    try {
      const atualizados = await SP.getPedidos(semanaId);
      return [...atualizados, ...virtuais];
    } catch (_) {
      return [...pedidos, ...virtuais];
    }
  },

  async _carregar(semanaId) {
    const tbody = document.getElementById("operacaoTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const dia = AdminUtils.getVal("operacaoDia") || AdminUtils.DIA_HOJE();
      let pedidos = await SP.getPedidos(semanaId);
      pedidos = await this._sincronizarExtrasParaPedidos(semanaId, dia, pedidos);

      const seenAuto = new Set();
      this._lista = (pedidos || []).filter(p => {
        if (this._norm(this._pick(p, "Dia")) !== this._norm(dia)) return false;

        // Deduplica somente a refeição extra automática do sistema.
        // Não remove Guarda, Investigador, Visitante ou outros extras manuais.
        if (this._isExtraAutomatico(p)) {
          const k = `auto-${this._norm(dia)}`;
          if (seenAuto.has(k)) return false;
          seenAuto.add(k);
        }
        return true;
      });

      this._renderTotais(dia);
      this._renderTabela();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell" style="color:#ff8080">Erro: ${this._esc(e.message || e)}</td></tr>`;
    }
  },

  _renderTotais(dia) {
    const STATUS_PROD = ["confirmado", "extra", "aprovado"];
    const STATUS_CANC = ["cancelado", "afastado", "ferias", "nao vai almocar", "bloqueado", "travado"];

    const isConf = p => STATUS_PROD.includes(this._norm(this._pick(p, "Status") || "")) || (SP.isTrue && SP.isTrue(this._pick(p, "Confirmado")));
    const isCanc = p => STATUS_CANC.includes(this._norm(this._pick(p, "Status") || ""));
    const conf = this._lista.filter(isConf);

    const setCard = (id, val) => AdminUtils.setTxt(id, val);
    setCard("opTotalConfirmado", conf.length);
    setCard("opTotalPrincipal", conf.filter(p => this._norm(this._pick(p, "Opcao")) === "principal").length);
    setCard("opTotalLight", conf.filter(p => this._norm(this._pick(p, "Opcao")) === "light").length);
    setCard("opTotalCarne", conf.filter(p => this._norm(this._pick(p, "Opcao")) === "carne").length);
    setCard("opTotalMassa", conf.filter(p => this._norm(this._pick(p, "Opcao")) === "massa").length);
    setCard("opTotalCancelado", this._lista.filter(isCanc).length);

    const cardLanche = document.getElementById("cardOpLanche");
    if (cardLanche) {
      cardLanche.style.display = this._norm(dia) === "sexta" ? "" : "none";
      setCard("opTotalLanche", conf.filter(p => this._norm(this._pick(p, "Opcao")) === "lanche").length);
    }
  },

  _renderTabela() {
    const tbody = document.getElementById("operacaoTable");
    if (!tbody) return;

    const statusFiltro = this._norm(AdminUtils.getVal("operacaoFiltroStatus"));
    const busca = this._norm(AdminUtils.getVal("operacaoBusca"));

    let lista = this._lista;
    if (statusFiltro) lista = lista.filter(p => this._norm(this._pick(p, "Status") || "") === statusFiltro);
    if (busca) lista = lista.filter(p =>
      this._norm([
        this._pick(p, "Colaborador_nome", "Title", "Nome"),
        this._centroCustoPedido(p),
        this._pick(p, "Origem", "tipo", "Tipo")
      ].join(" ")).includes(busca)
    );

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Nenhum pedido encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(p => {
      const id = this._esc(p.id || "");
      const nome = this._esc(this._pick(p, "Colaborador_nome", "Title", "Nome") || "—");
      const cc = this._esc(this._centroCustoPedido(p));
      const opcao = this._esc(this._pick(p, "Opcao") || "—");
      const prato = this._esc(this._pick(p, "Nome_Prato") || "—");
      const status = this._pick(p, "Status") || "Pendente";
      const origem = this._esc(this._pick(p, "Origem", "tipo", "Tipo") || "Refeitório");
      const isEx = this._isExtraPedido(p);
      const disabled = p._virtualExtra ? "disabled title='Extra sem pedido espelho no SharePoint'" : "";

      return `<tr>
        <td${isEx ? ' style="color:#ffd36d;font-weight:700"' : ""}>${nome}</td>
        <td>${cc}</td>
        <td><span class="badge badge-blue">${opcao}</span></td>
        <td>${prato}</td>
        <td>${AdminUtils.badge(status)}</td>
        <td>${origem}</td>
        <td><div class="table-actions">
          <button class="btn-icon" ${disabled} title="Confirmar" onclick="AdminOperacao.alterarStatus('${id}','Confirmado')">✅</button>
          <button class="btn-icon danger" ${disabled} title="Cancelar" onclick="AdminOperacao.alterarStatus('${id}','Cancelado')">❌</button>
          <button class="btn-icon" ${disabled} title="Não vai almoçar" onclick="AdminOperacao.alterarStatus('${id}','Não vai almoçar')">🚫</button>
          <button class="btn-icon danger" ${disabled} title="Excluir" onclick="AdminOperacao.excluir('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  async alterarStatus(id, status) {
    if (!id) return;
    try {
      await SP.init();
      await SP.updatePedido(id, {
        Status: status,
        Confirmado: ["Confirmado", "Extra"].includes(status),
        Alterado_Por: SP.getUserName(),
        Origem: "Admin"
      });
      const p = this._lista.find(x => String(x.id) === String(id));
      if (p) { p.Status = status; p.Confirmado = ["Confirmado", "Extra"].includes(status); }
      const dia = AdminUtils.getVal("operacaoDia") || AdminUtils.DIA_HOJE();
      this._renderTotais(dia);
      this._renderTabela();
      AdminUtils.toast(`Status: ${status}`, "success");
    } catch (e) {
      AdminUtils.toast("Erro: " + (e.message || e), "error");
    }
  },

  async excluir(id) {
    if (!id) return;
    if (!confirm("Excluir este pedido?")) return;
    try {
      await SP.init();
      await SP.deletePedido(id);
      this._lista = this._lista.filter(p => String(p.id) !== String(id));
      const dia = AdminUtils.getVal("operacaoDia") || AdminUtils.DIA_HOJE();
      this._renderTotais(dia);
      this._renderTabela();
      AdminUtils.toast("Pedido excluído.", "success");
    } catch (e) {
      AdminUtils.toast("Erro: " + (e.message || e), "error");
    }
  },

  _bindControles(semanaId) {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundOp) { el.dataset.boundOp = "1"; el.addEventListener(ev, fn); }
    };

    const diaEl = document.getElementById("operacaoDia");
    if (diaEl && !diaEl.dataset.boundOp && !diaEl.value) diaEl.value = AdminUtils.DIA_HOJE();

    bind("operacaoDia", "change", () => this._carregar(AdminState.getSemanaId()));
    bind("operacaoFiltroStatus", "change", () => this._renderTabela());
    bind("operacaoBusca", "input", () => this._renderTabela());
    bind("btnRecarregarOperacao", "click", () => this._carregar(AdminState.getSemanaId()));
  }
};
