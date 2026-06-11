// admin-pedidos.js — Pedidos do Admin Homy
// Correção: oculta pedidos quebrados antigos e mostra extras/guarda/investigador na tela de Pedidos.

const AdminPedidos = window.AdminPedidos = {
  _lista: [],
  _extrasVirtuais: [],

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

  _valorValido(v) {
    const s = String(v ?? "").trim();
    if (!s) return false;
    const n = this._norm(s);
    return !["-", "_", "—", "sem", "sem setor", "undefined", "null"].includes(n);
  },

  _nomeGenericoQuebrado(nome) {
    const n = this._norm(nome);
    return !n || ["pedido", "pedidos", "sem nome", "undefined", "null", "-", "—"].includes(n);
  },

  _pedidoTemConteudo(p) {
    const nome = this._pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title");
    const dia = this._pick(p, "Dia");
    const data = this._pick(p, "Data_Hora", "Data", "Data_Referencia");
    const colabId = this._pick(p, "Colaborador_id", "colaboradorId", "ColaboradorId");

    // Registros quebrados antigos vinham como Title/Colaborador = "Pedido",
    // sem dia, sem colaborador real e sem centro de custo. Eles não devem aparecer.
    if (this._nomeGenericoQuebrado(nome)) return false;
    if (!this._valorValido(dia) && !this._valorValido(data)) return false;

    // Se parecer um pedido artificial quebrado, também remove.
    if (this._norm(nome) === "pedido" && !this._valorValido(colabId)) return false;
    return true;
  },

  _isExtraLike(p) {
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
    const nome = this._norm(this._pick(p, "Colaborador_nome", "Nome", "Title"));
    const id = this._norm(this._pick(p, "Colaborador_id", "colaboradorId"));
    return origem.includes("extra") || origem.includes("guarda") || origem.includes("investigador") ||
           origem.includes("visitante") || origem.includes("prestador") || origem.includes("motorista") ||
           nome.includes("refeicao extra") || nome.includes("guarda") || nome.includes("investigador") ||
           id.startsWith("extra-");
  },

  _formatarCC(valor) {
    const raw = String(valor || "").replace(/[–—]/g, " - ").trim();
    if (!raw) return "";
    const m = raw.match(/(\d{6})/);
    if (!m) return raw;
    const codigo = m[1];
    const nome = this._CC_MAPA[codigo] || raw.split(" - ").map(x => x.trim()).find(x => !/^\d{6}$/.test(x)) || "SETOR";
    return `${codigo} - ${nome}`;
  },

  _centroCustoExtra(extra) {
    const direto = this._pick(extra, "Centro_Custo", "CentroCusto", "Setor", "Departamento");
    if (direto) return this._formatarCC(direto);

    const obs = this._pick(extra, "Observacao", "Observação");
    const m = String(obs || "").match(/(\d{6})(?:\s*-\s*([A-Za-zÀ-ÿ\s/.-]+))?/);
    if (m) return this._formatarCC(m[0]);

    // Regras do refeitório: guarda, investigador e extras sem CC entram na Portaria.
    return "120602 - PORTARIA";
  },

  async load(semanaId) {
    await this._buscar(semanaId);
    this._bindFiltros();
    this._bindExportar();
  },

  async _buscar(semanaId) {
    const tbody = document.getElementById("pedidosTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const pedidos = typeof SP.getPedidos === "function" ? await SP.getPedidos(semanaId) : [];
      const validos = (pedidos || []).filter(p => this._pedidoTemConteudo(p));
      const virtuais = await this._extrasComoPedidosVirtuais(semanaId, validos);

      this._lista = [...validos, ...virtuais].sort((a, b) => {
        const da = this._diaOrdem(this._pick(a, "Dia"));
        const db = this._diaOrdem(this._pick(b, "Dia"));
        if (da !== db) return da - db;
        return this._norm(this._pick(a, "Colaborador_nome", "Nome", "Title")).localeCompare(
          this._norm(this._pick(b, "Colaborador_nome", "Nome", "Title")), "pt-BR"
        );
      });

      this._render();
      this._atualizarInfo(`Exibindo pedidos da semana ${semanaId}.`);
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell" style="color:#ff8080">Erro: ${this._esc(e.message || e)}</td></tr>`;
    }
  },

  async _extrasComoPedidosVirtuais(semanaId, pedidosValidos) {
    let extras = [];
    try {
      if (typeof SP.getExtras === "function") extras = await SP.getExtras(semanaId);
      else if (typeof SP.getItems === "function") {
        const todos = await SP.getItems("Extras");
        extras = (todos || []).filter(e => this._pick(e, "Semana_id", "Semana") === semanaId);
      }
    } catch (e) {
      console.warn("[Pedidos] Não foi possível carregar Extras para espelho visual.", e);
      return [];
    }

    if (!extras.length) return [];

    const cardapio = await this._getCardapioSeguro(semanaId);
    const virtuais = [];

    for (const extra of extras) {
      const nome = this._pick(extra, "Nome", "Title") || "Refeição Extra";
      const dia = this._pick(extra, "Dia") || "";
      const tipo = this._pick(extra, "tipo", "Tipo", "Origem") || "Extra";
      const opcao = this._pick(extra, "Opcao", "Opção") || "principal";
      const extraId = String(extra.id || extra.ID || "");

      // Se já existe um pedido espelho válido, não duplica na tela.
      const existe = (pedidosValidos || []).some(p => this._pedidoCorrespondeExtra(p, extra));
      if (existe) continue;

      virtuais.push({
        id: `extra:${extraId || this._norm(`${nome}-${dia}-${tipo}`)}`,
        _virtualExtra: true,
        _extraId: extraId,
        Semana_id: semanaId,
        Colaborador_id: extraId ? `extra-${extraId}` : `extra-${this._norm(`${nome}-${dia}`)}`,
        Colaborador_nome: nome,
        Dia: dia,
        Opcao: opcao,
        Nome_Prato: this._pratoCardapio(cardapio, dia, opcao),
        Status: this._pick(extra, "Status") || "Confirmado",
        Confirmado: true,
        Centro_Custo: this._centroCustoExtra(extra),
        Observacao: this._pick(extra, "Observacao", "Observação") || "Criado no módulo Extras",
        Origem: tipo || "Extra"
      });
    }

    return virtuais;
  },

  _pedidoCorrespondeExtra(p, extra) {
    const pDia = this._norm(this._pick(p, "Dia"));
    const eDia = this._norm(this._pick(extra, "Dia"));
    if (pDia !== eDia) return false;

    const pObs = this._norm(this._pick(p, "Observacao", "Observação"));
    const extraId = String(extra.id || extra.ID || "").toLowerCase();
    if (extraId && pObs.includes(`extraid:${extraId}`)) return true;

    const pId = this._norm(this._pick(p, "Colaborador_id", "colaboradorId"));
    if (extraId && pId === `extra-${extraId}`) return true;

    const pNome = this._norm(this._pick(p, "Colaborador_nome", "Nome", "Title"));
    const eNome = this._norm(this._pick(extra, "Nome", "Title"));
    const eTipo = this._norm(this._pick(extra, "tipo", "Tipo"));

    if (pNome && eNome && pNome === eNome) return true;
    if (pNome && eNome && pNome.includes(eNome)) return true;
    return !!(pNome && eTipo && pNome.includes(eTipo));
  },

  async _getCardapioSeguro(semanaId) {
    try {
      return typeof SP.getCardapio === "function" ? await SP.getCardapio(semanaId) : [];
    } catch (_) {
      return [];
    }
  },

  _pratoCardapio(cardapio, dia, opcao) {
    const item = (cardapio || []).find(c =>
      this._norm(this._pick(c, "Dia")) === this._norm(dia) &&
      this._norm(this._pick(c, "Opcao", "Opção")) === this._norm(opcao)
    );
    return this._pick(item, "Nome_Prato", "Descricao", "Descrição", "Title") || "Cardápio do Dia";
  },

  _diaOrdem(dia) {
    const mapa = { segunda: 1, terca: 2, terça: 2, quarta: 3, quinta: 4, sexta: 5 };
    return mapa[this._norm(dia)] || 9;
  },

  async _buscarPorPeriodo(ini, fim) {
    const tbody = document.getElementById("pedidosTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const todos = await SP.getItems("Pedidos");
      this._lista = (todos || []).filter(p => {
        if (!this._pedidoTemConteudo(p)) return false;
        const d = String(this._pick(p, "Data_Hora") || "").slice(0, 10);
        return d && d >= ini && d <= fim;
      });
      this._render();
      this._atualizarInfo(`Período ${this._brDate(ini)} a ${this._brDate(fim)} — ${this._lista.length} pedidos.`);
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell" style="color:#ff8080">Erro: ${this._esc(e.message || e)}</td></tr>`;
    }
  },

  _render() {
    const tbody = document.getElementById("pedidosTable");
    if (!tbody) return;

    const dia = this._norm(AdminUtils.getVal("filtroDia"));
    const opcao = this._norm(AdminUtils.getVal("filtroOpcao"));

    const lista = this._lista.filter(p => {
      const okDia = !dia || this._norm(this._pick(p, "Dia")) === dia;
      const okOpcao = !opcao || this._norm(this._pick(p, "Opcao")) === opcao;
      return okDia && okOpcao;
    });

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhum pedido encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(p => {
      const id = this._esc(p.id || "");
      const nome = this._esc(this._pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "—");
      const dia = this._esc(this._pick(p, "Dia") || "—");
      const opcao = this._esc(this._pick(p, "Opcao") || "—");
      const prato = this._esc(this._pick(p, "Nome_Prato", "Prato") || "—");
      const status = this._pick(p, "Status") || (SP.isTrue?.(this._pick(p, "Confirmado")) ? "Confirmado" : "Pendente");

      const acoes = p._virtualExtra
        ? `<button class="btn-icon" title="Extra gerenciado em Extras" onclick="AdminPedidos.abrirExtraNaLista()">➕</button>
           ${p._extraId ? `<button class="btn-icon danger" title="Excluir extra" onclick="AdminPedidos.excluirExtra('${this._esc(p._extraId)}')">🗑️</button>` : ""}`
        : `<button class="btn-icon" title="Editar" onclick="AdminPedidos.abrirEdicao('${id}','${nome}','${dia}','${opcao}')">✏️</button>
           <button class="btn-icon danger" title="Excluir" onclick="AdminPedidos.excluir('${id}')">🗑️</button>`;

      return `<tr>
        <td>${nome}</td>
        <td>${dia}</td>
        <td><span class="badge badge-blue">${opcao}</span></td>
        <td>${prato}</td>
        <td>${AdminUtils.badge(status)}</td>
        <td><div class="table-actions">${acoes}</div></td>
      </tr>`;
    }).join("");
  },

  abrirExtraNaLista() {
    AdminUtils.toast("Este registro vem do módulo Extras.", "info");
  },

  async excluirExtra(extraId) {
    if (!confirm("Excluir este extra?")) return;
    try {
      await SP.init();
      if (typeof SP.deleteExtra === "function") await SP.deleteExtra(extraId);
      else await SP.deleteItem("Extras", extraId);
      AdminUtils.toast("Extra excluído.", "success");
      await this._buscar(AdminState.getSemanaId());
    } catch (e) {
      AdminUtils.toast("Erro ao excluir extra: " + (e.message || e), "error");
    }
  },

  _atualizarInfo(texto) {
    const el = document.getElementById("pedidosFiltroInfo");
    if (el) el.innerHTML = `ℹ️ ${this._esc(texto)}`;
  },

  _brDate(iso) {
    return iso ? iso.split("-").reverse().join("/") : "—";
  },

  _editandoId: null,

  abrirEdicao(id, nome, dia, opcao) {
    this._editandoId = id;
    AdminUtils.setVal("editPedidoId", id);
    AdminUtils.setVal("editPedidoNome", nome);
    AdminUtils.setVal("editPedidoDia", dia);
    AdminUtils.setVal("editPedidoOpcao", opcao);
    AdminUtils.openModal("modalEditarPedido");
  },

  async salvarEdicao() {
    const id = AdminUtils.getVal("editPedidoId");
    const dia = AdminUtils.getVal("editPedidoDia");
    const opcao = AdminUtils.getVal("editPedidoOpcao");
    if (!id) { AdminUtils.toast("Pedido inválido.", "error"); return; }

    try {
      await SP.init();
      await SP.updatePedido(id, { dia, opcao, Alterado_Por: SP.getUserName() });
      AdminUtils.closeModal("modalEditarPedido");
      AdminUtils.toast("Pedido atualizado.", "success");
      await this._buscar(AdminState.getSemanaId());
    } catch (e) {
      AdminUtils.toast("Erro ao salvar: " + (e.message || e), "error");
    }
  },

  async excluir(id) {
    if (!confirm("Excluir este pedido?")) return;
    try {
      await SP.init();
      await SP.deletePedido(id);
      AdminUtils.toast("Pedido excluído.", "success");
      this._lista = this._lista.filter(p => String(p.id) !== String(id));
      this._render();
    } catch (e) {
      AdminUtils.toast("Erro ao excluir: " + (e.message || e), "error");
    }
  },

  _bindFiltros() {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundPed) {
        el.dataset.boundPed = "1";
        el.addEventListener(ev, fn);
      }
    };

    bind("btnBuscarPedidosHistorico", "click", async () => {
      const semana = AdminUtils.getVal("filtroSemanaHistorico");
      const ini = AdminUtils.getVal("filtroDataInicioPedidos");
      const fim = AdminUtils.getVal("filtroDataFimPedidos");
      if (semana) await this._buscar(semana);
      else if (ini && fim) await this._buscarPorPeriodo(ini, fim);
      else await this._buscar(AdminState.getSemanaId());
    });

    bind("btnLimparPedidosHistorico", "click", async () => {
      ["filtroSemanaHistorico", "filtroDataInicioPedidos", "filtroDataFimPedidos", "filtroDia", "filtroOpcao"]
        .forEach(id => AdminUtils.setVal(id, ""));
      await this._buscar(AdminState.getSemanaId());
    });

    bind("filtroDia", "change", () => this._render());
    bind("filtroOpcao", "change", () => this._render());
    bind("salvarPedidoEdicao", "click", () => this.salvarEdicao());
    bind("cancelarPedidoEdicao", "click", () => AdminUtils.closeModal("modalEditarPedido"));
  },

  _bindExportar() {
    const btn = document.getElementById("btnExportarPedidos");
    if (btn && !btn.dataset.bound) {
      btn.dataset.bound = "1";
      btn.addEventListener("click", () => this._exportar());
    }
  },

  _exportar() {
    if (!this._lista.length) { AdminUtils.toast("Nenhum dado para exportar.", "info"); return; }
    if (typeof XLSX === "undefined") { AdminUtils.toast("Biblioteca XLSX não carregou.", "error"); return; }

    const linhas = this._lista.map(p => ({
      Semana: this._pick(p, "Semana_id") || "",
      Colaborador: this._pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "",
      Dia: this._pick(p, "Dia") || "",
      Opcao: this._pick(p, "Opcao") || "",
      Prato: this._pick(p, "Nome_Prato") || "",
      Status: this._pick(p, "Status") || "",
      Origem: this._pick(p, "Origem") || "",
      Centro_Custo: this._formatarCC(this._pick(p, "Centro_Custo")) || "",
      DataHora: this._pick(p, "Data_Hora") || ""
    }));

    const ws = XLSX.utils.json_to_sheet(linhas);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pedidos");
    XLSX.writeFile(wb, `pedidos-${AdminState.getSemanaId()}.xlsx`);
    AdminUtils.toast("Excel exportado.", "success");
  }
};
