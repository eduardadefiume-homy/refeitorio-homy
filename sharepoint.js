// ============================================================
// sharepoint.js — Refeitório Homy · Microsoft Graph API
// v: homy-final-20260610-1
// ============================================================

const SP = {
  clientId:  "aa37acf9-f3bd-4d1e-968a-fde57f79094c",
  tenantId:  "a2850abc-334a-4805-b6b2-420b4aef68a9",
  siteUrl:   "homyquimica.sharepoint.com",
  sitePath:  "/sites/Refeitrio-Homy",
  scopes:    ["https://graph.microsoft.com/Sites.ReadWrite.All", "User.Read"],

  // redirectUri fixo na raiz — NUNCA window.location.pathname
  redirectUri: "https://eduardadefiume-homy.github.io/refeitorio-homy/index.html",

  _msalInstance: null,
  _account:      null,
  _siteId:       null,
  _listIds:      {},

  // ============================================================
  // UTILITÁRIOS
  // ============================================================
  pick(obj, ...keys) {
    for (const key of keys) {
      if (obj && obj[key] !== undefined && obj[key] !== null) return obj[key];
    }
    return null;
  },

  isTrue(value) {
    if (value === true || value === 1) return true;
    const v = String(value ?? "").trim().toLowerCase();
    return v === "sim" || v === "true" || v === "yes" || v === "1";
  },

  getSemanaId(date = new Date()) {
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
    const week1 = new Date(d.getFullYear(), 0, 4);
    const weekNum = 1 + Math.round(
      ((d - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7
    );
    return `${d.getFullYear()}-W${String(weekNum).padStart(2, "0")}`;
  },

  getCurrentWeekId() {
    return this.getSemanaId(new Date());
  },

  getWeekDates(semanaId) {
    const [year, week] = semanaId.split("-W").map(Number);
    const jan4  = new Date(year, 0, 4);
    const start = new Date(jan4);
    start.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1 + (week - 1) * 7);
    return Array.from({ length: 5 }, (_, i) => {
      const d = new Date(start);
      d.setDate(start.getDate() + i);
      return d;
    });
  },

  getDataRefBySemanaDia(semanaId, dia) {
    const ordem = { segunda: 0, terca: 1, "terça": 1, quarta: 2, quinta: 3, sexta: 4 };
    const norm  = v => String(v || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    const datas = this.getWeekDates(semanaId);
    const idx   = ordem[norm(dia)];
    if (idx === undefined || !datas[idx]) return new Date().toISOString().slice(0, 10);
    return datas[idx].toISOString().slice(0, 10);
  },

  getUserName()  { return this._account?.name     || this._account?.username || "Usuário Homy"; },
  getUserEmail() { return this._account?.username || ""; },

  isExtraPedido(p) {
    const norm   = v => String(v || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    const origem = norm(this.pick(p, "Origem", "tipo", "Tipo") || "");
    const nome   = norm(this.pick(p, "Colaborador_nome", "Title", "Nome") || "");
    return origem.includes("extra") || origem.includes("investigador") ||
           origem.includes("guarda") || nome.includes("refeicao extra");
  },

  // ============================================================
  // AUTENTICAÇÃO — popup (NUNCA redirect em GitHub Pages)
  // cacheLocation: "localStorage" para compartilhar sessão entre
  // páginas da mesma origem sem exigir novo login
  // ============================================================
  async init() {
    if (this._msalInstance) return !!this._account;

    if (!window.msal) {
      throw new Error("MSAL não carregou. Verifique se msal-browser.min.js está antes de sharepoint.js.");
    }

    this._msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId:                  this.clientId,
        authority:                 `https://login.microsoftonline.com/${this.tenantId}`,
        redirectUri:               this.redirectUri,
        navigateToLoginRequestUrl: false
      },
      // localStorage compartilha o token entre cardapio-semana, marcar-refeicao,
      // cardapio-dia, cozinha — o usuário faz login uma única vez
      cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
    });

    await this._msalInstance.initialize();

    // Nunca usar handleRedirectPromise com popup — pode conflitar
    // Reaproveitamos apenas conta já existente no cache
    const active = this._msalInstance.getActiveAccount();
    if (active) { this._account = active; return true; }

    const accounts = this._msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      this._account = accounts[0];
      this._msalInstance.setActiveAccount(this._account);
      return true;
    }

    return false;
  },

  async login() {
    await this.init();
    const result = await this._msalInstance.loginPopup({
      scopes: this.scopes,
      prompt: "select_account"
    });
    if (!result?.account) throw new Error("Login retornou sem conta.");
    this._account = result.account;
    this._msalInstance.setActiveAccount(this._account);
    return true;
  },

  async ensureLogin() {
    await this.init();
    if (!this._account) await this.login();
    return true;
  },

  async logout() {
    await this.init();
    const account = this._account;
    this._account = null;
    this._siteId  = null;
    this._listIds = {};
    if (account) await this._msalInstance.logoutPopup({ account });
  },

  async getToken() {
    await this.init();
    if (!this._account) await this.login();

    try {
      const r = await this._msalInstance.acquireTokenSilent({
        scopes:  this.scopes,
        account: this._account
      });
      return r.accessToken;
    } catch (e) {
      const r = await this._msalInstance.acquireTokenPopup({ scopes: this.scopes });
      this._account = r.account || this._account;
      return r.accessToken;
    }
  },

  // ============================================================
  // GRAPH — camada HTTP base
  // ============================================================
  async graph(method, endpoint, body = null) {
    const token = await this.getToken();
    if (!token) return null;

    let safeBody = body;
    if (body && (method === "PATCH" || method === "POST")) {
      if (body.fields) {
        safeBody = { ...body, fields: this._cleanFields(body.fields) };
      } else {
        safeBody = this._cleanFields(body);
      }
    }

    const res = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
      method,
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type":  "application/json"
      },
      body: safeBody ? JSON.stringify(safeBody) : null
    });

    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Graph ${method} ${endpoint} → ${res.status}: ${err}`);
    }
    if (res.status === 204) return null;
    return res.json();
  },

  // Remove campos read-only antes de PATCH/POST
  _cleanFields(fields) {
    const READONLY = [
      "@odata.etag", "ComplianceAssetId", "AuthorId", "EditorId",
      "Created", "Modified", "id", "ID", "odata.type",
      "odata.id", "odata.editLink", "FileSystemObjectType",
      "ServerRedirectedEmbedUri", "ServerRedirectedEmbedUrl",
      "ContentTypeId", "OData__UIVersionString",
      "Attachments", "GUID", "_ModerationStatus",
      "_ModerationComments", "LinkTitleNoMenu", "LinkTitle"
    ];
    const out = {};
    for (const [k, v] of Object.entries(fields || {})) {
      if (!READONLY.includes(k) && !k.startsWith("@odata")) {
        out[k] = v;
      }
    }
    return out;
  },

  // ============================================================
  // SITE E LISTAS
  // ============================================================
  async getSiteId() {
    if (this._siteId) return this._siteId;

    const paths = [
      this.sitePath,
      "/sites/Refeitorio-Homy",
      "/sites/refeitorio-homy",
      "/sites/Refeitrio-Homy"
    ].filter(Boolean);

    const unique  = [...new Set(paths)];
    let lastErr   = null;

    for (const path of unique) {
      try {
        const data = await this.graph("GET", `/sites/${this.siteUrl}:${path}`);
        if (data?.id) {
          this._siteId  = data.id;
          this.sitePath = path;
          return this._siteId;
        }
      } catch (e) { lastErr = e; }
    }

    throw lastErr || new Error("Site SharePoint não encontrado.");
  },

  async getListId(listName) {
    if (this._listIds[listName]) return this._listIds[listName];

    const siteId = await this.getSiteId();
    const data   = await this.graph("GET", `/sites/${siteId}/lists?$select=id,displayName`);
    const list   = (data?.value || []).find(l => l.displayName === listName);
    if (!list) throw new Error(`Lista "${listName}" não encontrada.`);

    this._listIds[listName] = list.id;
    return list.id;
  },

  async getItems(listName) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);

    let   items    = [];
    let   endpoint = `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500`;

    while (endpoint) {
      const data = await this.graph("GET", endpoint);
      items.push(...(data?.value || []));
      endpoint = data?.["@odata.nextLink"]
        ? data["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "")
        : null;
    }

    return items.map(i => ({ id: i.id, ...i.fields }));
  },

  _validarPedidoFields(fields = {}) {
    const semana = String(fields.Semana_id || fields.semana_id || "").trim();
    const dia = String(fields.Dia || fields.dia || "").trim();
    const opcao = String(fields.Opcao || fields["Opção"] || fields.opcao || "").trim();
    const nome = String(fields.Colaborador_nome || fields.colaborador_nome || fields.Nome || "").trim();
    const colabId = String(fields.Colaborador_id || fields.colaborador_id || "").trim();

    if (!semana || !dia || !opcao || (!nome && !colabId)) {
      throw new Error("Pedido inválido bloqueado: faltam Semana_id, Dia, Opcao ou Colaborador.");
    }
  },

  async createItem(listName, fields) {
    if (String(listName || "").toLowerCase() === "pedidos") {
      this._validarPedidoFields(fields || {});
    }

    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    return this.graph("POST", `/sites/${siteId}/lists/${listId}/items`, {
      fields: this._cleanFields(fields)
    });
  },

  async updateItem(listName, itemId, fields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    return this.graph("PATCH",
      `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
      this._cleanFields(fields)
    );
  },

  async deleteItem(listName, itemId) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    return this.graph("DELETE", `/sites/${siteId}/lists/${listId}/items/${itemId}`);
  },

  // ============================================================
  // COLABORADORES
  // ============================================================
  async getColaboradores() {
    const items = await this.getItems("Colaboradores");
    return items.filter(i => this.isTrue(this.pick(i, "Ativo")));
  },

  async getTodosColaboradores() {
    return this.getItems("Colaboradores");
  },

  async createColaborador(dados) {
    return this.createItem("Colaboradores", {
      Title:        dados.nome         || dados.Nome         || "",
      Nome:         dados.nome         || dados.Nome         || "",
      Departamento: dados.departamento || dados.Departamento || "",
      Email:        dados.email        || dados.Email        || "",
      Ativo:        true,
      tipo:         dados.tipo         || "Colaborador",
      Centro_Custo: dados.centroCusto  || dados.Centro_Custo || ""
    });
  },

  async updateColaborador(id, dados) {
    const fields = {};
    if (dados.nome        !== undefined) { fields.Title = dados.nome; fields.Nome = dados.nome; }
    if (dados.Nome        !== undefined) { fields.Title = dados.Nome; fields.Nome = dados.Nome; }
    if (dados.departamento !== undefined) fields.Departamento = dados.departamento;
    if (dados.Departamento !== undefined) fields.Departamento = dados.Departamento;
    if (dados.email        !== undefined) fields.Email = dados.email;
    if (dados.Email        !== undefined) fields.Email = dados.Email;
    if (dados.ativo        !== undefined) fields.Ativo = dados.ativo;
    if (dados.Ativo        !== undefined) fields.Ativo = dados.Ativo;
    if (dados.tipo         !== undefined) fields.tipo  = dados.tipo;
    if (dados.centroCusto  !== undefined) fields.Centro_Custo = dados.centroCusto;
    if (dados.Centro_Custo !== undefined) fields.Centro_Custo = dados.Centro_Custo;
    return this.updateItem("Colaboradores", id, fields);
  },

  async desativarColaborador(id) {
    return this.updateItem("Colaboradores", id, { Ativo: false });
  },

  async deleteColaborador(id) {
    return this.deleteItem("Colaboradores", id);
  },

  // ============================================================
  // CARDÁPIO
  // ============================================================
  async getCardapio(semanaId) {
    const items = await this.getItems("Cardapio");
    return items.filter(i => this.pick(i, "Semana_id") === semanaId);
  },

  async saveCardapio(semanaId, dia, opcao, nomePrato, detalhes = "") {
    const existentes = await this.getCardapio(semanaId);
    const norm = v => String(v || "").toLowerCase().trim();
    const existing = existentes.find(i =>
      norm(this.pick(i, "Dia"))  === norm(dia) &&
      norm(this.pick(i, "Opcao")) === norm(opcao)
    );

    const fields = {
      Title:     `${semanaId}-${dia}-${opcao}`,
      Semana_id: semanaId,
      Dia:       dia,
      Opcao:     opcao,
      Nome_Prato: nomePrato,
      Detalhes:   detalhes
    };

    if (existing) return this.updateItem("Cardapio", existing.id, fields);
    return this.createItem("Cardapio", fields);
  },

  async clearCardapio(semanaId) {
    const items = await this.getCardapio(semanaId);
    for (const item of items) await this.deleteItem("Cardapio", item.id);
  },

  // ============================================================
  // PEDIDOS
  // ============================================================
  async getPedidos(semanaId) {
    const items = await this.getItems("Pedidos");
    return items.filter(i => this.pick(i, "Semana_id") === semanaId);
  },

  async getPedidoColaborador(semanaId, colaboradorId) {
    const items = await this.getPedidos(semanaId);
    return items.filter(i =>
      String(this.pick(i, "Colaborador_id")) === String(colaboradorId)
    );
  },

  async savePedido(semanaId, colaboradorId, colaboradorNome, dia, opcao, nomePrato, extras = {}) {
    const fields = {
      Title:            `${semanaId}-${colaboradorId}-${dia}`,
      Semana_id:        semanaId,
      Colaborador_id:   String(colaboradorId),
      Colaborador_nome: colaboradorNome,
      Dia:              dia,
      Opcao:            opcao,
      Nome_Prato:       nomePrato || "",
      Confirmado:       extras.confirmado   ?? extras.Confirmado   ?? false,
      Data_Hora:        extras.dataHora     ?? extras.Data_Hora    ?? new Date().toISOString(),
      Centro_Custo:     extras.centroCusto  ?? extras.Centro_Custo ?? "",
      Status:           extras.status       ?? extras.Status       ?? "Confirmado",
      Observacao:       extras.observacao   ?? extras.Observacao   ?? "",
      Origem:           extras.origem       ?? extras.Origem       ?? "Refeitório",
      Alterado_Por:     extras.alteradoPor  ?? extras.Alterado_Por ?? this.getUserName()
    };
    this._validarPedidoFields(fields);
    return this.createItem("Pedidos", fields);
  },

  async updatePedido(id, dados) {
    const map = {
      Semana_id:        ["Semana_id",        "semanaId"],
      Colaborador_id:   ["Colaborador_id",   "colaboradorId"],
      Colaborador_nome: ["Colaborador_nome", "colaboradorNome"],
      Dia:              ["Dia",              "dia"],
      Opcao:            ["Opcao",            "opcao"],
      Nome_Prato:       ["Nome_Prato",       "nomePrato"],
      Confirmado:       ["Confirmado",       "confirmado"],
      Data_Hora:        ["Data_Hora",        "dataHora"],
      Centro_Custo:     ["Centro_Custo",     "centroCusto"],
      Status:           ["Status",           "status"],
      Observacao:       ["Observacao",       "observacao"],
      Origem:           ["Origem",           "origem"],
      Alterado_Por:     ["Alterado_Por",     "alteradoPor"]
    };

    const fields = {};
    for (const [col, aliases] of Object.entries(map)) {
      for (const alias of aliases) {
        if (dados[alias] !== undefined) {
          fields[col] = col === "Colaborador_id" ? String(dados[alias]) : dados[alias];
          break;
        }
      }
    }
    return this.updateItem("Pedidos", id, fields);
  },

  async deletePedido(id) {
    return this.deleteItem("Pedidos", id);
  },

  async confirmarPedidos(semanaId, colaboradorId) {
    const pedidos = await this.getPedidoColaborador(semanaId, colaboradorId);
    for (const p of pedidos) {
      await this.updateItem("Pedidos", p.id, {
        Confirmado:   true,
        Status:       this.pick(p, "Status") || "Confirmado",
        Alterado_Por: this.getUserName()
      });
    }
  },

  // ============================================================
  // EXTRAS
  // ============================================================
  async getExtras(semanaId, dia = null) {
    const items = await this.getItems("Extras");
    return items.filter(i =>
      this.pick(i, "Semana_id") === semanaId &&
      (!dia || this.pick(i, "Dia") === dia)
    );
  },

  async addExtra(semanaId, dia, nome, tipo, opcao, observacao, adicionadoPor) {
    return this.createItem("Extras", {
      Title:         `${semanaId}-${dia}-${nome}`,
      Semana_id:     semanaId,
      Dia:           dia,
      Nome:          nome,
      tipo:          tipo,
      Opcao:         opcao || "principal",
      Observacao:    observacao || "",
      Adicionado_Por: adicionadoPor || this.getUserName()
    });
  },

  async updateExtra(id, dados) {
    const fields = {};
    if (dados.nome       !== undefined) fields.Nome  = dados.nome;
    if (dados.tipo       !== undefined) fields.tipo  = dados.tipo;
    if (dados.opcao      !== undefined) fields.Opcao = dados.opcao;
    if (dados.observacao !== undefined) fields.Observacao = dados.observacao;
    if (dados.Status     !== undefined) fields.Status = dados.Status;
    if (dados.status     !== undefined) fields.Status = dados.status;
    return this.updateItem("Extras", id, fields);
  },

  async deleteExtra(id) {
    return this.deleteItem("Extras", id);
  },

  // ============================================================
  // CONFIGURAÇÕES
  // ============================================================
  async getConfig(chave) {
    const items = await this.getItems("Configurações");
    const item  = items.find(i =>
      this.pick(i, "Chave") === chave || this.pick(i, "Title") === chave
    );
    return item ? this.pick(item, "Valor") : null;
  },

  async setConfig(chave, valor) {
    const items   = await this.getItems("Configurações");
    const existing = items.find(i =>
      this.pick(i, "Chave") === chave || this.pick(i, "Title") === chave
    );
    if (existing) return this.updateItem("Configurações", existing.id, { Valor: valor });
    return this.createItem("Configurações", { Title: chave, Chave: chave, Valor: valor });
  },

  async isCardapioLiberado() {
    for (const chave of ["cardapio_liberado", "marcacao_liberada", "pedidos_liberados"]) {
      const v = await this.getConfig(chave);
      if (this.isTrue(v)) return true;
    }
    return false;
  },

  async setMarcacaoLiberada(liberado) {
    const v = liberado ? "sim" : "nao";
    await this.setConfig("cardapio_liberado", v);
    await this.setConfig("marcacao_liberada", v);
    await this.setConfig("pedidos_liberados", v);
    return true;
  },

  async getPrazoMarcacao()      { return this.getConfig("prazo_limite"); },
  async setPrazoMarcacao(valor) { return this.setConfig("prazo_limite", valor); },

  // ============================================================
  // VALORES DE REFEIÇÃO
  // Detecção dinâmica de colunas para suportar renomeações
  // ============================================================
  async _resolveColunasValores() {
    const siteId = await this.getSiteId();
    const listId = await this.getListId("Valores de Refeição");
    const data   = await this.graph("GET",
      `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`
    );
    const cols   = data?.value || [];
    const find   = (...candidates) =>
      cols.find(c => candidates.includes(c.name) || candidates.includes(c.displayName))?.name || null;

    return {
      titulo:   find("Title", "Título", "titulo"),
      inicio:   find("Data_Inicio", "DataInicio", "Data Inicio"),
      fim:      find("Data_Fim",    "DataFim",    "Data Fim"),
      vascon:   find("Valor_Vascon", "ValorVascon"),
      desconto: find("Valor_Desconto_Funcionário", "Valor_Desconto_Funcionario"),
      obs:      find("Observacao", "Observação"),
      ativo:    find("Ativo")
    };
  },

  async getValoresRefeicao() {
    return this.getItems("Valores de Refeição");
  },

  async createValorRefeicao(dados) {
    const cols   = await this._resolveColunasValores();
    const fields = {};
    if (cols.titulo)   fields[cols.titulo]   = dados.title || dados.titulo || dados.Title || "";
    if (cols.inicio)   fields[cols.inicio]   = dados.dataInicio || dados.Data_Inicio;
    if (cols.fim)      fields[cols.fim]      = dados.dataFim    || dados.Data_Fim;
    if (cols.vascon)   fields[cols.vascon]   = Number(dados.valorVascon   ?? dados.Valor_Vascon   ?? 0);
    if (cols.desconto) fields[cols.desconto] = Number(dados.valorDesconto ?? dados.Valor_Desconto_Funcionario ?? 0);
    if (cols.obs)      fields[cols.obs]      = dados.observacao || dados.Observacao || "";
    if (cols.ativo)    fields[cols.ativo]    = dados.ativo !== false && dados.Ativo !== false;
    return this.createItem("Valores de Refeição", fields);
  },

  async updateValorRefeicao(id, dados) {
    const cols   = await this._resolveColunasValores();
    const fields = {};
    if ((dados.title     || dados.titulo)     && cols.titulo)   fields[cols.titulo]   = dados.title || dados.titulo;
    if ((dados.dataInicio !== undefined)       && cols.inicio)   fields[cols.inicio]   = dados.dataInicio;
    if ((dados.Data_Inicio !== undefined)      && cols.inicio)   fields[cols.inicio]   = dados.Data_Inicio;
    if ((dados.dataFim    !== undefined)       && cols.fim)      fields[cols.fim]      = dados.dataFim;
    if ((dados.Data_Fim   !== undefined)       && cols.fim)      fields[cols.fim]      = dados.Data_Fim;
    if ((dados.valorVascon  !== undefined)     && cols.vascon)   fields[cols.vascon]   = Number(dados.valorVascon);
    if ((dados.valorDesconto !== undefined)    && cols.desconto) fields[cols.desconto] = Number(dados.valorDesconto);
    if ((dados.observacao !== undefined)       && cols.obs)      fields[cols.obs]      = dados.observacao;
    if ((dados.ativo      !== undefined)       && cols.ativo)    fields[cols.ativo]    = dados.ativo;
    return this.updateItem("Valores de Refeição", id, fields);
  },

  // ============================================================
  // AUSÊNCIAS
  // Lista: Ausencias do Refeitorio
  // Colunas: Title, Colaborador_id, Colaborador_nome,
  //          Data_Inicio, Data_Fim, Motivo, Observacao, Ativo, Criado_Por
  // Motivos válidos: ferias | atestado | falta | licenca | afastamento |
  //                  nao_vai_almocar | homy_office | banco_horas | outro
  // ============================================================
  async getAusencias(apenasAtivas = true) {
    const items = await this.getItems("Ausencias do Refeitorio");
    return apenasAtivas ? items.filter(i => this.isTrue(this.pick(i, "Ativo"))) : items;
  },

  async getAusenciasColaborador(colaboradorId, dataRef = null) {
    const items = await this.getAusencias(true);
    return items.filter(i => {
      if (String(this.pick(i, "Colaborador_id")) !== String(colaboradorId)) return false;
      if (!dataRef) return true;
      const d   = new Date(dataRef);
      const ini = this.pick(i, "Data_Inicio") ? new Date(this.pick(i, "Data_Inicio")) : null;
      const fim = this.pick(i, "Data_Fim")    ? new Date(this.pick(i, "Data_Fim"))    : null;
      return ini && fim && d >= ini && d <= fim;
    });
  },

  async colaboradorEstaAusente(colaboradorId, dataRef = new Date()) {
    const aus = await this.getAusenciasColaborador(colaboradorId, dataRef);
    return aus.length > 0 ? aus[0] : null;
  },

  async createAusencia(dados) {
    const nome   = dados.colaboradorNome || dados.Colaborador_nome || "";
    const motivo = dados.motivo || dados.Motivo || "nao_vai_almocar";
    return this.createItem("Ausencias do Refeitorio", {
      Title:            dados.title || `${nome} - ${motivo}`,
      Colaborador_id:   String(dados.colaboradorId || dados.Colaborador_id || ""),
      Colaborador_nome: nome,
      Centro_Custo:     dados.centroCusto || dados.Centro_Custo || "",
      Data_Inicio:      dados.dataInicio  || dados.Data_Inicio,
      Data_Fim:         dados.dataFim     || dados.Data_Fim,
      Motivo:           motivo,
      Observacao:       dados.observacao  || dados.Observacao || "",
      Ativo:            dados.ativo       ?? dados.Ativo      ?? true,
      Criado_Por:       dados.criadoPor   || dados.Criado_Por || this.getUserName()
    });
  },

  async updateAusencia(id, dados) {
    const fields = {};
    if (dados.ativo       !== undefined) fields.Ativo     = dados.ativo;
    if (dados.Ativo       !== undefined) fields.Ativo     = dados.Ativo;
    if (dados.motivo      !== undefined) fields.Motivo    = dados.motivo;
    if (dados.observacao  !== undefined) fields.Observacao = dados.observacao;
    if (dados.dataInicio  !== undefined) fields.Data_Inicio = dados.dataInicio;
    if (dados.dataFim     !== undefined) fields.Data_Fim    = dados.dataFim;
    return this.updateItem("Ausencias do Refeitorio", id, fields);
  },

  async deleteAusencia(id) {
    return this.deleteItem("Ausencias do Refeitorio", id);
  },

  // ============================================================
  // CHECK-IN
  // Lista: CheckIn
  // ============================================================
  async getCheckIn(semanaId, dia) {
    const items = await this.getItems("CheckIn");
    return items.filter(i =>
      this.pick(i, "Semana_id") === semanaId &&
      this.pick(i, "Dia")       === dia
    );
  },

  async registrarCheckIn(semanaId, colaboradorId, colaboradorNome, dia, confirmadoPor) {
    const existing = await this.getCheckIn(semanaId, dia);
    const found    = existing.find(i =>
      String(this.pick(i, "Colaborador_id")) === String(colaboradorId)
    );

    const fields = {
      Retirou:            true,
      Data_Hora_Retirada: new Date().toISOString(),
      Confirmado_Por:     confirmadoPor
    };

    if (found) return this.updateItem("CheckIn", found.id, fields);

    return this.createItem("CheckIn", {
      Title:            `${semanaId}-${colaboradorId}-${dia}`,
      Semana_id:        semanaId,
      Colaborador_id:   String(colaboradorId),
      Colaborador_nome: colaboradorNome,
      Dia:              dia,
      ...fields
    });
  },

  async saveCheckIn(semanaId, colaboradorId, colaboradorNome, dia, confirmadoPor) {
    return this.registrarCheckIn(semanaId, colaboradorId, colaboradorNome, dia, confirmadoPor);
  }

};
