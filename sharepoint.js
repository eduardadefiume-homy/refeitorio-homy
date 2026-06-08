// ============================================================
// sharepoint.js — Refeitório Homy · Microsoft Graph API
// Versão: 2026-06-08
//
// App registration: Refeitorio Homy
// clientId: aa37acf9-f3bd-4d1e-968a-fde57f79094c
//
// Decisão técnica:
// - Login sempre por redirect
// - Sem loginPopup
// - Sem acquireTokenPopup
// - SharePoint continua sendo a fonte oficial dos dados
// - no PATCH/POST, sem funções duplicadas.
// ============================================================

const SP = window.SP = {

  // ============================================================
  // CREDENCIAIS
  // ============================================================
  clientId: "aa37acf9-f3bd-4d1e-968a-fde57f79094c",
  tenantId: "a2850abc-334a-4805-b6b2-420b4aef68a9",
  scopes: ["Sites.ReadWrite.All", "User.Read"],

  siteUrl: "homyquimica.sharepoint.com",

  _sitePathCandidates: [
    "/sites/Refeitrio-Homy",
    "/sites/Refeitorio-Homy",
    "/sites/Refeitório-Homy"
  ],

  // ============================================================
  // ESTADO INTERNO
  // ============================================================
  _msalInstance: null,
  _account: null,
  _siteId: null,
  _sitePath: null,
  _listIds: {},
  _columnsCache: {},
  _colunasValores: null,

  _msalReadyPromise: null,
  _redirectHandled: false,
  _loginEmAndamento: false,

  // ============================================================
  // CAMPOS READ ONLY
  // ============================================================
  _readOnlyFields: new Set([
    "ComplianceAssetId",
    "id",
    "ID",
    "Created",
    "Modified",
    "Author",
    "Editor",
    "AuthorId",
    "EditorId",
    "FileSystemObjectType",
    "ContentTypeId",
    "_UIVersionString",
    "_UIVersion",
    "Attachments",
    "CheckoutUser",
    "GUID",
    "UniqueId",
    "owshiddenversion",
    "AppAuthor",
    "AppEditor",
    "DocIcon",
    "LinkTitle",
    "LinkTitleNoMenu",
    "Edit",
    "SelectTitle",
    "SelectFilename",
    "ItemChildCount",
    "FolderChildCount"
  ]),

  _cleanFields(fields) {
    if (!fields || typeof fields !== "object") return fields;

    return Object.fromEntries(
      Object.entries(fields).filter(([k]) =>
        !this._readOnlyFields.has(k) &&
        !k.startsWith("@") &&
        !k.startsWith("OData_")
      )
    );
  },

  // ============================================================
  // UTILITÁRIOS
  // ============================================================
  pick(obj, ...keys) {
    for (const k of keys) {
      if (obj && obj[k] !== undefined && obj[k] !== null) return obj[k];
    }
    return null;
  },

  norm(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim();
  },

  isTrue(value) {
    if (value === true || value === 1) return true;
    if (value === false || value === 0) return false;

    const v = String(value ?? "").trim().toLowerCase();
    return ["sim", "true", "yes", "1", "ativo"].includes(v);
  },

  getUserName() {
    return this._account?.name || this._account?.username || "Usuário Homy";
  },

  getUserEmail() {
    return this._account?.username || "";
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
    const jan4 = new Date(year, 0, 4);
    const start = new Date(jan4);

    start.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1 + (week - 1) * 7);

    return Array.from({ length: 5 }, (_, i) => {
      const d = new Date(start);
      d.setDate(start.getDate() + i);
      return d;
    });
  },

  getDataRefBySemanaDia(semanaId, dia) {
    const ordem = {
      segunda: 0,
      terca: 1,
      terça: 1,
      quarta: 2,
      quinta: 3,
      sexta: 4
    };

    const datas = this.getWeekDates(semanaId);
    const idx = ordem[this.norm(dia)];

    if (idx === undefined || !datas[idx]) {
      return new Date().toISOString().slice(0, 10);
    }

    return datas[idx].toISOString().slice(0, 10);
  },

  isExtraPedido(p) {
    const origem = this.norm(this.pick(p, "Origem", "tipo", "Tipo") || "");
    const nome = this.norm(this.pick(p, "Colaborador_nome", "Title", "Nome") || "");

    return origem.includes("extra") ||
      origem.includes("investigador") ||
      origem.includes("guarda") ||
      nome.includes("refeicao extra");
  },

  _toSharePointNumber(value) {
    if (value === null || value === undefined || String(value).trim() === "") return null;

    let s = String(value)
      .replace(/R\$/gi, "")
      .replace(/\s/g, "")
      .trim();

    if (s.includes(",")) {
      s = s.replace(/\./g, "").replace(",", ".");
    }

    s = s.replace(/[^0-9.-]/g, "");

    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  },

  _toSharePointBool(value) {
    return this.isTrue(value);
  },

  // ============================================================
  // AUTENTICAÇÃO MSAL
  // ============================================================
  _getRedirectUri() {
    return window.location.origin + window.location.pathname;
  },

  _getAllStorageKeys() {
    return [
      ...Object.keys(sessionStorage),
      ...Object.keys(localStorage)
    ];
  },

  _hasInteractionInProgress() {
    return this._getAllStorageKeys().some(k => {
      const key = k.toLowerCase();
      return key.includes("interaction.status") ||
        key.includes("msal.interaction.status");
    });
  },

  _clearMsalInteractionOnly() {
    for (const storage of [sessionStorage, localStorage]) {
      Object.keys(storage).forEach(k => {
        const key = k.toLowerCase();

        if (
          key.includes("interaction.status") ||
          key.includes("msal.interaction.status") ||
          key.includes("request.state") ||
          key.includes("nonce.id_token") ||
          key.includes("authority") ||
          key.includes("urlhash")
        ) {
          storage.removeItem(k);
        }
      });
    }
  },

  async _wait(ms = 200) {
    return new Promise(resolve => setTimeout(resolve, ms));
  },

  async init() {
    if (this._msalReadyPromise) return this._msalReadyPromise;

    this._msalReadyPromise = (async () => {
      if (!window.msal) {
        throw new Error("MSAL não carregou. Verifique se msal-browser.min.js está antes de sharepoint.js.");
      }

      if (!this._msalInstance) {
        this._msalInstance = new msal.PublicClientApplication({
          auth: {
            clientId: this.clientId,
            authority: `https://login.microsoftonline.com/${this.tenantId}`,
            redirectUri: this._getRedirectUri(),
            navigateToLoginRequestUrl: false
          },
          cache: {
            cacheLocation: "sessionStorage",
            storeAuthStateInCookie: true
          },
          system: {
            allowNativeBroker: false
          }
        });

        await this._msalInstance.initialize();
      }

      if (!this._redirectHandled) {
        this._redirectHandled = true;

        try {
          const result = await this._msalInstance.handleRedirectPromise();

          if (result?.account) {
            this._account = result.account;
            this._msalInstance.setActiveAccount(result.account);
            this._loginEmAndamento = false;
            this._clearMsalInteractionOnly();
            return true;
          }
        } catch (e) {
          console.warn("[SP] handleRedirectPromise falhou:", e);

          const msg = String(e?.errorCode || e?.message || e || "");

          if (msg.includes("interaction_in_progress")) {
            this._clearMsalInteractionOnly();
          }
        }
      }

      const active = this._msalInstance.getActiveAccount();
      if (active) {
        this._account = active;
        return true;
      }

      const accounts = this._msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        this._account = accounts[0];
        this._msalInstance.setActiveAccount(this._account);
        return true;
      }

      return false;
    })();

    return this._msalReadyPromise;
  },

  async login() {
    await this.init();

    if (this._account) return true;

    if (this._loginEmAndamento || this._hasInteractionInProgress()) {
      this._clearMsalInteractionOnly();
      await this._wait(500);
    }

    this._loginEmAndamento = true;

    const request = {
      scopes: this.scopes,
      prompt: "select_account"
    };

    try {
      await this._msalInstance.loginRedirect(request);
      return false;

    } catch (e) {
      this._loginEmAndamento = false;

      const msg = String(e?.errorCode || e?.message || e || "");

      if (msg.includes("interaction_in_progress")) {
        this._clearMsalInteractionOnly();
        throw new Error("Login Microsoft já estava em andamento. Recarregue a página e tente novamente.");
      }

      throw e;
    }
  },

  async ensureLogin() {
    const logado = await this.init();

    if (logado && this._account) return true;

    await this.login();
    return false;
  },

  async getToken() {
    await this.init();

    if (!this._account) return null;

    try {
      const result = await this._msalInstance.acquireTokenSilent({
        scopes: this.scopes,
        account: this._account
      });

      return result.accessToken;

    } catch (e) {
      const msg = String(e?.errorCode || e?.message || e || "");
      console.warn("[SP] acquireTokenSilent falhou:", e);

      if (msg.includes("interaction_in_progress")) {
        this._clearMsalInteractionOnly();
        return null;
      }

      return null;
    }
  },

  async getTokenInterativo() {
    await this.init();

    if (!this._account) {
      await this.login();
      return null;
    }

    try {
      const silent = await this._msalInstance.acquireTokenSilent({
        scopes: this.scopes,
        account: this._account
      });

      return silent.accessToken;

    } catch (e) {
      const msg = String(e?.errorCode || e?.message || e || "");

      if (msg.includes("interaction_in_progress")) {
        this._clearMsalInteractionOnly();
        return null;
      }

      await this._msalInstance.acquireTokenRedirect({
        scopes: this.scopes,
        account: this._account
      });

      return null;
    }
  },

  async logout() {
    await this.init();

    const account = this._account;

    this._account = null;
    this._loginEmAndamento = false;
    this._msalReadyPromise = null;
    this._redirectHandled = false;

    sessionStorage.clear();

    if (account) {
      await this._msalInstance.logoutRedirect({
        account,
        postLogoutRedirectUri: this._getRedirectUri()
      });
    }
  },

  resetAuthLocal() {
    this._account = null;
    this._loginEmAndamento = false;
    this._msalReadyPromise = null;
    this._redirectHandled = false;

    sessionStorage.clear();

    for (const storage of [localStorage]) {
      Object.keys(storage).forEach(k => {
        if (k.toLowerCase().includes("msal")) {
          storage.removeItem(k);
        }
      });
    }
  },

  // ============================================================
  // GRAPH
  // ============================================================
  async graph(method, endpoint, body = null, options = {}) {
    const token = options.interativo
      ? await this.getTokenInterativo()
      : await this.getToken();

    if (!token) {
      throw new Error("Usuário não autenticado no Microsoft. Entre novamente antes de acessar o SharePoint.");
    }

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
        "Content-Type": "application/json"
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

  async graphInterativo(method, endpoint, body = null) {
    return this.graph(method, endpoint, body, { interativo: true });
  },

  // ============================================================
  // SITE E LISTAS
  // ============================================================
  async getSiteId() {
    if (this._siteId) return this._siteId;

    const paths = this._sitePath
      ? [this._sitePath, ...this._sitePathCandidates]
      : this._sitePathCandidates;

    const unique = [...new Set(paths)];
    let lastErr = null;

    for (const path of unique) {
      try {
        const data = await this.graph("GET", `/sites/${this.siteUrl}:${path}`);

        if (data?.id) {
          this._siteId = data.id;
          this._sitePath = path;
          console.log(`[SP] Site encontrado em: ${path}`);
          return this._siteId;
        }

      } catch (e) {
        lastErr = e;
      }
    }

    throw lastErr || new Error(`Site SharePoint não encontrado. Testados: ${unique.join(", ")}`);
  },

  async getListId(listName) {
    if (this._listIds[listName]) return this._listIds[listName];

    const siteId = await this.getSiteId();

    const data = await this.graph(
      "GET",
      `/sites/${siteId}/lists?$filter=displayName eq '${listName}'&$select=id,displayName`
    );

    if (!data?.value?.length) {
      throw new Error(`Lista não encontrada: "${listName}"`);
    }

    this._listIds[listName] = data.value[0].id;
    return this._listIds[listName];
  },

  async getItems(listName) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);

    let endpoint = `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=999`;
    const items = [];

    while (endpoint) {
      const data = await this.graph("GET", endpoint);
      items.push(...(data?.value || []));

      endpoint = data?.["@odata.nextLink"]
        ? data["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "")
        : null;
    }

    return items.map(i => ({ id: i.id, ...i.fields }));
  },

  async createItem(listName, fields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);

    return this.graphInterativo("POST", `/sites/${siteId}/lists/${listId}/items`, {
      fields: this._cleanFields(fields)
    });
  },

  async updateItem(listName, itemId, fields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);

    return this.graphInterativo(
      "PATCH",
      `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
      this._cleanFields(fields)
    );
  },

  async deleteItem(listName, itemId) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);

    return this.graphInterativo("DELETE", `/sites/${siteId}/lists/${listId}/items/${itemId}`);
  },

  // ============================================================
  // COLUNAS SHAREPOINT
  // ============================================================
  async getListColumns(listName) {
    if (this._columnsCache[listName]) return this._columnsCache[listName];

    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);

    const data = await this.graph(
      "GET",
      `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`
    );

    const cols = data?.value || [];
    this._columnsCache[listName] = cols;

    return cols;
  },

  _normColumn(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[_\s\-]/g, "")
      .replace(/[^a-zA-Z0-9]/g, "")
      .toLowerCase()
      .trim();
  },

  async findColumnName(listName, candidates) {
    const cols = await this.getListColumns(listName);
    const wanted = candidates.map(c => this._normColumn(c));

    for (const col of cols) {
      const internalName = this._normColumn(col.name);
      const displayName = this._normColumn(col.displayName);

      if (wanted.includes(internalName) || wanted.includes(displayName)) {
        return col.name;
      }
    }

    return null;
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
      Title: dados.nome || dados.Nome || "",
      Nome: dados.nome || dados.Nome || "",
      Departamento: dados.departamento || dados.Departamento || "",
      Email: dados.email || dados.Email || "",
      Ativo: true,
      tipo: dados.tipo || "Colaborador",
      Centro_Custo: dados.centroCusto || dados.Centro_Custo || ""
    });
  },

  async updateColaborador(id, dados) {
    const fields = {};

    if (dados.nome !== undefined) {
      fields.Title = dados.nome;
      fields.Nome = dados.nome;
    }

    if (dados.Nome !== undefined) {
      fields.Title = dados.Nome;
      fields.Nome = dados.Nome;
    }

    if (dados.departamento !== undefined) fields.Departamento = dados.departamento;
    if (dados.Departamento !== undefined) fields.Departamento = dados.Departamento;
    if (dados.email !== undefined) fields.Email = dados.email;
    if (dados.Email !== undefined) fields.Email = dados.Email;
    if (dados.ativo !== undefined) fields.Ativo = dados.ativo;
    if (dados.Ativo !== undefined) fields.Ativo = dados.Ativo;
    if (dados.tipo !== undefined) fields.tipo = dados.tipo;
    if (dados.centroCusto !== undefined) fields.Centro_Custo = dados.centroCusto;
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

    const existing = existentes.find(i =>
      this.norm(this.pick(i, "Dia")) === this.norm(dia) &&
      this.norm(this.pick(i, "Opcao")) === this.norm(opcao)
    );

    const fields = {
      Title: `${semanaId}-${dia}-${opcao}`,
      Semana_id: semanaId,
      Dia: dia,
      Opcao: opcao,
      Nome_Prato: nomePrato,
      Detalhes: detalhes
    };

    if (existing) {
      return this.updateItem("Cardapio", existing.id, fields);
    }

    return this.createItem("Cardapio", fields);
  },

  async clearCardapio(semanaId) {
    const items = await this.getCardapio(semanaId);

    for (const item of items) {
      await this.deleteItem("Cardapio", item.id);
    }
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
    return this.createItem("Pedidos", {
      Title: `${semanaId}-${colaboradorId}-${dia}`,
      Semana_id: semanaId,
      Colaborador_id: String(colaboradorId),
      Colaborador_nome: colaboradorNome,
      Dia: dia,
      Opcao: opcao,
      Nome_Prato: nomePrato || "",
      Confirmado: extras.confirmado ?? extras.Confirmado ?? false,
      Data_Hora: extras.dataHora ?? extras.Data_Hora ?? new Date().toISOString(),
      Centro_Custo: extras.centroCusto ?? extras.Centro_Custo ?? "",
      Status: extras.status ?? extras.Status ?? "Confirmado",
      Observacao: extras.observacao ?? extras.Observacao ?? "",
      Origem: extras.origem ?? extras.Origem ?? "Refeitório",
      Alterado_Por: extras.alteradoPor ?? extras.Alterado_Por ?? this.getUserName()
    });
  },

  async updatePedido(id, dados) {
    const map = {
      Semana_id: ["Semana_id", "semanaId"],
      Colaborador_id: ["Colaborador_id", "colaboradorId"],
      Colaborador_nome: ["Colaborador_nome", "colaboradorNome"],
      Dia: ["Dia", "dia"],
      Opcao: ["Opcao", "opcao"],
      Nome_Prato: ["Nome_Prato", "nomePrato"],
      Confirmado: ["Confirmado", "confirmado"],
      Data_Hora: ["Data_Hora", "dataHora"],
      Centro_Custo: ["Centro_Custo", "centroCusto"],
      Status: ["Status", "status"],
      Observacao: ["Observacao", "observacao"],
      Origem: ["Origem", "origem"],
      Alterado_Por: ["Alterado_Por", "alteradoPor"]
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
        Confirmado: true,
        Status: this.pick(p, "Status") || "Confirmado",
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
      Title: `${semanaId}-${dia}-${nome}`,
      Semana_id: semanaId,
      Dia: dia,
      Nome: nome,
      tipo: tipo,
      Opcao: opcao,
      Observacao: observacao || "",
      Adicionado_Por: adicionadoPor || this.getUserName()
    });
  },

  async addExtraPedido(semanaId, dia, nome, tipo, opcao, observacao, adicionadoPor) {
    const user = adicionadoPor || this.getUserName();

    await this.addExtra(semanaId, dia, nome, tipo, opcao, observacao, user);

    await this.savePedido(
      semanaId,
      `EXTRA-${Date.now()}`,
      nome,
      dia,
      opcao,
      "",
      {
        confirmado: true,
        status: "Confirmado",
        origem: tipo,
        observacao: observacao || nome,
        alteradoPor: user
      }
    );
  },

  async removeExtra(id) {
    return this.deleteItem("Extras", id);
  },

  async deleteExtra(id) {
    return this.deleteItem("Extras", id);
  },

  async deleteExtraComPedido(extra) {
    if (extra?.id) await this.deleteItem("Extras", extra.id);

    const semanaId = this.pick(extra, "Semana_id");
    const dia = this.pick(extra, "Dia");
    const nome = this.pick(extra, "Nome", "Title");

    if (semanaId && dia && nome) {
      const pedidos = await this.getPedidos(semanaId);

      const vinculado = pedidos.find(p =>
        this.norm(this.pick(p, "Dia")) === this.norm(dia) &&
        this.norm(this.pick(p, "Colaborador_nome", "Title")) === this.norm(nome)
      );

      if (vinculado) await this.deleteItem("Pedidos", vinculado.id);
    }
  },

  // ============================================================
  // CONFIGURAÇÕES
  // ============================================================
  async getConfig(chave) {
    const items = await this.getItems("Configurações");

    const item = items.find(i =>
      this.pick(i, "Chave") === chave ||
      this.pick(i, "Title") === chave
    );

    return item ? this.pick(item, "Valor") : null;
  },

  async setConfig(chave, valor) {
    const items = await this.getItems("Configurações");

    const existing = items.find(i =>
      this.pick(i, "Chave") === chave ||
      this.pick(i, "Title") === chave
    );

    if (existing) {
      return this.updateItem("Configurações", existing.id, {
        Valor: String(valor)
      });
    }

    return this.createItem("Configurações", {
      Title: chave,
      Chave: chave,
      Valor: String(valor)
    });
  },

  async isCardapioLiberado() {
    for (const chave of ["cardapio_liberado", "marcacao_liberada", "pedidos_liberados"]) {
      if (this.isTrue(await this.getConfig(chave))) return true;
    }

    return false;
  },

  async setMarcacaoLiberada(liberado) {
    const v = liberado ? "sim" : "nao";

    await this.setConfig("cardapio_liberado", v);
    await this.setConfig("marcacao_liberada", v);
    await this.setConfig("pedidos_liberados", v);
  },

  async setCardapioVisivel(visivel) {
    return this.setConfig("cardapio_visivel", visivel ? "sim" : "nao");
  },

  async getPrazoMarcacao() {
    return this.getConfig("prazo_limite");
  },

  async setPrazoMarcacao(valor) {
    await this.setConfig("prazo_limite", valor);
    await this.setConfig("prazo_marcacao", valor);
  },

  // ============================================================
  // VALORES DE REFEIÇÃO
  // ============================================================
  async _resolveColunasValores() {
    if (this._colunasValores) return this._colunasValores;

    const listName = "Valores de Refeição";

    const colunas = {
      titulo: "Title",

      inicio: await this.findColumnName(listName, [
        "Data_Inicio",
        "Data Inicio",
        "Data início",
        "Inicio",
        "Início"
      ]),

      fim: await this.findColumnName(listName, [
        "Data_Fim",
        "Data Fim",
        "Fim"
      ]),

      vascon: await this.findColumnName(listName, [
        "Valor_Vascon",
        "Valor Vascon",
        "Vascon"
      ]),

      desconto: await this.findColumnName(listName, [
        "Valor_Desconto_Funcionário",
        "Valor_Desconto_Funcionario",
        "Valor Desconto Funcionário",
        "Valor Desconto Funcionario",
        "Desconto Funcionário",
        "Desconto Funcionario"
      ]),

      obs: await this.findColumnName(listName, [
        "Observacao",
        "Observação",
        "Observacoes",
        "Observações",
        "Obs"
      ]),

      ativo: await this.findColumnName(listName, [
        "Ativo",
        "Status"
      ])
    };

    const obrigatorias = ["inicio", "fim", "vascon", "desconto", "ativo"];
    const faltando = obrigatorias.filter(k => !colunas[k]);

    if (faltando.length) {
      const cols = await this.getListColumns(listName);
      console.error("[SP] Colunas encontradas na lista Valores de Refeição:", cols);

      throw new Error(
        "Não consegui localizar as colunas obrigatórias da lista Valores de Refeição: " +
        faltando.join(", ") +
        ". Abra o console para ver as colunas encontradas."
      );
    }

    this._colunasValores = colunas;
    console.log("[SP] Colunas Valores de Refeição resolvidas:", colunas);

    return this._colunasValores;
  },

  async getValoresRefeicao(apenasAtivos = true) {
    const items = await this.getItems("Valores de Refeição");
    const cols = await this._resolveColunasValores();

    const normalizados = items.map(item => ({
      ...item,
      ValorDataInicio: this.pick(item, cols.inicio),
      ValorDataFim: this.pick(item, cols.fim),
      ValorVascon: this.pick(item, cols.vascon),
      ValorDescontoFuncionario: this.pick(item, cols.desconto),
      ValorObservacao: cols.obs ? this.pick(item, cols.obs) : "",
      ValorAtivo: this.pick(item, cols.ativo)
    }));

    if (!apenasAtivos) return normalizados;

    return normalizados.filter(i => this.isTrue(i.ValorAtivo));
  },

  async getValorRefeicaoVigente(dataRef = new Date()) {
    const data = new Date(dataRef);
    const items = await this.getValoresRefeicao(true);

    return items.find(i => {
      const ini = i.ValorDataInicio ? new Date(i.ValorDataInicio) : null;
      const fim = i.ValorDataFim ? new Date(i.ValorDataFim) : null;

      return ini && fim && data >= ini && data <= fim;
    }) || null;
  },

  async createValorRefeicao(dados) {
    const cols = await this._resolveColunasValores();

    if (this._toSharePointBool(dados.ativo)) {
      await this._desativarValoresRefeicaoAtivos();
    }

    const fields = {
      [cols.titulo]: dados.title || dados.titulo || "Valor refeição",
      [cols.inicio]: dados.dataInicio || dados.Data_Inicio || null,
      [cols.fim]: dados.dataFim || dados.Data_Fim || null,
      [cols.vascon]: this._toSharePointNumber(dados.valorVascon ?? dados.Valor_Vascon),
      [cols.desconto]: this._toSharePointNumber(dados.valorDesconto ?? dados.Valor_Desconto_Funcionario),
      [cols.ativo]: this._toSharePointBool(dados.ativo ?? dados.Ativo ?? true)
    };

    if (cols.obs) {
      fields[cols.obs] = dados.observacao || dados.Observacao || "";
    }

    return this.createItem("Valores de Refeição", fields);
  },

  async updateValorRefeicao(id, dados) {
    const cols = await this._resolveColunasValores();

    if (this._toSharePointBool(dados.ativo)) {
      await this._desativarValoresRefeicaoAtivos(id);
    }

    const fields = {
      [cols.titulo]: dados.title || dados.titulo || "Valor refeição"
    };

    if (dados.dataInicio !== undefined || dados.Data_Inicio !== undefined) {
      fields[cols.inicio] = dados.dataInicio || dados.Data_Inicio || null;
    }

    if (dados.dataFim !== undefined || dados.Data_Fim !== undefined) {
      fields[cols.fim] = dados.dataFim || dados.Data_Fim || null;
    }

    if (dados.valorVascon !== undefined || dados.Valor_Vascon !== undefined) {
      fields[cols.vascon] = this._toSharePointNumber(dados.valorVascon ?? dados.Valor_Vascon);
    }

    if (dados.valorDesconto !== undefined || dados.Valor_Desconto_Funcionario !== undefined) {
      fields[cols.desconto] = this._toSharePointNumber(dados.valorDesconto ?? dados.Valor_Desconto_Funcionario);
    }

    if (dados.observacao !== undefined || dados.Observacao !== undefined) {
      if (cols.obs) fields[cols.obs] = dados.observacao || dados.Observacao || "";
    }

    if (dados.ativo !== undefined || dados.Ativo !== undefined) {
      fields[cols.ativo] = this._toSharePointBool(dados.ativo ?? dados.Ativo);
    }

    return this.updateItem("Valores de Refeição", id, fields);
  },

  async deleteValorRefeicao(id) {
    return this.deleteItem("Valores de Refeição", id);
  },

  async _desativarValoresRefeicaoAtivos(excetoId = null) {
    const cols = await this._resolveColunasValores();
    const valores = await this.getValoresRefeicao(false);

    const ativos = valores.filter(v => {
      if (excetoId && String(v.id) === String(excetoId)) return false;
      return this.isTrue(v.ValorAtivo);
    });

    for (const item of ativos) {
      await this.updateItem("Valores de Refeição", item.id, {
        [cols.ativo]: false
      });
    }
  },

  // ============================================================
  // AUSÊNCIAS
  // ============================================================
  async getAusencias(apenasAtivas = true) {
    const items = await this.getItems("Ausencias do Refeitorio");
    return apenasAtivas ? items.filter(i => this.isTrue(this.pick(i, "Ativo"))) : items;
  },

  async getAusenciasRefeitorio(apenasAtivas = true) {
    return this.getAusencias(apenasAtivas);
  },

  async getAusenciasColaborador(colaboradorId, dataRef = null) {
    const items = await this.getAusencias(true);

    return items.filter(i => {
      if (String(this.pick(i, "Colaborador_id")) !== String(colaboradorId)) return false;
      if (!dataRef) return true;

      const d = new Date(dataRef);
      const ini = this.pick(i, "Data_Inicio") ? new Date(this.pick(i, "Data_Inicio")) : null;
      const fim = this.pick(i, "Data_Fim") ? new Date(this.pick(i, "Data_Fim")) : null;

      return ini && fim && d >= ini && d <= fim;
    });
  },

  async colaboradorEstaAusente(colaboradorId, dataRef = new Date()) {
    const aus = await this.getAusenciasColaborador(colaboradorId, dataRef);
    return aus.length > 0 ? aus[0] : null;
  },

  async createAusencia(dados) {
    const nome = dados.colaboradorNome || dados.Colaborador_nome || "";
    const motivo = dados.motivo || dados.Motivo || "Ausência";

    return this.createItem("Ausencias do Refeitorio", {
      Title: dados.title || `${nome} ${motivo}`,
      Colaborador_id: String(dados.colaboradorId || dados.Colaborador_id || ""),
      Colaborador_nome: nome,
      Data_Inicio: dados.dataInicio || dados.Data_Inicio,
      Data_Fim: dados.dataFim || dados.Data_Fim,
      Motivo: motivo,
      Observacao: dados.observacao || dados.Observacao || "",
      Ativo: dados.ativo ?? dados.Ativo ?? true,
      Criado_Por: dados.criadoPor || dados.Criado_Por || this.getUserName()
    });
  },

  async createAusenciaRefeitorio(dados) {
    return this.createAusencia(dados);
  },

  async updateAusenciaRefeitorio(id, dados) {
    const fields = {};

    if (dados.colaboradorId !== undefined) fields.Colaborador_id = String(dados.colaboradorId);
    if (dados.Colaborador_id !== undefined) fields.Colaborador_id = String(dados.Colaborador_id);
    if (dados.colaboradorNome !== undefined) fields.Colaborador_nome = dados.colaboradorNome;
    if (dados.Colaborador_nome !== undefined) fields.Colaborador_nome = dados.Colaborador_nome;
    if (dados.dataInicio !== undefined) fields.Data_Inicio = dados.dataInicio;
    if (dados.Data_Inicio !== undefined) fields.Data_Inicio = dados.Data_Inicio;
    if (dados.dataFim !== undefined) fields.Data_Fim = dados.dataFim;
    if (dados.Data_Fim !== undefined) fields.Data_Fim = dados.Data_Fim;
    if (dados.motivo !== undefined) fields.Motivo = dados.motivo;
    if (dados.Motivo !== undefined) fields.Motivo = dados.Motivo;
    if (dados.observacao !== undefined) fields.Observacao = dados.observacao;
    if (dados.Observacao !== undefined) fields.Observacao = dados.Observacao;
    if (dados.ativo !== undefined) fields.Ativo = dados.ativo;
    if (dados.Ativo !== undefined) fields.Ativo = dados.Ativo;

    return this.updateItem("Ausencias do Refeitorio", id, fields);
  },

  async deleteAusencia(id) {
    return this.deleteItem("Ausencias do Refeitorio", id);
  },

  // ============================================================
  // CHECK IN
  // ============================================================
  async getCheckIn(semanaId, dia) {
    const items = await this.getItems("CheckIn");

    return items.filter(i =>
      this.pick(i, "Semana_id") === semanaId &&
      this.pick(i, "Dia") === dia
    );
  },

  async getCheckIns() {
    return this.getItems("CheckIn");
  },

  async registrarCheckIn(semanaId, colaboradorId, colaboradorNome, dia, confirmadoPor) {
    const existing = await this.getCheckIn(semanaId, dia);

    const found = existing.find(i =>
      String(this.pick(i, "Colaborador_id")) === String(colaboradorId)
    );

    const fields = {
      Retirou: true,
      Data_Hora_Retirada: new Date().toISOString(),
      Confirmado_Por: confirmadoPor
    };

    if (found) {
      return this.updateItem("CheckIn", found.id, fields);
    }

    return this.createItem("CheckIn", {
      Title: `${semanaId}-${colaboradorId}-${dia}`,
      Semana_id: semanaId,
      Colaborador_id: String(colaboradorId),
      Colaborador_nome: colaboradorNome,
      Dia: dia,
      ...fields
    });
  },

  async saveCheckIn(semanaId, colaboradorId, colaboradorNome, dia, confirmadoPor) {
    return this.registrarCheckIn(semanaId, colaboradorId, colaboradorNome, dia, confirmadoPor);
  },

  // ============================================================
  // DASHBOARD
  // ============================================================
  async getDashboardResumo(semanaId) {
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const diasPt = ["domingo", "segunda", "terca", "quarta", "quinta", "sexta", "sabado"];
    const diaHoje = diasPt[hoje.getDay()];
    const diasUteis = ["segunda", "terca", "quarta", "quinta", "sexta"];
    const diaHojeUtil = diasUteis.includes(diaHoje) ? diaHoje : "segunda";

    const [colabs, pedidos, extras, checkIns] = await Promise.all([
      this.getColaboradores().catch(() => []),
      this.getPedidos(semanaId).catch(() => []),
      this.getExtras(semanaId).catch(() => []),
      this.getCheckIn(semanaId, diaHojeUtil).catch(() => [])
    ]);

    const isConfirmado = p => {
      const s = this.norm(this.pick(p, "Status") || "");
      return s === "confirmado" || s === "extra" || s === "aprovado" || this.isTrue(this.pick(p, "Confirmado"));
    };

    const isCancelado = p => {
      const s = this.norm(this.pick(p, "Status") || "");
      return ["cancelado", "afastado", "ferias", "nao vai almocar", "bloqueado", "travado"].includes(s);
    };

    const isPendente = p => !isConfirmado(p) && !isCancelado(p);
    const isExtraP = p => this.isExtraPedido(p);

    const pedidosColab = pedidos.filter(p => !isExtraP(p));
    const confirmadosColab = pedidosColab.filter(isConfirmado);
    const idsConf = new Set(confirmadosColab.map(p => String(this.pick(p, "Colaborador_id") || "")));

    const pedidosHoje = pedidos.filter(p => this.norm(this.pick(p, "Dia")) === diaHojeUtil);
    const confirmadosHoje = pedidosHoje.filter(isConfirmado);
    const extrasHoje = extras.filter(e => this.norm(this.pick(e, "Dia")) === diaHojeUtil);

    const countOpcao = (lista, opcao) =>
      lista.filter(p => this.norm(this.pick(p, "Opcao")) === opcao).length;

    return {
      colaboradoresAtivos: colabs.length,
      pedidosConfirmadosColaboradores: idsConf.size,
      pendentesColaboradores: Math.max(0, colabs.length - idsConf.size),
      totalPedidosSemana: pedidos.filter(isConfirmado).length,

      checkinsHoje: checkIns.filter(c => this.isTrue(this.pick(c, "Retirou"))).length,
      totalPedidosHoje: confirmadosHoje.length,
      ausenciasHoje: pedidosHoje.filter(isCancelado).length,
      pendentesHoje: pedidosHoje.filter(isPendente).length,

      principalHoje: countOpcao(confirmadosHoje, "principal"),
      lightHoje: countOpcao(confirmadosHoje, "light"),
      outrasHoje: Math.max(
        0,
        confirmadosHoje.length -
        countOpcao(confirmadosHoje, "principal") -
        countOpcao(confirmadosHoje, "light")
      ),

      extrasAtivos: extrasHoje.length,
      extrasConfirmados: extrasHoje.filter(isConfirmado).length,
      extrasPendentes: extrasHoje.filter(isPendente).length,

      porDia: diasUteis.reduce((acc, dia) => {
        const lista = pedidos.filter(p => this.norm(this.pick(p, "Dia")) === dia);
        const conf = lista.filter(isConfirmado);

        acc[dia] = {
          total: conf.length,
          principal: countOpcao(conf, "principal"),
          light: countOpcao(conf, "light"),
          pendentes: lista.filter(isPendente).length,
          cancelados: lista.filter(isCancelado).length
        };

        return acc;
      }, {}),

      setoresHoje: (() => {
        const map = {};

        confirmadosHoje.forEach(p => {
          const setor = this.pick(p, "Centro_Custo") || "Sem CC";
          map[setor] = (map[setor] || 0) + 1;
        });

        return Object.entries(map).sort((a, b) => b[1] - a[1]);
      })(),

      semanaId,
      diaHoje: diaHojeUtil,
      geradoEm: new Date().toISOString()
    };
  },

  // ============================================================
  // COMPATIBILIDADE
  // ============================================================
  async addItem(listName, fields) {
    return this.createItem(listName, fields);
  },

  async patchListItem(listName, id, fields) {
    return this.updateItem(listName, id, fields);
  },

  async removeItem(listName, id) {
    return this.deleteItem(listName, id);
  },

  async cleanupExtraAutomaticoSemana(semanaId) {
    const pedidos = await this.getPedidos(semanaId);
    const seenAuto = new Set();
    const diasUteis = ["segunda", "terca", "quarta", "quinta", "sexta"];

    for (const p of pedidos) {
      const nome = this.norm(this.pick(p, "Colaborador_nome", "Title") || "");
      const origem = this.norm(this.pick(p, "Origem", "tipo") || "");
      const dia = this.norm(this.pick(p, "Dia") || "");

      if (!diasUteis.includes(dia)) continue;

      const isAuto = nome === "refeicaoextra" || (nome.includes("extra") && origem.includes("extra"));
      if (!isAuto) continue;

      const key = `auto-${dia}`;

      if (seenAuto.has(key)) {
        await this.deleteItem("Pedidos", p.id).catch(() => {});
      } else {
        seenAuto.add(key);
      }
    }
  }
};

window.SP = SP;

