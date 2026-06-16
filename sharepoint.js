// ============================================================
// sharepoint.js — Refeitório Homy · Microsoft Graph API
// v: sync-operacao-pedidos-ausencias-20260616
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
  _columnCache:   {},

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

  norm(value) {
    return String(value || "")
      .normalize("NFD").replace(/[̀-ͯ]/g, "")
      .toLowerCase().trim();
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

  // Compatibilidade com páginas públicas antigas/novas do refeitório.
  // Faz apenas a tentativa silenciosa de reaproveitar uma sessão Microsoft já existente.
  // Não abre popup sozinho; quando não houver sessão, a própria página mostra o botão de login.
  async ensureLoginSilenciosoOuAviso() {
    await this.init();
    return !!this._account;
  },

  async ensureLoginSilencioso() {
    return this.ensureLoginSilenciosoOuAviso();
  },

  async loginSilencioso() {
    return this.ensureLoginSilenciosoOuAviso();
  },

  async logout() {
    await this.init();
    const account = this._account;
    this._account = null;
    this._siteId      = null;
    this._listIds     = {};
    this._columnCache = {};
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

  async getListColumns(listName) {
    if (this._columnCache[listName]) return this._columnCache[listName];

    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const data = await this.graph("GET",
      `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`
    );
    const cols = data?.value || [];
    this._columnCache[listName] = cols;
    return cols;
  },

  _normCampo(v) {
    return String(v || "")
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "");
  },

  _aliasesCampo(campo) {
    const k = String(campo || "");
    const mapa = {
      Centro_Custo: [
        "Centro_Custo", "Centro Custo", "Centro de Custo", "CentroCusto",
        "Centro_x0020_Custo", "Centro_x005f_Custo", "Centro_x005f_Custo0",
        "CC", "Centro custo"
      ],
      centroCusto: [
        "Centro_Custo", "Centro Custo", "Centro de Custo", "CentroCusto",
        "Centro_x0020_Custo", "Centro_x005f_Custo", "CC"
      ]
    };
    return [k, ...(mapa[k] || [])];
  },

  _isCampoCentroCusto(campo) {
    return this._normCampo(campo) === "centrocusto" ||
           this._normCampo(campo) === "cc" ||
           this._normCampo(campo) === "centrodecusto";
  },

  _isListaAusencias(listName) {
    const n = this._normCampo(listName);
    return n.includes("ausencia") || n.includes("ausencias");
  },

  _isListaConfiguracoes(listName) {
    const n = this._normCampo(listName);
    return n.includes("configuracao") || n.includes("configuracoes");
  },

  async _mapFieldsToListColumns(listName, fields) {
    const clean = this._cleanFields(fields || {});
    // Ausencias do Refeitorio agora possui Centro_Custo.
    // Mantemos o campo quando a coluna existir; se alguma lista antiga não tiver,
    // o mapeamento abaixo omite apenas o Centro_Custo para não quebrar o salvamento.
    let cols = [];
    try {
      cols = await this.getListColumns(listName);
    } catch (e) {
      console.warn(`[SharePoint] Não foi possível ler colunas de ${listName}; usando campos originais.`, e);
      return clean;
    }

    const byName = new Map(cols.map(c => [String(c.name), c.name]));
    const byNorm = new Map();
    for (const c of cols) {
      byNorm.set(this._normCampo(c.name), c.name);
      byNorm.set(this._normCampo(c.displayName), c.name);
    }

    const out = {};
    for (const [campo, valor] of Object.entries(clean)) {
      let resolved = byName.get(campo) || null;

      if (!resolved) {
        for (const alias of this._aliasesCampo(campo)) {
          resolved = byName.get(alias) || byNorm.get(this._normCampo(alias)) || null;
          if (resolved) break;
        }
      }

      if (resolved) {
        out[resolved] = valor;
        continue;
      }

      // Centro_Custo existe com nomes internos diferentes em algumas listas.
      // Se a lista não tiver essa coluna, o campo é omitido para não quebrar o salvamento.
      if (this._isCampoCentroCusto(campo)) {
        console.warn(`[SharePoint] Campo Centro_Custo não existe em ${listName}; envio omitido.`);
        continue;
      }

      // Configurações tem variações de coluna entre ambientes.
      // Nessa lista, campo não reconhecido deve ser omitido, não enviado ao Graph.
      if (this._isListaConfiguracoes(listName)) {
        console.warn(`[SharePoint] Campo ${campo} não existe em ${listName}; envio omitido.`);
        continue;
      }

      out[campo] = valor;
    }

    return out;
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
    const mappedFields = await this._mapFieldsToListColumns(listName, fields);
    return this.graph("POST", `/sites/${siteId}/lists/${listId}/items`, {
      fields: mappedFields
    });
  },

  async updateItem(listName, itemId, fields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const mappedFields = await this._mapFieldsToListColumns(listName, fields);
    return this.graph("PATCH",
      `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
      mappedFields
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
    let centroCusto = extras.centroCusto ?? extras.Centro_Custo ?? "";
    if (!centroCusto) {
      centroCusto = await this._resolverCentroCustoColaborador(colaboradorId, colaboradorNome);
    }

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
      Centro_Custo:     centroCusto || "",
      Status:           extras.status       ?? extras.Status       ?? "Confirmado",
      Observacao:       extras.observacao   ?? extras.Observacao   ?? "",
      Origem:           extras.origem       ?? extras.Origem       ?? "Refeitório",
      Alterado_Por:     extras.alteradoPor  ?? extras.Alterado_Por ?? this.getUserName()
    };
    this._validarPedidoFields(fields);
    return this.createItem("Pedidos", fields);
  },

  async _getItemFieldsById(listName, itemId) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const data = await this.graph("GET", `/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields`);
    return { id: data?.id || itemId, ...(data?.fields || {}) };
  },

  async _nomePratoCardapioPorOpcao(semanaId, dia, opcao) {
    if (!semanaId || !dia || !opcao) return "";
    const norm = v => String(v || "")
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .toLowerCase().trim();

    let cardapio = [];
    try {
      cardapio = typeof this.getCardapio === "function" ? await this.getCardapio(semanaId) : [];
    } catch (e) {
      console.warn("[SharePoint] Não foi possível buscar cardápio para atualizar Nome_Prato.", e);
      cardapio = [];
    }

    const item = (cardapio || []).find(c =>
      norm(this.pick(c, "Dia", "dia")) === norm(dia) &&
      norm(this.pick(c, "Opcao", "Opção", "opcao")) === norm(opcao)
    );

    return this.pick(item, "Nome_Prato", "nomePrato", "Descricao", "Descrição", "Title") || "";
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

    // Quando o Admin altera Dia ou Opção, o prato precisa acompanhar o cardápio salvo.
    // Ex.: se mudou Principal -> Massa, Nome_Prato deve virar o prato da Massa daquele dia.
    const mudouDiaOuOpcao = fields.Dia !== undefined || fields.Opcao !== undefined;
    const pratoFoiInformado = fields.Nome_Prato !== undefined && String(fields.Nome_Prato || "").trim() !== "";
    if (mudouDiaOuOpcao && !pratoFoiInformado) {
      let atual = {};
      try { atual = await this._getItemFieldsById("Pedidos", id); }
      catch (e) { console.warn("[SharePoint] Não foi possível ler pedido atual para atualizar Nome_Prato.", e); }

      const semanaId = fields.Semana_id || this.pick(atual, "Semana_id", "Semana", "semanaId");
      const dia = fields.Dia || this.pick(atual, "Dia", "dia");
      const opcao = fields.Opcao || this.pick(atual, "Opcao", "Opção", "opcao");
      const nomePrato = await this._nomePratoCardapioPorOpcao(semanaId, dia, opcao);

      if (nomePrato) {
        fields.Nome_Prato = nomePrato;
      } else if (opcao && typeof this._pratoPadraoPorOpcao === "function") {
        fields.Nome_Prato = this._pratoPadraoPorOpcao(opcao);
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

  async addExtra(semanaId, dia, nome, tipo, opcao, observacao, adicionadoPor, centroCusto = "") {
    return this.createItem("Extras", {
      Title:          `${semanaId}-${dia}-${nome}`,
      Semana_id:      semanaId,
      Dia:            dia,
      Nome:           nome,
      tipo:           tipo,
      Opcao:          opcao || "principal",
      Observacao:     observacao || "",
      Adicionado_Por: adicionadoPor || this.getUserName(),
      Centro_Custo:   centroCusto || ""
    });
  },

  _normTexto(v) {
    return String(v || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
  },

  _centroCustoPadraoExtra(nome, tipo) {
    const n = this._normTexto(`${nome} ${tipo}`);
    if (n.includes("guarda") || n.includes("investigador") || n.includes("portaria")) {
      return "120602 - PORTARIA";
    }
    return "120101 - ADM GERAL";
  },

  _pratoPadraoPorOpcao(opcao) {
    const op = this._normTexto(opcao);
    if (op === "light") return "Opção Light";
    if (op === "carne") return "Opção Carne";
    if (op === "massa") return "Opção Massa";
    if (op === "lanche") return "Lanche";
    return "Prato Principal";
  },

  _extraBate(e, semanaId, dia, nome, tipo, opcao) {
    const nDia = this._normTexto(dia);
    const nNome = this._normTexto(nome);
    const nTipo = this._normTexto(tipo);
    const nOpcao = this._normTexto(opcao || "principal");
    return String(this.pick(e, "Semana_id", "Semana") || "") === String(semanaId) &&
      this._normTexto(this.pick(e, "Dia", "dia")) === nDia &&
      this._normTexto(this.pick(e, "Nome", "Title", "Colaborador_nome")) === nNome &&
      (!nTipo || this._normTexto(this.pick(e, "tipo", "Tipo", "Origem")).includes(nTipo)) &&
      this._normTexto(this.pick(e, "Opcao", "opcao") || "principal") === nOpcao;
  },

  _pedidoExtraBate(p, semanaId, dia, nome, tipo, opcao, extraId = "") {
    const obs = String(this.pick(p, "Observacao", "Observação", "observacao") || "");
    if (extraId && obs.includes(`ExtraID:${extraId}`)) return true;

    return String(this.pick(p, "Semana_id", "Semana") || "") === String(semanaId) &&
      this._normTexto(this.pick(p, "Dia", "dia")) === this._normTexto(dia) &&
      this._normTexto(this.pick(p, "Colaborador_nome", "Nome", "Title")) === this._normTexto(nome) &&
      this._normTexto(this.pick(p, "Opcao", "opcao") || "principal") === this._normTexto(opcao || "principal") &&
      this._normTexto(this.pick(p, "Origem", "tipo", "Tipo") || "").includes(this._normTexto(tipo));
  },

  async _addExtraPedidoCC(semanaId, dia, nome, tipo, opcao = "principal", observacao = "", centroCusto = "", adicionadoPor = "") {
    const cc = centroCusto || this._centroCustoPadraoExtra(nome, tipo);
    const user = adicionadoPor || this.getUserName();

    let extras = [];
    try { extras = await this.getExtras(semanaId); } catch (_) { extras = []; }

    let extra = (extras || []).find(e => this._extraBate(e, semanaId, dia, nome, tipo, opcao));
    if (!extra) {
      extra = await this.addExtra(semanaId, dia, nome, tipo, opcao, observacao, user, cc);
      // createItem retorna o item bruto do Graph; normaliza id e campos para reaproveitar abaixo.
      extra = { id: extra?.id, Semana_id: semanaId, Dia: dia, Nome: nome, tipo, Opcao: opcao, Observacao: observacao, Centro_Custo: cc };
    }

    const extraId = String(this.pick(extra, "id", "ID") || "");
    let pedidos = [];
    try { pedidos = await this.getPedidos(semanaId); } catch (_) { pedidos = []; }

    const jaTemPedido = (pedidos || []).some(p =>
      this._pedidoExtraBate(p, semanaId, dia, nome, tipo, opcao, extraId)
    );

    let pedido = null;
    if (!jaTemPedido) {
      const obsPedido = `${observacao || ""}${extraId ? ` | ExtraID:${extraId}` : ""}`.trim();
      pedido = await this.savePedido(
        semanaId,
        extraId ? `extra-${extraId}` : `extra-${this._normTexto(nome)}-${dia}`,
        nome,
        dia,
        opcao || "principal",
        this._pratoPadraoPorOpcao(opcao),
        {
          confirmado:  true,
          status:      "Confirmado",
          origem:      tipo || "extra",
          centroCusto: cc,
          observacao:  obsPedido,
          dataHora:    new Date().toISOString(),
          alteradoPor: user
        }
      );
    }

    return { extra, pedido, criadoPedido: !!pedido };
  },

  async addExtraPedido(semanaId, dia, nome, tipo, opcao, observacao, adicionadoPor, centroCusto = "") {
    return this._addExtraPedidoCC(semanaId, dia, nome, tipo, opcao, observacao, centroCusto, adicionadoPor);
  },

  async deleteExtraComPedido(extra) {
    if (!extra) return false;
    const semanaId = this.pick(extra, "Semana_id", "Semana");
    const dia = this.pick(extra, "Dia", "dia");
    const nome = this.pick(extra, "Nome", "Title", "Colaborador_nome") || "Refeição Extra";
    const tipo = this.pick(extra, "tipo", "Tipo", "Origem") || "extra";
    const opcao = this.pick(extra, "Opcao", "opcao") || "principal";
    const extraId = String(this.pick(extra, "id", "ID") || "");

    if (semanaId) {
      let pedidos = [];
      try { pedidos = await this.getPedidos(semanaId); } catch (_) { pedidos = []; }
      const alvos = (pedidos || []).filter(p =>
        this._pedidoExtraBate(p, semanaId, dia, nome, tipo, opcao, extraId)
      );
      for (const p of alvos) {
        const id = this.pick(p, "id", "ID");
        if (id) { try { await this.deletePedido(id); } catch (e) { console.warn("Falha ao remover pedido espelho do extra:", e); } }
      }
    }

    if (extraId) await this.deleteExtra(extraId);
    return true;
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

  async getMarcacaoLiberada() {
    return this.isCardapioLiberado();
  },

  async getCardapioVisivel() {
    const v = await this.getConfig("cardapio_visivel").catch(() => null);
    if (v !== null && v !== undefined && String(v) !== "") return this.isTrue(v);

    // Compatibilidade com instalações antigas que usavam cardapio_liberado.
    const legado = await this.getConfig("cardapio_liberado").catch(() => null);
    return this.isTrue(legado);
  },

  async setCardapioVisivel(visivel) {
    await this.setConfig("cardapio_visivel", visivel ? "sim" : "nao");
    return true;
  },

  async isMarcacaoLiberada() {
    return this.isCardapioLiberado();
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

  _pickCentroCusto(obj) {
    return this.pick(
      obj,
      "Centro_Custo", "Centro Custo", "Centro de Custo", "CentroCusto",
      "Centro_x0020_Custo", "Centro_x005f_Custo", "Centro_x005f_Custo0",
      "CC", "Setor", "Departamento"
    ) || "";
  },

  async _resolverCentroCustoColaborador(colaboradorId, colaboradorNome) {
    const id = String(colaboradorId || "").trim();
    const nomeNorm = this.norm(colaboradorNome || "");
    if (!id && !nomeNorm) return "";

    try {
      const colaboradores = await this.getTodosColaboradores();
      const encontrado = (colaboradores || []).find(c => {
        const cid = String(c.id || this.pick(c, "ID", "Id") || "").trim();
        const cnome = this.norm(this.pick(c, "Nome", "Title", "Colaborador_nome") || "");
        return (id && cid === id) || (nomeNorm && cnome === nomeNorm);
      });
      return this._pickCentroCusto(encontrado) || "";
    } catch (e) {
      console.warn("[SharePoint] Não foi possível buscar centro de custo do colaborador:", e.message || e);
      return "";
    }
  },

  async createAusencia(dados) {
    const nome   = dados.colaboradorNome || dados.Colaborador_nome || "";
    const motivo = dados.motivo || dados.Motivo || "nao_vai_almocar";
    const colaboradorId = String(dados.colaboradorId || dados.Colaborador_id || "");
    let centroCusto = dados.centroCusto || dados.Centro_Custo || dados.centro_custo || "";

    if (!centroCusto) {
      centroCusto = await this._resolverCentroCustoColaborador(colaboradorId, nome);
    }

    return this.createItem("Ausencias do Refeitorio", {
      Title:            dados.title || `${nome} - ${motivo}`,
      Colaborador_id:   colaboradorId,
      Colaborador_nome: nome,
      Centro_Custo:     centroCusto || "",
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
  // OPERAÇÃO DO DIA — sincronização Pedido x Ausência
  // Regra: Pedidos é a base; Operação altera o próprio Pedido.
  // Quando vira ausência, cria/ativa Ausencias do Refeitorio.
  // Quando volta a comer, inativa a ausência do dia.
  // ============================================================
  _isStatusAusenciaOperacao(status) {
    const s = this.norm(status);
    return [
      "nao vai almocar", "não vai almoçar", "nao_vai_almocar",
      "ferias", "férias", "afastado", "atestado", "licenca", "licença", "ausente"
    ].includes(s);
  },

  _motivoAusenciaOperacao(status) {
    const s = this.norm(status);
    if (s === "nao vai almocar" || s === "não vai almoçar" || s === "nao_vai_almocar" || s === "ausente") return "Não vai almoçar";
    if (s === "ferias" || s === "férias") return "Férias";
    if (s === "afastado") return "Afastado";
    if (s === "atestado") return "Atestado";
    if (s === "licenca" || s === "licença") return "Licença";
    return status || "Não vai almoçar";
  },

  _pedidoDataRefeicao(semanaId, dia, pedido = {}) {
    if (semanaId && dia && typeof this.getDataRefBySemanaDia === "function") {
      return this.getDataRefBySemanaDia(semanaId, dia);
    }
    const data = this.pick(pedido, "Data_Refeicao", "Data_Referencia", "Data") || "";
    if (/^\d{4}-\d{2}-\d{2}/.test(String(data))) return String(data).slice(0, 10);
    return new Date().toISOString().slice(0, 10);
  },

  _mesmoColaboradorAusenciaPedido(a, p) {
    const aid = String(this.pick(a, "Colaborador_id", "ColaboradorId", "colaborador_id") || "").trim();
    const pid = String(this.pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    if (aid && pid && aid === pid) return true;

    const an = this.norm(this.pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title") || "");
    const pn = this.norm(this.pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "");
    return !!an && !!pn && an === pn;
  },

  _ausenciaAtivaOperacao(a) {
    const ativo = this.pick(a, "Ativo", "ativo");
    const status = this.norm(this.pick(a, "Status", "status") || "");
    if (["inativo", "cancelado", "false", "nao", "não", "0"].includes(status)) return false;
    if (ativo === null || ativo === undefined || ativo === "") return true;
    return this.isTrue(ativo);
  },

  _ausenciaCobreDataOperacao(a, dataRef) {
    const iniRaw = this.pick(a, "Data_Inicio", "Inicio", "DataInicio", "Data") || "";
    const fimRaw = this.pick(a, "Data_Fim", "Fim", "DataFim", "Data") || iniRaw;
    const ini = String(iniRaw).slice(0, 10);
    const fim = String(fimRaw).slice(0, 10);
    return !!ini && !!fim && ini <= dataRef && fim >= dataRef;
  },

  async _buscarAusenciasDoPedidoOperacao(pedido, semanaId, dia) {
    const dataRef = this._pedidoDataRefeicao(semanaId, dia, pedido);
    let ausencias = [];
    try { ausencias = await this.getAusencias(false); }
    catch (e) { console.warn("[SharePoint] Não foi possível buscar ausências para sincronizar operação:", e); }
    return (ausencias || []).filter(a =>
      this._ausenciaAtivaOperacao(a) &&
      this._mesmoColaboradorAusenciaPedido(a, pedido) &&
      this._ausenciaCobreDataOperacao(a, dataRef)
    );
  },

  async _inativarAusenciasDoPedidoOperacao(pedido, semanaId, dia) {
    const ausencias = await this._buscarAusenciasDoPedidoOperacao(pedido, semanaId, dia);
    for (const a of ausencias) {
      const id = this.pick(a, "id", "ID");
      if (!id) continue;
      try { await this.updateAusencia(id, { Ativo: false }); }
      catch (e) { console.warn("[SharePoint] Falha ao inativar ausência após ajuste da operação:", e); }
    }
    return ausencias.length;
  },

  async _garantirAusenciaDoPedidoOperacao(pedido, status, semanaId, dia) {
    const motivo = this._motivoAusenciaOperacao(status);
    const dataRef = this._pedidoDataRefeicao(semanaId, dia, pedido);
    const existentes = await this._buscarAusenciasDoPedidoOperacao(pedido, semanaId, dia);
    const idExistente = this.pick(existentes[0], "id", "ID") || "";

    const colaboradorId = String(this.pick(pedido, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    const nome = this.pick(pedido, "Colaborador_nome", "Colaborador", "Nome", "Title") || "";
    let centroCusto = this.pick(pedido, "Centro_Custo", "CentroCusto", "Setor", "Departamento") || "";
    if (!centroCusto) centroCusto = await this._resolverCentroCustoColaborador(colaboradorId, nome);

    if (idExistente) {
      await this.updateAusencia(idExistente, {
        Ativo: true,
        Motivo: motivo,
        Data_Inicio: dataRef,
        Data_Fim: dataRef,
        Observacao: this.pick(existentes[0], "Observacao", "Observação") || "Sincronizada pela Operação do Dia."
      });
      return idExistente;
    }

    const criada = await this.createAusencia({
      colaboradorId,
      colaboradorNome: nome,
      centroCusto,
      dataInicio: dataRef,
      dataFim: dataRef,
      motivo,
      observacao: "Sincronizada pela Operação do Dia.",
      ativo: true,
      criadoPor: this.getUserName()
    });
    return this.pick(criada, "id", "ID") || "";
  },



  async _sincronizarPedidosDuplicadosOperacao(semanaId, dia, pedidoBase, fields, idPrincipal = "") {
    // Corrige dados antigos com mais de um Pedido para o mesmo colaborador/dia.
    // A Operação do Dia precisa deixar todos com a mesma verdade para Pedidos, Cozinha e Relatórios.
    const colabId = String(this.pick(pedidoBase, "Colaborador_id", "ColaboradorId", "colaboradorId") || fields.Colaborador_id || "").trim();
    const nomeNorm = this.norm(this.pick(pedidoBase, "Colaborador_nome", "Colaborador", "Nome", "Title") || fields.Colaborador_nome || "");
    if (String(colabId).startsWith("extra-") || (!colabId && !nomeNorm)) return 0;

    let pedidos = [];
    try { pedidos = await this.getPedidos(semanaId); }
    catch (e) { console.warn("[SharePoint] Não foi possível verificar duplicidades de pedidos:", e); return 0; }

    let atualizados = 0;
    for (const p of (pedidos || [])) {
      const pid = String(this.pick(p, "id", "ID") || "");
      if (!pid || String(pid) === String(idPrincipal || "")) continue;
      if (this.norm(this.pick(p, "Dia", "dia") || "") !== this.norm(dia)) continue;

      const pId = String(this.pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
      const pNome = this.norm(this.pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "");
      const mesmo = (colabId && pId && pId === colabId) || (!colabId && nomeNorm && pNome === nomeNorm) || (colabId && !pId && nomeNorm && pNome === nomeNorm);
      if (!mesmo) continue;

      try {
        await this.updatePedido(pid, {
          Status: fields.Status,
          Confirmado: fields.Confirmado,
          Opcao: fields.Opcao,
          Nome_Prato: fields.Nome_Prato,
          Centro_Custo: fields.Centro_Custo,
          Data_Hora: fields.Data_Hora,
          Origem: fields.Origem,
          Observacao: fields.Observacao,
          Alterado_Por: fields.Alterado_Por || this.getUserName()
        });
        atualizados++;
      } catch (e) {
        console.warn("[SharePoint] Falha ao sincronizar pedido duplicado da operação:", e);
      }
    }
    return atualizados;
  },

  async alterarPedidoOperacao(id, status, pedidoAtual = {}) {
    const isVirtual = !id || String(id).startsWith("ausencia-") || pedidoAtual?._virtualAusencia;

    let atual = {};
    if (!isVirtual) {
      try { atual = await this._getItemFieldsById("Pedidos", id); }
      catch (e) { console.warn("[SharePoint] Não foi possível ler pedido atual na Operação do Dia:", e); }
    }

    const base = { ...(atual || {}), ...(pedidoAtual || {}) };
    const semanaId = this.pick(base, "Semana_id", "Semana", "semanaId") || this.getCurrentWeekId();
    const dia = this.pick(base, "Dia", "dia") || "segunda";
    const colaboradorId = String(this.pick(base, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    const colaboradorNome = this.pick(base, "Colaborador_nome", "Colaborador", "Nome", "Title") || "";
    const opcao = this.pick(base, "Opcao", "Opção", "opcao") || "principal";
    const origemAtual = this.pick(base, "Origem", "origem") || "Refeitório";
    const centroCusto = this.pick(base, "Centro_Custo", "CentroCusto", "centroCusto", "Setor", "Departamento") || await this._resolverCentroCustoColaborador(colaboradorId, colaboradorNome);
    const dataHora = `${this._pedidoDataRefeicao(semanaId, dia, base)}T12:00:00`;

    if (!colaboradorId && !colaboradorNome) throw new Error("Não foi possível identificar o colaborador do pedido.");

    const statusNorm = this.norm(status);
    let fields = {
      Semana_id: semanaId,
      Colaborador_id: colaboradorId || this.norm(colaboradorNome),
      Colaborador_nome: colaboradorNome,
      Dia: dia,
      Opcao: opcao,
      Centro_Custo: centroCusto || "",
      Data_Hora: dataHora,
      Alterado_Por: this.getUserName()
    };

    if (this._isStatusAusenciaOperacao(status)) {
      const motivo = this._motivoAusenciaOperacao(status);
      fields = {
        ...fields,
        Nome_Prato: motivo,
        Confirmado: false,
        Status: motivo,
        Origem: origemAtual || "Refeitório",
        Observacao: this.pick(base, "Observacao", "Observação") || "Marcado como ausência pela Operação do Dia."
      };

      let pedido = null;
      if (!isVirtual && id) pedido = await this.updatePedido(id, fields);
      else pedido = await this.savePedido(semanaId, fields.Colaborador_id, colaboradorNome, dia, opcao, motivo, fields);
      const pedidoIdFinal = !isVirtual && id ? id : this.pick(pedido, "id", "ID");
      await this._sincronizarPedidosDuplicadosOperacao(semanaId, dia, { ...base, ...fields }, fields, pedidoIdFinal);
      await this._garantirAusenciaDoPedidoOperacao({ ...base, ...fields, id: pedidoIdFinal }, status, semanaId, dia);
      try { localStorage.setItem("homy_refeitorio_sync", JSON.stringify({ tipo:"pedido", ts:Date.now(), semanaId, dia })); } catch (_) {}
      return pedido;
    }

    if (statusNorm === "confirmado" || statusNorm === "extra") {
      const nomePrato = await this._nomePratoCardapioPorOpcao(semanaId, dia, opcao) || this._pratoPadraoPorOpcao(opcao);
      fields = {
        ...fields,
        Nome_Prato: nomePrato,
        Confirmado: true,
        Status: "Confirmado",
        Origem: origemAtual && this.norm(origemAtual) !== "ausencia" ? origemAtual : "Operação do Dia",
        Observacao: this.pick(base, "Observacao", "Observação") || "Confirmado pela Operação do Dia."
      };

      let pedido = null;
      if (!isVirtual && id) pedido = await this.updatePedido(id, fields);
      else pedido = await this.savePedido(semanaId, fields.Colaborador_id, colaboradorNome, dia, opcao, nomePrato, fields);
      const pedidoIdFinal = !isVirtual && id ? id : this.pick(pedido, "id", "ID");
      await this._sincronizarPedidosDuplicadosOperacao(semanaId, dia, { ...base, ...fields }, fields, pedidoIdFinal);
      await this._inativarAusenciasDoPedidoOperacao({ ...base, ...fields, id: pedidoIdFinal }, semanaId, dia);
      try { localStorage.setItem("homy_refeitorio_sync", JSON.stringify({ tipo:"pedido", ts:Date.now(), semanaId, dia })); } catch (_) {}
      return pedido;
    }

    if (statusNorm === "cancelado") {
      fields = {
        ...fields,
        Confirmado: false,
        Status: "Cancelado",
        Origem: origemAtual || "Operação do Dia",
        Observacao: this.pick(base, "Observacao", "Observação") || "Cancelado pela Operação do Dia."
      };
      let pedido = null;
      if (!isVirtual && id) pedido = await this.updatePedido(id, fields);
      else pedido = await this.savePedido(semanaId, fields.Colaborador_id, colaboradorNome, dia, opcao, this.pick(base, "Nome_Prato") || "Cancelado", fields);
      const pedidoIdFinal = !isVirtual && id ? id : this.pick(pedido, "id", "ID");
      await this._sincronizarPedidosDuplicadosOperacao(semanaId, dia, { ...base, ...fields }, fields, pedidoIdFinal);
      await this._inativarAusenciasDoPedidoOperacao({ ...base, ...fields, id: pedidoIdFinal }, semanaId, dia);
      try { localStorage.setItem("homy_refeitorio_sync", JSON.stringify({ tipo:"pedido", ts:Date.now(), semanaId, dia })); } catch (_) {}
      return pedido;
    }

    return this.updatePedido(id, { Status: status, Alterado_Por: this.getUserName() });
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

// Export explícito para páginas que verificam window.SP antes do login.
if (typeof window !== 'undefined') window.SP = SP;
