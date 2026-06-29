// ============================================================
// sharepoint.js — Refeitório Homy · Microsoft Graph API
// v: base-limpa-v10-4-20260629
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

  // Cache curto para reduzir chamadas repetidas ao Graph.
  // A fonte da verdade continua sendo o SharePoint; o cache é só de leitura e expira rápido.
  _itemsCache: {},
  _pendingItemRequests: {},
  _repairCache: {},
  _repairRunning: {},
  _ausenciasEncerradasLastRun: 0,
  _ITEMS_CACHE_TTL_MS: 15000,
  _CONFIG_CACHE_TTL_MS: 45000,
  _REPAIR_TTL_MS: 180000,

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
  // REGRAS DE INTEGRIDADE OPERACIONAL v7
  // ============================================================
  _dataISOOperacional(v) {
    if (!v) return "";
    if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
    const s = String(v || "").trim();
    const mIso = s.match(/^(\d{4}-\d{2}-\d{2})/);
    if (mIso) return mIso[1];
    const mBr = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (mBr) return `${mBr[3]}-${mBr[2]}-${mBr[1]}`;
    const d = new Date(s);
    return isNaN(d) ? "" : d.toISOString().slice(0, 10);
  },

  _diaOperacionalPassado(diaInfoOuData) {
    const data = this._dataISOOperacional(diaInfoOuData?.data || diaInfoOuData);
    if (!data) return false;
    const hoje = typeof this._hojeISO === "function" ? this._hojeISO() : new Date().toISOString().slice(0, 10);
    return data < hoje;
  },

  _podeGerarRetornoAutomatico(diaInfo, options = {}) {
    if (options && options.permitirRetornoRetroativo === true) return true;
    return !this._diaOperacionalPassado(diaInfo);
  },

  _statusBloqueiaProducao(status) {
    const s = this.norm(status);
    return [
      "cancelado", "bloqueado",
      "nao vai almocar", "não vai almoçar", "nao_vai_almocar",
      "ausente", "ferias", "férias", "afastado", "atestado",
      "licenca", "licença", "banco horas", "banco_horas",
      "homy office", "homy_office", "falta"
    ].includes(s);
  },

  _pedidoProdutivoValido(p) {
    if (!p) return false;
    const status = this.pick(p, "Status", "status") || "";
    const statusNorm = this.norm(status);
    const origemNorm = this.norm(this.pick(p, "Origem", "origem", "tipo", "Tipo") || "");
    if (statusNorm === "travado" && origemNorm.includes("travamento")) return true;
    if (this._statusBloqueiaProducao(status)) return false;
    const confirmado = this.isTrue(this.pick(p, "Confirmado", "confirmado"));
    return confirmado || ["confirmado", "aprovado", "extra", "travado"].includes(statusNorm);
  },

  _extraIdempotenteKey(semanaId, dia, nome, tipo, opcao) {
    return [semanaId, dia, nome, tipo, opcao || "principal"]
      .map(v => this.norm(v || ""))
      .join("|");
  },

  _pedidoExtraAtivoEquivalente(p, semanaId, dia, nome, tipo, opcao, extraId = "") {
    if (!this._pedidoExtraBate(p, semanaId, dia, nome, tipo, opcao, extraId)) return false;
    return this._pedidoProdutivoValido(p);
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
  // PERFORMANCE / CACHE
  // ============================================================
  _cloneItems(items) {
    return (items || []).map(i => ({ ...i }));
  },

  _cacheTTL(listName) {
    const n = String(listName || "").toLowerCase();
    if (n.includes("config")) return this._CONFIG_CACHE_TTL_MS;
    return this._ITEMS_CACHE_TTL_MS;
  },

  clearListCache(listName) {
    if (!listName) return;
    delete this._itemsCache[listName];
    if (String(listName).toLowerCase().includes("config")) delete this._itemsCache["Configurações"];
  },

  clearAllCache() {
    this._itemsCache = {};
    this._pendingItemRequests = {};
  },

  _emitSync(tipo = "dados", detalhe = "") {
    const payload = { tipo, detalhe, ts: Date.now() };
    try { localStorage.setItem("homy_refeitorio_sync", JSON.stringify(payload)); } catch (_) {}
    try {
      if (!this._syncChannel) this._syncChannel = new BroadcastChannel("homy-refeitorio-sync");
      this._syncChannel.postMessage(payload);
    } catch (_) {}
  },

  _reparoMudouDados(r) {
    if (!r) return false;
    return !!(
      r.criados || r.atualizados || r.ignoradosAtualizados ||
      r.ausenciasCriadas || r.ausenciasAtualizadas ||
      r.retornoCriados || r.retornoAtualizados ||
      r.duplicadasInativadas || r.encerradas ||
      r.pedidosAtualizados || r.pedidosCancelados
    );
  },

  async repararIntegridadeSemana(semanaId, options = {}) {
    if (!semanaId) return { executado: false, motivo: "sem semana" };
    const force = !!options.force;
    const now = Date.now();
    const last = this._repairCache[semanaId] || 0;

    if (!force && last && (now - last) < this._REPAIR_TTL_MS) {
      return { executado: false, motivo: "reparo recente" };
    }
    if (this._repairRunning[semanaId]) {
      return this._repairRunning[semanaId];
    }

    this._repairCache[semanaId] = now;
    const run = (async () => {
      const resultado = {
        executado: true,
        extras: null,
        ausencias: null,
        mudouDados: false
      };
      try {
        let pedidosBase = options.pedidosBase || null;
        if (!pedidosBase) {
          const todos = await this.getItems("Pedidos", { force: !!force });
          pedidosBase = (todos || []).filter(i => this.pick(i, "Semana_id") === semanaId);
        }

        if (typeof this.sincronizarAusenciasPedidosSemana === "function") {
          resultado.ausencias = await this.sincronizarAusenciasPedidosSemana(semanaId, pedidosBase)
            .catch(e => ({ erro: e.message || String(e) }));
        } else if (typeof this.sincronizarAusenciasEncerradas === "function") {
          resultado.ausencias = await this.sincronizarAusenciasEncerradas(null, { force })
            .catch(e => ({ erro: e.message || String(e) }));
        }

        if (typeof this.garantirExtrasComoPedidos === "function") {
          resultado.extras = await this.garantirExtrasComoPedidos(semanaId, pedidosBase)
            .catch(e => ({ erro: e.message || String(e) }));
        }

        resultado.mudouDados = this._reparoMudouDados(resultado.ausencias) || this._reparoMudouDados(resultado.extras);
        if (resultado.mudouDados) {
          this.clearListCache("Pedidos");
          this.clearListCache("Ausencias do Refeitorio");
          this.clearListCache("Extras");
          this._emitSync("integridade", semanaId);
        }
        return resultado;
      } finally {
        delete this._repairRunning[semanaId];
      }
    })();

    this._repairRunning[semanaId] = run;
    return run;
  },

  agendarReparoIntegridadeSemana(semanaId, pedidosBase = null) {
    // v10.4 — leitura nunca agenda gravação em segundo plano.
    // Mantida apenas por compatibilidade com chamadas antigas. Qualquer reparo
    // que altere Pedidos/Ausências/Extras deve ser acionado por botão explícito,
    // com confirmação e auditoria.
    console.info("[SharePoint] Reparo automático não agendado: leituras são somente leitura.", semanaId || "");
    return { agendado: false, somenteLeitura: true, motivo: "reparo automático desativado" };
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

  async getItems(listName, options = {}) {
    const force = !!options.force;
    const ttl = Number.isFinite(options.ttl) ? options.ttl : this._cacheTTL(listName);
    const key = String(listName || "");
    const now = Date.now();
    const cached = this._itemsCache[key];

    if (!force && cached && (now - cached.ts) < ttl) {
      return this._cloneItems(cached.data);
    }

    if (!force && this._pendingItemRequests[key]) {
      const data = await this._pendingItemRequests[key];
      return this._cloneItems(data);
    }

    const request = (async () => {
      const siteId = await this.getSiteId();
      const listId = await this.getListId(listName);

      let items = [];
      let endpoint = `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500`;

      while (endpoint) {
        const data = await this.graph("GET", endpoint);
        items.push(...(data?.value || []));
        endpoint = data?.["@odata.nextLink"]
          ? data["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "")
          : null;
      }

      const mapped = items.map(i => ({ id: i.id, ...i.fields }));
      this._itemsCache[key] = { ts: Date.now(), data: mapped };
      return mapped;
    })();

    this._pendingItemRequests[key] = request;
    try {
      const data = await request;
      return this._cloneItems(data);
    } finally {
      delete this._pendingItemRequests[key];
    }
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

      // Campo não encontrado na lista real do SharePoint.
      // Regra de integridade: nunca enviar campo inexistente ao Graph,
      // porque isso quebra o salvamento inteiro da lista. A coluna pode ter
      // nome visual diferente, nome interno diferente ou simplesmente não existir.
      console.warn(`[SharePoint] Campo ${campo} não existe em ${listName}; envio omitido para preservar a gravação.`);
      continue;
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
    const result = await this.graph("POST", `/sites/${siteId}/lists/${listId}/items`, {
      fields: mappedFields
    });
    this.clearListCache(listName);
    this._emitSync(String(listName || "").toLowerCase(), "create");
    return result;
  },

  async updateItem(listName, itemId, fields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const mappedFields = await this._mapFieldsToListColumns(listName, fields);
    const result = await this.graph("PATCH",
      `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
      mappedFields
    );
    this.clearListCache(listName);
    this._emitSync(String(listName || "").toLowerCase(), "update");
    return result;
  },

  async deleteItem(listName, itemId) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const result = await this.graph("DELETE", `/sites/${siteId}/lists/${listId}/items/${itemId}`);
    this.clearListCache(listName);
    this._emitSync(String(listName || "").toLowerCase(), "delete");
    return result;
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
  async getPedidos(semanaId, options = {}) {
    const items = await this.getItems("Pedidos", { force: !!options.force, ttl: this._ITEMS_CACHE_TTL_MS });
    const pedidos = items.filter(i => this.pick(i, "Semana_id") === semanaId);

    // v10.4 — leitura nunca grava.
    // Antes, getPedidos() podia agendar repararIntegridadeSemana(), que por sua vez
    // podia criar/cancelar/atualizar pedidos em segundo plano apenas ao abrir telas.
    // A partir daqui, qualquer reparo/travamento/correção precisa de ação explícita.
    return pedidos;
  },

  _extraValor(extra, ...keys) {
    return this.pick(extra, ...keys) ?? "";
  },

  _extraInativo(extra) {
    const status = this.norm(this._extraValor(extra, "Status", "status") || "");
    if (["cancelado", "bloqueado", "excluido", "excluído", "inativo", "false", "nao", "não", "0"].includes(status)) return true;
    const ativo = this._extraValor(extra, "Ativo", "ativo");
    if (ativo === null || ativo === undefined || String(ativo).trim() === "") return false;
    return !this.isTrue(ativo);
  },

  _extraDiaValor(extra) {
    return this._extraValor(extra, "Dia", "dia") || "";
  },

  _extraNomeValor(extra) {
    return this._extraValor(extra, "Nome", "Title", "Colaborador_nome") || "Refeição Extra";
  },

  _extraTipoValor(extra) {
    return this._extraValor(extra, "tipo", "Tipo", "Origem", "origem") || "extra";
  },

  _extraOpcaoValor(extra) {
    return this._extraValor(extra, "Opcao", "Opção", "opcao") || "principal";
  },

  _extraCentroCustoValor(extra) {
    const raw = this._extraValor(extra, "Centro_Custo", "CentroCusto", "Setor", "Departamento") || "";
    if (raw) return raw;
    const obs = String(this._extraValor(extra, "Observacao", "Observação", "observacao") || "");
    const m = obs.match(/(\d{6})(?:\s*-\s*([A-Za-zÀ-ÿ\s/.-]+))?/);
    if (m) return m[0];
    return this._centroCustoPadraoExtra(this._extraNomeValor(extra), this._extraTipoValor(extra));
  },

  _pedidoEspelhoDoExtra(pedido, extra, semanaId) {
    const dia = this._extraDiaValor(extra);
    const nome = this._extraNomeValor(extra);
    const tipo = this._extraTipoValor(extra);
    const opcao = this._extraOpcaoValor(extra);
    const extraId = String(this.pick(extra, "id", "ID") || "").trim();
    const colabId = String(this.pick(pedido, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    const obs = String(this.pick(pedido, "Observacao", "Observação", "observacao") || "");

    if (extraId && colabId === `extra-${extraId}`) return true;
    if (extraId && this.norm(obs).includes(`extraid:${this.norm(extraId)}`)) return true;
    if (typeof this._pedidoExtraBate === "function" && this._pedidoExtraBate(pedido, semanaId, dia, nome, tipo, opcao, extraId)) return true;

    return String(this.pick(pedido, "Semana_id", "Semana") || "") === String(semanaId) &&
      this.norm(this.pick(pedido, "Dia", "dia") || "") === this.norm(dia) &&
      this.norm(this.pick(pedido, "Colaborador_nome", "Nome", "Title") || "") === this.norm(nome) &&
      this.norm(this.pick(pedido, "Origem", "tipo", "Tipo") || "").includes(this.norm(tipo)) &&
      this.norm(this.pick(pedido, "Opcao", "opcao") || "principal") === this.norm(opcao || "principal");
  },

  async garantirExtrasComoPedidos(semanaId, pedidosBase = null) {
    if (!semanaId) return { criados: 0, atualizados: 0, ignorados: 0 };
    if (this._sincronizandoExtrasComoPedidos) return { criados: 0, atualizados: 0, ignorados: 0 };

    this._sincronizandoExtrasComoPedidos = true;
    let criados = 0, atualizados = 0, ignorados = 0;

    try {
      const [extras, pedidosIniciais] = await Promise.all([
        this.getItems("Extras").catch(() => []),
        pedidosBase ? Promise.resolve(pedidosBase) : this.getItems("Pedidos").then(items => items.filter(i => this.pick(i, "Semana_id") === semanaId)).catch(() => [])
      ]);

      const pedidos = Array.isArray(pedidosIniciais) ? [...pedidosIniciais] : [];
      const extrasSemana = (extras || []).filter(e =>
        String(this.pick(e, "Semana_id", "Semana") || "") === String(semanaId) &&
        !this._extraInativo(e) &&
        String(this._extraDiaValor(e) || "").trim() &&
        String(this._extraNomeValor(e) || "").trim()
      );

      const extrasProcessados = new Set();

      for (const extra of extrasSemana) {
        const dia = this._extraDiaValor(extra);
        const nome = this._extraNomeValor(extra);
        const tipo = this._extraTipoValor(extra);
        const opcao = this._extraOpcaoValor(extra);
        const extraKey = this._extraIdempotenteKey(semanaId, dia, nome, tipo, opcao);
        if (extrasProcessados.has(extraKey)) {
          ignorados++;
          continue;
        }
        extrasProcessados.add(extraKey);
        const extraId = String(this.pick(extra, "id", "ID") || "").trim();
        const colabId = extraId ? `extra-${extraId}` : `extra-${this.norm(tipo)}-${this.norm(nome)}-${this.norm(dia)}`;
        const centroCusto = this._extraCentroCustoValor(extra);
        const obsBase = this._extraValor(extra, "Observacao", "Observação", "observacao") || "Criado a partir da lista Extras";
        const observacao = [obsBase, extraId ? `ExtraID:${extraId}` : ""].filter(Boolean).join(" | ");
        const nomePrato = await this._nomePratoCardapioPorOpcao(semanaId, dia, opcao).catch(() => "") || this._pratoPadraoPorOpcao(opcao);
        const dataHora = `${this.getDataRefBySemanaDia(semanaId, dia)}T12:00:00`;
        const existente = pedidos.find(p => this._pedidoEspelhoDoExtra(p, extra, semanaId));

        const fields = {
          Title:            `${semanaId}-${colabId}-${dia}`,
          Semana_id:        semanaId,
          Colaborador_id:   colabId,
          Colaborador_nome: nome,
          Dia:              dia,
          Opcao:            opcao || "principal",
          Nome_Prato:       nomePrato,
          Confirmado:       true,
          Data_Hora:        dataHora,
          Centro_Custo:     centroCusto || "",
          Status:           "Confirmado",
          Observacao:       observacao,
          Origem:           tipo || "extra",
          Alterado_Por:     this.getUserName ? this.getUserName() : "Sistema"
        };

        if (existente?.id) {
          const precisaAtualizar = ["Colaborador_id", "Colaborador_nome", "Dia", "Opcao", "Nome_Prato", "Centro_Custo", "Status", "Origem"].some(k =>
            String(this.pick(existente, k) || "") !== String(fields[k] || "")
          ) || !this.isTrue(this.pick(existente, "Confirmado"));

          if (precisaAtualizar) {
            await this.updateItem("Pedidos", existente.id, fields);
            Object.assign(existente, fields);
            atualizados++;
          } else {
            ignorados++;
          }
        } else {
          const criado = await this.createItem("Pedidos", fields);
          pedidos.push({ id: criado?.id || criado?.ID || "", ...fields });
          criados++;
        }
      }
    } finally {
      this._sincronizandoExtrasComoPedidos = false;
    }

    return { criados, atualizados, ignorados };
  },

  async getPedidoColaborador(semanaId, colaboradorId) {
    const items = await this.getPedidos(semanaId);
    return items.filter(i =>
      String(this.pick(i, "Colaborador_id")) === String(colaboradorId)
    );
  },

  _pedidoMesmoColaboradorDia(p, semanaId, colaboradorId, colaboradorNome, dia) {
    const pSemana = String(this.pick(p, "Semana_id", "Semana", "semanaId") || "").trim();
    const pDia = this.norm(this.pick(p, "Dia", "dia") || "");
    if (pSemana !== String(semanaId || "").trim()) return false;
    if (pDia !== this.norm(dia || "")) return false;

    const idAlvo = String(colaboradorId || "").trim();
    const nomeAlvo = this.norm(colaboradorNome || "");
    const pId = String(this.pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    const pNome = this.norm(this.pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "");

    if (idAlvo && pId && pId === idAlvo) return true;
    if (!idAlvo && nomeAlvo && pNome === nomeAlvo) return true;
    if (idAlvo && !pId && nomeAlvo && pNome === nomeAlvo) return true;
    return false;
  },

  _isPedidoAdicionalColaborador(p) {
    const origem = this.norm(this.pick(p, "Origem", "origem", "tipo", "Tipo") || "");
    const obs = this.norm(this.pick(p, "Observacao", "Observação", "observacao") || "");
    const cid = this.norm(this.pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId") || "");
    return origem.includes("segunda refeicao") || origem.includes("segunda refeição") ||
           origem.includes("transferencia para hoje") || origem.includes("transferência para hoje") ||
           origem.includes("refeicao adicional") || origem.includes("refeição adicional") ||
           obs.includes("adicionalid:") || obs.includes("colaboradorbaseid:") ||
           cid.includes("-adicional-");
  },

  _pedidoTimestampOrdem(p) {
    const raw = this.pick(p, "Modified", "modified", "Data_Hora", "DataHora", "Created", "created", "Data") || "";
    const dt = raw ? new Date(raw) : null;
    if (dt && !isNaN(dt)) return dt.getTime();
    const id = Number(this.pick(p, "id", "ID") || 0);
    return Number.isFinite(id) ? id : 0;
  },

  _pedidoStatusPeso(p) {
    const status = this.norm(this.pick(p, "Status", "status") || "");
    const origem = this.norm(this.pick(p, "Origem", "origem", "tipo", "Tipo") || "");
    const confirmado = this.isTrue(this.pick(p, "Confirmado", "confirmado"));
    if (["nao vai almocar", "não vai almoçar", "ausente", "ferias", "férias", "afastado", "atestado", "licenca", "licença"].includes(status)) return 90;
    if (["confirmado", "aprovado", "extra"].includes(status) || confirmado) return origem.includes("travamento") ? 70 : 80;
    if (status === "travado") return origem.includes("travamento") ? 70 : 65;
    if (status === "cancelado" || status === "bloqueado") return 50;
    return 10;
  },

  _ordenarPedidoMaisAtual(a, b) {
    const ta = this._pedidoTimestampOrdem(a);
    const tb = this._pedidoTimestampOrdem(b);
    if (ta !== tb) return tb - ta;
    return this._pedidoStatusPeso(b) - this._pedidoStatusPeso(a);
  },

  async _buscarPedidoExistenteParaUpsert(semanaId, colaboradorId, colaboradorNome, dia) {
    let pedidos = [];
    try { pedidos = await this.getPedidos(semanaId); }
    catch (e) { console.warn("[SharePoint] Não foi possível verificar pedido existente antes de salvar.", e); return null; }

    const candidatos = (pedidos || []).filter(p => !this._isPedidoAdicionalColaborador(p) && this._pedidoMesmoColaboradorDia(p, semanaId, colaboradorId, colaboradorNome, dia));
    if (!candidatos.length) return null;
    candidatos.sort((a, b) => this._ordenarPedidoMaisAtual(a, b));
    return candidatos[0];
  },

  async savePedido(semanaId, colaboradorId, colaboradorNome, dia, opcao, nomePrato, extras = {}) {
    if (!semanaId) {
      semanaId = await this.getSemanaAlvoMarcacao().catch(() => this.getCurrentWeekId());
    }
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

    // Refeição adicional vinculada ao colaborador é exceção controlada:
    // deve criar uma segunda linha intencional no mesmo dia, sem sobrescrever o pedido normal.
    if (this._isPedidoAdicionalColaborador(fields)) {
      return this.createItem("Pedidos", fields);
    }

    // Regra importante: Pedido normal é 1 registro por colaborador + semana + dia.
    // Antes o sistema criava novos registros sempre. Quando ficavam duplicados
    // com status/origem diferentes, uma tela podia ler o antigo e o colaborador "sumia".
    const existente = await this._buscarPedidoExistenteParaUpsert(semanaId, fields.Colaborador_id, fields.Colaborador_nome, dia);
    if (existente?.id) {
      return this.updatePedido(existente.id, fields);
    }

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
    // v7 — Extra idempotente: não criar outro registro igual para a mesma semana/dia/nome/tipo/opção.
    try {
      const extras = await this.getItems("Extras", { force: true });
      const existente = (extras || []).find(e =>
        !this._extraInativo(e) &&
        this._extraBate(e, semanaId, dia, nome, tipo, opcao || "principal")
      );
      if (existente?.id) return existente;
    } catch (e) {
      console.warn("[SharePoint] Não foi possível verificar extra existente antes de criar.", e);
    }

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
    // v7 — Extra idempotente:
    // - trava chamadas paralelas com a mesma chave operacional;
    // - consulta SharePoint sem cache antes de criar;
    // - reaproveita pedido produtivo equivalente quando já existir.
    if (!this._extraPedidoLocks) this._extraPedidoLocks = {};
    const lockKey = this._extraIdempotenteKey(semanaId, dia, nome, tipo, opcao);
    if (this._extraPedidoLocks[lockKey]) return this._extraPedidoLocks[lockKey];

    const run = (async () => {
      const cc = centroCusto || this._centroCustoPadraoExtra(nome, tipo);
      const user = adicionadoPor || this.getUserName();

      let extras = [];
      try {
        extras = (await this.getItems("Extras", { force: true }) || [])
          .filter(e => String(this.pick(e, "Semana_id", "Semana") || "") === String(semanaId));
      } catch (_) {
        extras = [];
      }

      let extra = (extras || []).find(e =>
        !this._extraInativo(e) &&
        this._extraBate(e, semanaId, dia, nome, tipo, opcao)
      );

      if (!extra) {
        extra = await this.addExtra(semanaId, dia, nome, tipo, opcao, observacao, user, cc);
        extra = {
          id: extra?.id || extra?.ID,
          Semana_id: semanaId,
          Dia: dia,
          Nome: nome,
          tipo,
          Opcao: opcao,
          Observacao: observacao,
          Centro_Custo: cc
        };
      }

      const extraId = String(this.pick(extra, "id", "ID") || "");
      let pedidos = [];
      try {
        pedidos = (await this.getItems("Pedidos", { force: true }) || [])
          .filter(i => String(this.pick(i, "Semana_id", "Semana") || "") === String(semanaId));
      } catch (_) {
        pedidos = [];
      }

      const pedidoAtivo = (pedidos || []).find(p =>
        this._pedidoExtraAtivoEquivalente(p, semanaId, dia, nome, tipo, opcao, extraId)
      );

      if (pedidoAtivo?.id) {
        return { extra, pedido: pedidoAtivo, criadoPedido: false, reaproveitado: true };
      }

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
    })();

    this._extraPedidoLocks[lockKey] = run;
    try {
      return await run;
    } finally {
      delete this._extraPedidoLocks[lockKey];
    }
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
    const items = await this.getItems("Configurações", { ttl: this._CONFIG_CACHE_TTL_MS });
    const item = items.find(i =>
      this.pick(i, "Chave") === chave || this.pick(i, "Title") === chave
    );
    return item ? this.pick(item, "Valor") : null;
  },

  async setConfig(chave, valor, descricao = "") {
    const items = await this.getItems("Configurações", { force: true });
    const existing = items.find(i =>
      this.pick(i, "Chave") === chave || this.pick(i, "Title") === chave
    );
    const fields = { Valor: valor };
    if (descricao) fields.Descricao = descricao;
    let result;
    if (existing) result = await this.updateItem("Configurações", existing.id, fields);
    else result = await this.createItem("Configurações", { Title: chave, Chave: chave, Valor: valor, Descricao: descricao || "" });
    this.clearListCache("Configurações");
    this._emitSync("configuracoes", chave);
    return result;
  },

  addDias(date, dias) {
    const d = new Date(date || new Date());
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() + Number(dias || 0));
    return d;
  },

  getNextWeekId(date = new Date()) {
    return this.getSemanaId(this.addDias(date, 7));
  },

  async getSemanaAlvoMarcacao(fallback = null) {
    const alvo = await this.getConfig("marcacao_semana_alvo").catch(() => null);
    if (alvo && /^\d{4}-W\d{1,2}$/i.test(String(alvo).trim())) return String(alvo).trim().replace(/W(\d)$/i, "W0$1");

    const visivel = await this.getConfig("cardapio_semana_visivel").catch(() => null);
    if (visivel && /^\d{4}-W\d{1,2}$/i.test(String(visivel).trim())) return String(visivel).trim().replace(/W(\d)$/i, "W0$1");

    return fallback || this.getNextWeekId(new Date());
  },

  async setSemanaAlvoMarcacao(semanaId) {
    const semana = String(semanaId || this.getNextWeekId(new Date())).trim();
    await this.setConfig("marcacao_semana_alvo", semana, "Semana_id aberta para marcação pública dos colaboradores.");
    return semana;
  },

  async getSemanaVisivelCardapio(fallback = null) {
    const semana = await this.getConfig("cardapio_semana_visivel").catch(() => null);
    if (semana && /^\d{4}-W\d{1,2}$/i.test(String(semana).trim())) return String(semana).trim().replace(/W(\d)$/i, "W0$1");
    return this.getSemanaAlvoMarcacao(fallback);
  },

  async setSemanaVisivelCardapio(semanaId) {
    const semana = String(semanaId || await this.getSemanaAlvoMarcacao()).trim();
    await this.setConfig("cardapio_semana_visivel", semana, "Semana_id exibida no Cardápio da Semana público.");
    return semana;
  },

  async isCardapioLiberado() {
    return this.getMarcacaoLiberada();
  },

  async getMarcacaoLiberada() {
    const v = await this.getConfig("marcacao_liberada").catch(() => null);
    if (v !== null && v !== undefined && String(v).trim() !== "") return this.isTrue(v);

    // Compatibilidade com as chaves antigas já existentes no SharePoint.
    for (const chave of ["cardapio_liberado", "pedidos_liberados"]) {
      const legado = await this.getConfig(chave).catch(() => null);
      if (legado !== null && legado !== undefined && String(legado).trim() !== "") return this.isTrue(legado);
    }
    return false;
  },

  async setMarcacaoLiberada(liberado, semanaId = null) {
    const semana = semanaId ? await this.setSemanaAlvoMarcacao(semanaId) : await this.getSemanaAlvoMarcacao();
    const v = liberado ? "sim" : "nao";
    await this.setConfig("marcacao_liberada", v, "Libera ou bloqueia a marcação da semana alvo.");
    await this.setConfig("pedidos_liberados", v, "Compatibilidade: libera ou bloqueia pedidos públicos.");
    await this.setConfig("cardapio_liberado", v, "Compatibilidade: flag antiga de liberação.");
    if (liberado) await this.setSemanaVisivelCardapio(semana);
    return true;
  },

  async getCardapioVisivel() {
    const v = await this.getConfig("cardapio_visivel").catch(() => null);
    if (v !== null && v !== undefined && String(v).trim() !== "") return this.isTrue(v);

    const legado = await this.getConfig("cardapio_liberado").catch(() => null);
    return this.isTrue(legado);
  },

  async setCardapioVisivel(visivel, semanaId = null) {
    const v = visivel ? "sim" : "nao";
    await this.setConfig("cardapio_visivel", v, "Libera ou bloqueia a visualização pública do cardápio.");
    await this.setConfig("cardapio_liberado", v, "Compatibilidade: flag antiga de cardápio liberado.");
    if (semanaId) await this.setSemanaVisivelCardapio(semanaId);
    else if (visivel) await this.setSemanaVisivelCardapio(await this.getSemanaAlvoMarcacao());
    return true;
  },

  async getPrazoMarcacao(semanaId = null) {
    const alvo = semanaId || await this.getSemanaAlvoMarcacao().catch(() => null);
    if (alvo) {
      const especifico = await this.getConfig(`marcacao_prazo_limite_${alvo}`).catch(() => null);
      if (especifico) return especifico;
    }
    const novo = await this.getConfig("marcacao_prazo_limite").catch(() => null);
    if (novo) return novo;
    return this.getConfig("prazo_limite");
  },

  async setPrazoMarcacao(valor, semanaId = null) {
    const alvo = semanaId || await this.getSemanaAlvoMarcacao().catch(() => null);
    await this.setConfig("marcacao_prazo_limite", valor, "Prazo limite da semana alvo de marcação.");
    await this.setConfig("prazo_limite", valor, "Compatibilidade: prazo antigo usado pelas telas legadas.");
    if (alvo) await this.setConfig(`marcacao_prazo_limite_${alvo}`, valor, `Prazo limite específico da semana ${alvo}.`);
    return true;
  },

  async abrirMarcacaoSemana(semanaId, prazoLimite = null) {
    const semana = await this.setSemanaAlvoMarcacao(semanaId || this.getNextWeekId(new Date()));
    await this.setSemanaVisivelCardapio(semana);
    await this.setMarcacaoLiberada(true, semana);
    await this.setCardapioVisivel(true, semana);
    if (prazoLimite) await this.setPrazoMarcacao(prazoLimite, semana);
    return this.getEstadoMarcacao(semana);
  },

  async fecharMarcacaoSemana(semanaId = null) {
    const semana = semanaId || await this.getSemanaAlvoMarcacao().catch(() => null);
    await this.setMarcacaoLiberada(false, semana);
    return this.getEstadoMarcacao(semana);
  },

  async getEstadoMarcacao(semanaId = null) {
    const alvo = semanaId || await this.getSemanaAlvoMarcacao().catch(() => null);
    const liberada = await this.getMarcacaoLiberada().catch(() => false);
    const cardapioVisivel = await this.getCardapioVisivel().catch(() => false);
    const prazoLimite = await this.getPrazoMarcacao(alvo).catch(() => null);
    const prazo = prazoLimite ? new Date(prazoLimite) : null;
    const prazoVencido = !!(prazo && !isNaN(prazo) && new Date() > prazo);
    return { semanaId: alvo, liberada, cardapioVisivel, prazoLimite, prazoVencido };
  },

  async isMarcacaoLiberada() {
    const estado = await this.getEstadoMarcacao().catch(() => null);
    return !!(estado && estado.liberada && !estado.prazoVencido);
  },

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

  _dataISOAusencia(v) {
    if (!v) return "";
    if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
    const s = String(v || "").trim();
    const mIso = s.match(/^(\d{4}-\d{2}-\d{2})/);
    if (mIso) return mIso[1];
    const mBr = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (mBr) return `${mBr[3]}-${mBr[2]}-${mBr[1]}`;
    const d = new Date(s);
    return isNaN(d) ? "" : d.toISOString().slice(0, 10);
  },

  _hojeISO() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const dia = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${dia}`;
  },

  _ausenciaFimISO(a) {
    return this._dataISOAusencia(this.pick(a, "Data_Fim", "Fim", "DataFim", "Data")) ||
           this._dataISOAusencia(this.pick(a, "Data_Inicio", "Inicio", "DataInicio", "Data"));
  },

  _ausenciaAtivaPorCampo(a) {
    const ativo = this.pick(a, "Ativo", "ativo");
    const status = this.norm(this.pick(a, "Status", "status", "Status_Ausencia", "statusAusencia") || "");
    if ([
      "inativo", "cancelado", "encerrado", "periodo encerrado", "período encerrado",
      "duplicado inativado", "duplicidade inativada", "false", "nao", "não", "0"
    ].includes(status)) return false;
    if (ativo === null || ativo === undefined || ativo === "") return true;
    return this.isTrue(ativo);
  },

  ausenciaPeriodoEncerrado(a, dataRef = null) {
    const ref = dataRef || this._hojeISO();
    const fim = this._ausenciaFimISO(a);
    return !!fim && fim < ref;
  },

  async sincronizarAusenciasEncerradas(dataRef = null, options = {}) {
    // v10.4 — somente leitura.
    // Ausência encerrada não precisa receber Status/Status_Ausencia.
    // A validade é calculada por Data_Inicio <= data <= Data_Fim e Ativo.
    const hoje = dataRef || this._hojeISO();
    const lista = "Ausencias do Refeitorio";
    let items = [];
    try {
      items = await this.getItems(lista, { force: !!options.force });
    } catch (e) {
      console.warn("[SharePoint] Não foi possível verificar ausências encerradas.", e);
      return { verificadas: 0, encerradas: 0, somenteLeitura: true };
    }

    const encerradas = (items || []).filter(a =>
      this._ausenciaAtivaPorCampo(a) && this.ausenciaPeriodoEncerrado(a, hoje)
    );

    return {
      verificadas: (items || []).length,
      encerradas: encerradas.length,
      somenteLeitura: true,
      mensagem: "Ausências encerradas não são atualizadas automaticamente; o período é interpretado por data."
    };
  },

  async getAusencias(apenasAtivas = true) {
    // v10.4 — somente leitura.
    // Não chamar sincronizarAusenciasEncerradas() aqui. Carregar Dashboard/Operação/
    // Marcar Refeição não pode tentar atualizar a lista Ausencias do Refeitorio.
    const items = await this.getItems("Ausencias do Refeitorio", {
      force: false,
      ttl: this._ITEMS_CACHE_TTL_MS
    });
    if (!apenasAtivas) return items;
    const hoje = this._hojeISO();
    return items.filter(i => this._ausenciaAtivaPorCampo(i) && !this.ausenciaPeriodoEncerrado(i, hoje));
  },

  async getAusenciasColaborador(colaboradorId, dataRef = null) {
    // Usa a lista completa para permitir auditoria/consulta histórica por dataRef.
    const items = await this.getAusencias(false);
    return items.filter(i => {
      if (!this._ausenciaAtivaPorCampo(i)) return false;
      if (String(this.pick(i, "Colaborador_id")) !== String(colaboradorId)) return false;
      if (!dataRef) return !this.ausenciaPeriodoEncerrado(i, this._hojeISO());
      const ref = this._dataISOAusencia(dataRef);
      const ini = this._dataISOAusencia(this.pick(i, "Data_Inicio", "Inicio", "DataInicio", "Data"));
      const fim = this._dataISOAusencia(this.pick(i, "Data_Fim", "Fim", "DataFim", "Data")) || ini;
      return !!ref && !!ini && !!fim && ini <= ref && fim >= ref;
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


  // ============================================================
  // INTEGRIDADE — Ausências x Pedidos
  // ============================================================
  _diasSemanaOperacional(semanaId) {
    const nomes = ["segunda", "terca", "quarta", "quinta", "sexta"];
    const datas = this.getWeekDates(semanaId) || [];
    return nomes.map((dia, idx) => ({
      dia,
      data: datas[idx] ? this._dataISOAusencia(datas[idx]) : ""
    })).filter(d => d.data);
  },

  _semanaInicioFimISO(semanaId) {
    const dias = this._diasSemanaOperacional(semanaId);
    return {
      ini: dias[0]?.data || "",
      fim: dias[dias.length - 1]?.data || "",
      dias
    };
  },

  _colabIdValor(obj) {
    return String(this.pick(obj, "Colaborador_id", "ColaboradorId", "colaboradorId", "id", "ID", "Matricula", "Matrícula") || "").trim();
  },

  _colabNomeValor(obj) {
    return this.pick(obj, "Colaborador_nome", "Colaborador", "Nome", "Title") || "";
  },

  _colabKey(obj) {
    const id = this._colabIdValor(obj);
    if (id) return `id:${id}`;
    const nome = this.norm(this._colabNomeValor(obj));
    return nome ? `nome:${nome}` : "";
  },

  _colaboradorAtivo(c) {
    const ativo = this.pick(c, "Ativo", "ativo");
    if (ativo === null || ativo === undefined || String(ativo).trim() === "") return true;
    return this.isTrue(ativo);
  },

  _isPedidoAusenciaStatus(status) {
    const s = this.norm(status);
    return [
      "nao vai almocar", "não vai almoçar", "nao_vai_almocar",
      "ferias", "férias", "afastado", "atestado", "licenca", "licença",
      "banco horas", "banco_horas", "homy office", "homy_office", "ausente", "falta"
    ].includes(s);
  },

  _formatarMotivoAusenciaSistema(motivo) {
    const n = this.norm(motivo);
    if (n === "nao_vai_almocar" || n === "nao vai almocar" || n === "não vai almoçar" || n === "ausente") return "Não vai almoçar";
    if (n === "ferias" || n === "férias") return "Férias";
    if (n === "afastado" || n === "afastamento") return "Afastado";
    if (n === "atestado") return "Atestado";
    if (n === "licenca" || n === "licença") return "Licença";
    if (n === "banco_horas" || n === "banco horas") return "Banco de horas";
    if (n === "homy_office" || n === "homy office") return "Homy Office";
    if (n === "falta") return "Falta";
    return String(motivo || "Ausente");
  },

  _ausenciaInicioISO(a) {
    return this._dataISOAusencia(this.pick(a, "Data_Inicio", "Inicio", "DataInicio", "Data"));
  },

  _ausenciaFimISO(a) {
    return this._dataISOAusencia(this.pick(a, "Data_Fim", "Fim", "DataFim", "Data")) || this._ausenciaInicioISO(a);
  },

  _ausenciaCobreData(a, dataISO) {
    const ini = this._ausenciaInicioISO(a);
    const fim = this._ausenciaFimISO(a);
    return !!ini && !!fim && ini <= dataISO && fim >= dataISO;
  },

  _ausenciaStatusNorm(a) {
    return this.norm(this.pick(a, "Status", "status", "Status_Ausencia", "statusAusencia") || "");
  },

  _ausenciaFoiCanceladaOuDuplicada(a) {
    const status = this._ausenciaStatusNorm(a);
    return [
      "cancelado", "cancelada", "duplicado inativado", "duplicidade inativada",
      "duplicada inativada", "excluido", "excluído"
    ].includes(status);
  },

  _ausenciaPodeGerarHistorico(a) {
    if (!a || this._ausenciaFoiCanceladaOuDuplicada(a)) return false;
    const status = this._ausenciaStatusNorm(a);
    const ativo = this.pick(a, "Ativo", "ativo");
    // Período encerrado é apenas uma condição temporal/visual; ele continua
    // valendo para os dias em que a data da operação caiu dentro do período.
    if (status === "periodo encerrado" || status === "período encerrado") return true;
    if (["inativo", "false", "nao", "não", "0"].includes(status)) return false;
    if (ativo === null || ativo === undefined || ativo === "") return true;
    return this.isTrue(ativo);
  },

  _ausenciaConsideradaParaData(a, dataISO) {
    if (!this._ausenciaPodeGerarHistorico(a)) return false;
    return this._ausenciaCobreData(a, dataISO);
  },

  _ausenciaChaveDuplicada(a) {
    const colab = this._colabKey({
      Colaborador_id: this.pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId"),
      Colaborador_nome: this.pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title")
    });
    const ini = this._ausenciaInicioISO(a);
    const motivo = this.norm(this.pick(a, "Motivo", "motivo", "Status") || "");
    return colab && ini && motivo ? `${colab}|${ini}|${motivo}` : "";
  },

  _ausenciaPreferidaParaGrupo(grupo) {
    return [...grupo].sort((a, b) => {
      const aa = this._ausenciaAtivaPorCampo(a) ? 1 : 0;
      const bb = this._ausenciaAtivaPorCampo(b) ? 1 : 0;
      if (aa !== bb) return bb - aa;
      const fa = this._ausenciaFimISO(a);
      const fb = this._ausenciaFimISO(b);
      if (fa !== fb) return fb.localeCompare(fa);
      return Number(this.pick(b, "id", "ID") || 0) - Number(this.pick(a, "id", "ID") || 0);
    })[0];
  },

  _ausenciaIntervalosSobrepostos(a, b) {
    const ai = this._ausenciaInicioISO(a);
    const af = this._ausenciaFimISO(a);
    const bi = this._ausenciaInicioISO(b);
    const bf = this._ausenciaFimISO(b);
    return !!ai && !!af && !!bi && !!bf && ai <= bf && af >= bi;
  },

  _ausenciaPreferidaParaSobreposicao(grupo) {
    return [...grupo].sort((a, b) => {
      const aa = this._ausenciaAtivaPorCampo(a) ? 1 : 0;
      const bb = this._ausenciaAtivaPorCampo(b) ? 1 : 0;
      if (aa !== bb) return bb - aa;
      const ma = new Date(this.pick(a, "Modified", "modified", "Data_Hora", "Created", "created") || 0).getTime() || 0;
      const mb = new Date(this.pick(b, "Modified", "modified", "Data_Hora", "Created", "created") || 0).getTime() || 0;
      if (ma !== mb) return mb - ma;
      const fa = this._ausenciaFimISO(a);
      const fb = this._ausenciaFimISO(b);
      if (fa !== fb) return fb.localeCompare(fa);
      return Number(this.pick(b, "id", "ID") || 0) - Number(this.pick(a, "id", "ID") || 0);
    })[0];
  },

  async sincronizarAusenciasDuplicadas() {
    let ausencias = [];
    try { ausencias = await this.getItems("Ausencias do Refeitorio", { force: true }); }
    catch (e) {
      console.warn("[SharePoint] Não foi possível verificar duplicidade de ausências.", e);
      return { verificadas: 0, duplicadasInativadas: 0 };
    }

    const porColabMotivo = new Map();
    for (const a of ausencias || []) {
      const colab = this._colabKey({
        Colaborador_id: this.pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId"),
        Colaborador_nome: this.pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title")
      });
      const motivo = this.norm(this.pick(a, "Motivo", "motivo", "Status") || "");
      const ini = this._ausenciaInicioISO(a);
      const fim = this._ausenciaFimISO(a);
      if (!colab || !motivo || !ini || !fim) continue;
      const k = `${colab}|${motivo}`;
      if (!porColabMotivo.has(k)) porColabMotivo.set(k, []);
      porColabMotivo.get(k).push(a);
    }

    let duplicadasInativadas = 0;

    for (const grupoBase of porColabMotivo.values()) {
      const ativos = grupoBase.filter(a => this._ausenciaAtivaPorCampo(a));
      if (ativos.length < 2) continue;

      const visitados = new Set();
      for (const item of ativos) {
        const idItem = String(this.pick(item, "id", "ID") || "");
        if (!idItem || visitados.has(idItem)) continue;

        const cluster = ativos.filter(outro => this._ausenciaIntervalosSobrepostos(item, outro));
        cluster.forEach(x => visitados.add(String(this.pick(x, "id", "ID") || "")));
        if (cluster.length < 2) continue;

        const preferida = this._ausenciaPreferidaParaSobreposicao(cluster);
        const idPreferida = String(this.pick(preferida, "id", "ID") || "");

        for (const a of cluster) {
          const id = String(this.pick(a, "id", "ID") || "");
          if (!id || id === idPreferida) continue;
          const obsAtual = this.pick(a, "Observacao", "Observação", "Obs") || "";
          const obs = [obsAtual, `Duplicidade/sobreposição inativada automaticamente; ausência mantida: ${idPreferida}.`].filter(Boolean).join(" | ");
          try {
            await this.updateItem("Ausencias do Refeitorio", id, {
              Ativo: false,
              Observacao: obs
            });
            duplicadasInativadas++;
          } catch (e) {
            console.warn(`[SharePoint] Falha ao inativar ausência duplicada/sobreposta ${id}.`, e);
          }
        }
      }
    }

    return { verificadas: (ausencias || []).length, duplicadasInativadas };
  },

  _pedidoNormalDoColaboradorNoDia(p, semanaId, colabKey, dia) {
    if (!p || this._isPedidoAdicionalColaborador?.(p) || this.isExtraPedido?.(p)) return false;
    if (String(this.pick(p, "Semana_id", "Semana") || "") !== String(semanaId || "")) return false;
    if (this.norm(this.pick(p, "Dia", "dia") || "") !== this.norm(dia)) return false;
    return this._colabKey({
      Colaborador_id: this.pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId"),
      Colaborador_nome: this.pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title")
    }) === colabKey;
  },

  _pedidoAusenciaParaAtualizar(p) {
    const status = this.pick(p, "Status", "status") || "";
    const origem = this.norm(this.pick(p, "Origem", "origem", "tipo", "Tipo") || "");
    return this._isPedidoAusenciaStatus(status) || origem.includes("ausencia") || origem.includes("ausência");
  },

  async _criarOuAtualizarPedidoAusencia(semanaId, diaInfo, ausencia, pedidos) {
    const colabId = String(this.pick(ausencia, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    const nome = this.pick(ausencia, "Colaborador_nome", "Colaborador", "Nome", "Title") || "Colaborador ausente";
    const colabKey = this._colabKey({ Colaborador_id: colabId, Colaborador_nome: nome });
    if (!colabKey) return { criado: 0, atualizado: 0 };

    const motivo = this._formatarMotivoAusenciaSistema(this.pick(ausencia, "Motivo", "motivo", "Status") || "Ausente");
    const cc = this.pick(ausencia, "Centro_Custo", "CentroCusto", "Setor", "Departamento") || await this._resolverCentroCustoColaborador(colabId, nome);
    const existente = (pedidos || []).find(p => this._pedidoNormalDoColaboradorNoDia(p, semanaId, colabKey, diaInfo.dia));
    const dataHora = `${diaInfo.data}T12:00:00`;
    const obsBase = this.pick(ausencia, "Observacao", "Observação", "Obs") || "Ausência cadastrada no Admin.";
    const obs = [obsBase, `AusenciaID:${this.pick(ausencia, "id", "ID") || ""}`].filter(Boolean).join(" | ");

    const fields = {
      Semana_id: semanaId,
      Colaborador_id: colabId,
      Colaborador_nome: nome,
      Dia: diaInfo.dia,
      Opcao: "principal",
      Nome_Prato: motivo,
      Confirmado: false,
      Data_Hora: dataHora,
      Centro_Custo: cc || "",
      Status: motivo,
      Observacao: obs,
      Origem: "Ausência",
      Alterado_Por: this.getUserName ? this.getUserName() : "Sistema"
    };

    if (existente?.id) {
      const deveAtualizar = this._pedidoAusenciaParaAtualizar(existente) ||
        this.norm(this.pick(existente, "Status")) !== this.norm(motivo) ||
        this.norm(this.pick(existente, "Nome_Prato")) !== this.norm(motivo);
      if (deveAtualizar) {
        await this.updatePedido(existente.id, fields);
        Object.assign(existente, fields);
        return { criado: 0, atualizado: 1 };
      }
      return { criado: 0, atualizado: 0 };
    }

    const criado = await this.createItem("Pedidos", {
      Title: `${semanaId}-${colabId || this.norm(nome)}-${diaInfo.dia}-ausencia`,
      ...fields
    });
    pedidos.push({ id: criado?.id || criado?.ID || "", ...fields });
    return { criado: 1, atualizado: 0 };
  },

  async _criarOuAtualizarPedidoRetornoPrincipal(semanaId, diaInfo, ausencia, colaborador, pedidos, options = {}) {
    // v10.4 — retorno de ausência nunca cria Principal automaticamente.
    // O fim da ausência apenas libera marcação. Se ninguém marcar até o prazo,
    // Principal só é gerado por travamento/fechamento explícito.
    return {
      criado: 0,
      atualizado: 0,
      ignorado: 1,
      somenteLeitura: true,
      motivo: "Retorno de ausência não gera Principal automático; aguarda marcação ou travamento."
    };
  },




  _pedidoColaboradorNormalKey(p) {
    if (!p || this._isPedidoAdicionalColaborador?.(p) || this.isExtraPedido?.(p)) return "";
    const id = String(this.pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    if (id) return `id:${id}`;
    const nome = this.norm(this.pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title") || "");
    return nome ? `nome:${nome}` : "";
  },

  _pedidoStatusAusenciaOuOrigem(p) {
    const status = this.pick(p, "Status", "status") || "";
    const origem = this.norm(this.pick(p, "Origem", "origem", "tipo", "Tipo") || "");
    return this._isPedidoAusenciaStatus(status) || origem.includes("ausencia") || origem.includes("ausência");
  },

  _ausenciaVigenteParaKeyData(ausencias, colabKey, dataISO) {
    const candidatas = (ausencias || []).filter(a => {
      const key = this._colabKey({
        Colaborador_id: this.pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId"),
        Colaborador_nome: this.pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title")
      });
      if (key !== colabKey) return false;
      return this._ausenciaConsideradaParaData(a, dataISO);
    });
    if (!candidatas.length) return null;
    return this._ausenciaPreferidaParaSobreposicao(candidatas);
  },

  _pedidoPreferidoGrupoColaborador(grupo) {
    // v7 — Pedido produtivo válido sempre vence ausência/cancelado/obsoleto.
    return [...(grupo || [])].sort((a, b) => {
      const ap = this._pedidoProdutivoValido(a) ? 1 : 0;
      const bp = this._pedidoProdutivoValido(b) ? 1 : 0;
      if (ap !== bp) return bp - ap;

      const aa = this._pedidoStatusAusenciaOuOrigem(a) ? 1 : 0;
      const bb = this._pedidoStatusAusenciaOuOrigem(b) ? 1 : 0;
      if (aa !== bb) return aa - bb;

      const ca = this.norm(this.pick(a, "Status", "status")) === "cancelado" ? 1 : 0;
      const cb = this.norm(this.pick(b, "Status", "status")) === "cancelado" ? 1 : 0;
      if (ca !== cb) return ca - cb;

      const ta = new Date(this.pick(a, "Modified", "modified", "Data_Hora", "Created", "created") || 0).getTime() || Number(this.pick(a, "id", "ID") || 0) || 0;
      const tb = new Date(this.pick(b, "Modified", "modified", "Data_Hora", "Created", "created") || 0).getTime() || Number(this.pick(b, "id", "ID") || 0) || 0;
      return tb - ta;
    })[0];
  },


  async _cancelarPedidoDuplicadoAusencia(id, motivo = "Pedido duplicado/obsoleto inativado automaticamente.") {
    if (!id) return 0;
    try {
      await this.updatePedido(id, {
        Status: "Cancelado",
        Confirmado: false,
        Origem: "Duplicado inativado",
        Observacao: motivo,
        Alterado_Por: this.getUserName ? this.getUserName() : "Sistema"
      });
      return 1;
    } catch (e) {
      console.warn(`[SharePoint] Falha ao cancelar pedido duplicado/obsoleto ${id}.`, e);
      return 0;
    }
  },

  async _normalizarPedidosPorAusenciasSemana(semanaId, pedidos, ausencias, dias, colabPorKey) {
    let pedidosAtualizados = 0;
    let pedidosCancelados = 0;
    let retornosIgnoradosPorDiaPassado = 0;
    const porColabDia = new Map();

    for (const p of pedidos || []) {
      if (!p || this.isExtraPedido?.(p) || this._isPedidoAdicionalColaborador?.(p)) continue;
      if (String(this.pick(p, "Semana_id", "Semana") || "") !== String(semanaId || "")) continue;
      const diaNorm = this.norm(this.pick(p, "Dia", "dia") || "");
      const colabKey = this._pedidoColaboradorNormalKey(p);
      if (!diaNorm || !colabKey) continue;
      const k = `${colabKey}|${diaNorm}`;
      if (!porColabDia.has(k)) porColabDia.set(k, []);
      porColabDia.get(k).push(p);
    }

    const diaPorNorm = new Map((dias || []).map(d => [this.norm(d.dia), d]));

    for (const [grupoKey, grupo] of porColabDia.entries()) {
      const [colabKey, diaNorm] = grupoKey.split("|");
      const diaInfo = diaPorNorm.get(diaNorm);
      if (!diaInfo) continue;
      const diaPassado = this._diaOperacionalPassado(diaInfo);
      const ausenciaVigente = this._ausenciaVigenteParaKeyData(ausencias, colabKey, diaInfo.data);

      if (ausenciaVigente) {
        const motivo = this._formatarMotivoAusenciaSistema(this.pick(ausenciaVigente, "Motivo", "motivo", "Status") || "Ausente");
        const preferido = grupo.find(p => this._pedidoStatusAusenciaOuOrigem(p)) || this._pedidoPreferidoGrupoColaborador(grupo);
        const idPref = String(this.pick(preferido, "id", "ID") || "");
        if (!idPref) continue;

        const nome = this.pick(ausenciaVigente, "Colaborador_nome", "Colaborador", "Nome", "Title") || this.pick(preferido, "Colaborador_nome", "Colaborador", "Nome", "Title") || "Colaborador ausente";
        const colabId = String(this.pick(ausenciaVigente, "Colaborador_id", "ColaboradorId", "colaboradorId") || this.pick(preferido, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
        const cc = this.pick(ausenciaVigente, "Centro_Custo", "CentroCusto", "Setor", "Departamento") || this.pick(preferido, "Centro_Custo", "CentroCusto", "Setor", "Departamento") || await this._resolverCentroCustoColaborador(colabId, nome);
        const obsBase = this.pick(ausenciaVigente, "Observacao", "Observação", "Obs") || "Ausência cadastrada no Admin.";
        const fieldsAus = {
          Semana_id: semanaId,
          Colaborador_id: colabId,
          Colaborador_nome: nome,
          Dia: diaInfo.dia,
          Opcao: "principal",
          Nome_Prato: motivo,
          Confirmado: false,
          Data_Hora: `${diaInfo.data}T12:00:00`,
          Centro_Custo: cc || "",
          Status: motivo,
          Observacao: [obsBase, `AusenciaID:${this.pick(ausenciaVigente, "id", "ID") || ""}`].filter(Boolean).join(" | "),
          Origem: "Ausência",
          Alterado_Por: this.getUserName ? this.getUserName() : "Sistema"
        };
        await this.updatePedido(idPref, fieldsAus);
        Object.assign(preferido, fieldsAus);
        pedidosAtualizados++;

        for (const p of grupo) {
          const pid = String(this.pick(p, "id", "ID") || "");
          if (!pid || pid === idPref) continue;
          pedidosCancelados += await this._cancelarPedidoDuplicadoAusencia(pid, `Pedido duplicado/substituído por ausência vigente (${motivo}); pedido mantido: ${idPref}.`);
        }
      } else {
        const produtivos = grupo.filter(p => this._pedidoProdutivoValido(p));
        const normais = grupo.filter(p => !this._pedidoStatusAusenciaOuOrigem(p) && !["cancelado", "bloqueado"].includes(this.norm(this.pick(p, "Status", "status"))));
        const preferido = produtivos.length
          ? this._pedidoPreferidoGrupoColaborador(produtivos)
          : (normais.length ? this._pedidoPreferidoGrupoColaborador(normais) : this._pedidoPreferidoGrupoColaborador(grupo));

        const idPref = String(this.pick(preferido, "id", "ID") || "");
        if (!idPref) continue;

        if (this._pedidoStatusAusenciaOuOrigem(preferido)) {
          // v10.4 — ausência encerrada não vira Principal aqui.
          // Mantém o registro como está para histórico/auditoria. O colaborador fica pendente
          // na Operação/Marcar Refeição até escolher ou até o travamento/fechamento explícito.
          retornosIgnoradosPorDiaPassado++;
          continue;
        }

        if (diaPassado) continue;

        for (const p of grupo) {
          const pid = String(this.pick(p, "id", "ID") || "");
          if (!pid || pid === idPref) continue;
          if (this._pedidoProdutivoValido(p)) continue;
          if (this._pedidoStatusAusenciaOuOrigem(p) || this.norm(this.pick(p, "Status", "status")) === "cancelado") {
            pedidosCancelados += await this._cancelarPedidoDuplicadoAusencia(pid, `Pedido de ausência/duplicado obsoleto; pedido mantido: ${idPref}.`);
          }
        }
      }
    }

    return { pedidosAtualizados, pedidosCancelados, retornosIgnoradosPorDiaPassado };
  },


  async sincronizarAusenciasPedidosSemana(semanaId, pedidosBase = null) {
    if (!semanaId) return { ausenciasCriadas: 0, ausenciasAtualizadas: 0, retornoCriados: 0, retornoAtualizados: 0, duplicadasInativadas: 0, pedidosAtualizados: 0, pedidosCancelados: 0 };
    if (this._sincronizandoAusenciasPedidos) return { ausenciasCriadas: 0, ausenciasAtualizadas: 0, retornoCriados: 0, retornoAtualizados: 0, duplicadasInativadas: 0, pedidosAtualizados: 0, pedidosCancelados: 0 };

    this._sincronizandoAusenciasPedidos = true;
    let ausenciasCriadas = 0, ausenciasAtualizadas = 0, retornoCriados = 0, retornoAtualizados = 0, duplicadasInativadas = 0, pedidosAtualizados = 0, pedidosCancelados = 0;

    try {
      const dup = await this.sincronizarAusenciasDuplicadas().catch(e => {
        console.warn("[SharePoint] Reparação de duplicidade de ausências ignorada.", e);
        return { duplicadasInativadas: 0 };
      });
      duplicadasInativadas = dup?.duplicadasInativadas || 0;

      // v10.4 — não atualizar ausências encerradas durante sincronização.
      // Períodos vencidos são interpretados por Data_Inicio/Data_Fim.

      const { ini: semanaIni, fim: semanaFim, dias } = this._semanaInicioFimISO(semanaId);
      if (!semanaIni || !semanaFim) return { ausenciasCriadas, ausenciasAtualizadas, retornoCriados, retornoAtualizados, duplicadasInativadas, pedidosAtualizados, pedidosCancelados };

      const [ausenciasTodas, colaboradores, pedidosIniciais] = await Promise.all([
        this.getItems("Ausencias do Refeitorio").catch(() => []),
        this.getTodosColaboradores().catch(() => []),
        pedidosBase ? Promise.resolve(pedidosBase) : this.getItems("Pedidos").then(items => items.filter(i => this.pick(i, "Semana_id") === semanaId)).catch(() => [])
      ]);
      const pedidos = Array.isArray(pedidosIniciais) ? [...pedidosIniciais] : [];
      const colabsAtivos = (colaboradores || []).filter(c => this._colaboradorAtivo(c));
      const colabPorKey = new Map(colabsAtivos.map(c => [this._colabKey(c), c]).filter(([k]) => k));

      const ausenciasOrdenadas = (ausenciasTodas || []).filter(a => {
        const ini = this._ausenciaInicioISO(a);
        const fim = this._ausenciaFimISO(a);
        return !!ini && !!fim && ini <= semanaFim && fim >= semanaIni;
      });

      const ausenciasOperacionais = ausenciasOrdenadas.filter(a => this._ausenciaPodeGerarHistorico(a));

      for (const a of ausenciasOperacionais) {
        for (const diaInfo of dias) {
          if (!this._ausenciaConsideradaParaData(a, diaInfo.data)) continue;
          const r = await this._criarOuAtualizarPedidoAusencia(semanaId, diaInfo, a, pedidos);
          ausenciasCriadas += r.criado || 0;
          ausenciasAtualizadas += r.atualizado || 0;
        }
      }

      // v10.4 — retorno de ausência NÃO cria Principal automaticamente.
      // O fim da ausência apenas libera o colaborador para escolher.
      // Se ninguém marcar até o prazo, o Principal é gerado somente pelo travamento/fechamento explícito.

      const normPedidos = await this._normalizarPedidosPorAusenciasSemana(semanaId, pedidos, ausenciasOrdenadas, dias, colabPorKey)
        .catch(e => {
          console.warn("[SharePoint] Normalização de pedidos por ausência ignorada.", e);
          return { pedidosAtualizados: 0, pedidosCancelados: 0 };
        });
      pedidosAtualizados += normPedidos?.pedidosAtualizados || 0;
      pedidosCancelados  += normPedidos?.pedidosCancelados  || 0;

    } finally {
      this._sincronizandoAusenciasPedidos = false;
    }

    return { ausenciasCriadas, ausenciasAtualizadas, retornoCriados, retornoAtualizados, duplicadasInativadas, pedidosAtualizados, pedidosCancelados };
  },

  async _buscarAusenciaDuplicadaParaUpsert(colaboradorId, colaboradorNome, motivo, dataInicio, dataFim) {
    let ausencias = [];
    try { ausencias = await this.getItems("Ausencias do Refeitorio"); }
    catch (e) { console.warn("[SharePoint] Não foi possível verificar ausência duplicada antes de salvar.", e); return null; }

    const alvoKey = this._colabKey({ Colaborador_id: colaboradorId, Colaborador_nome: colaboradorNome });
    const motivoNorm = this.norm(motivo);
    const iniNovo = this._dataISOAusencia(dataInicio);
    const fimNovo = this._dataISOAusencia(dataFim) || iniNovo;
    const candidatos = (ausencias || []).filter(a => {
      if (!this._ausenciaAtivaPorCampo(a)) return false;
      const key = this._colabKey({
        Colaborador_id: this.pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId"),
        Colaborador_nome: this.pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title")
      });
      if (key !== alvoKey) return false;
      if (this.norm(this.pick(a, "Motivo", "motivo", "Status") || "") !== motivoNorm) return false;
      const ini = this._ausenciaInicioISO(a);
      const fim = this._ausenciaFimISO(a);
      return !!ini && !!fim && ini <= fimNovo && fim >= iniNovo;
    });

    if (!candidatos.length) return null;
    return this._ausenciaPreferidaParaGrupo(candidatos);
  },

  async createAusencia(dados) {
    const nome   = dados.colaboradorNome || dados.Colaborador_nome || "";
    const motivo = dados.motivo || dados.Motivo || "nao_vai_almocar";
    const colaboradorId = String(dados.colaboradorId || dados.Colaborador_id || "");
    let centroCusto = dados.centroCusto || dados.Centro_Custo || dados.centro_custo || "";

    if (!centroCusto) {
      centroCusto = await this._resolverCentroCustoColaborador(colaboradorId, nome);
    }

    const dataInicio = dados.dataInicio || dados.Data_Inicio;
    const dataFim = dados.dataFim || dados.Data_Fim || dataInicio;
    const semanaInicio = dataInicio ? this.getSemanaId(new Date(`${String(dataInicio).slice(0, 10)}T12:00:00`)) : "";
    const semanaFim = dataFim ? this.getSemanaId(new Date(`${String(dataFim).slice(0, 10)}T12:00:00`)) : semanaInicio;

    const fields = {
      Title:            dados.title || dados.Title || `${nome} - ${motivo}`,
      Colaborador_id:   colaboradorId,
      Colaborador_nome: nome,
      Centro_Custo:     centroCusto || "",
      Data_Inicio:      dataInicio,
      Data_Fim:         dataFim,
      Semana_id:        dados.semanaId || dados.Semana_id || semanaInicio,
      Semana_Inicio:    dados.semanaInicio || dados.Semana_Inicio || semanaInicio,
      Semana_Fim:       dados.semanaFim || dados.Semana_Fim || semanaFim,
      Motivo:           motivo,
      Observacao:       dados.observacao  || dados.Observacao || "",
      Ativo:            dados.ativo       ?? dados.Ativo      ?? true,
      Criado_Por:       dados.criadoPor   || dados.Criado_Por || this.getUserName()
    };

    const existente = await this._buscarAusenciaDuplicadaParaUpsert(colaboradorId, nome, motivo, dataInicio, dataFim);
    let result;
    if (existente?.id) {
      result = await this.updateItem("Ausencias do Refeitorio", existente.id, fields);
      result = { id: existente.id, ...fields };
    } else {
      result = await this.createItem("Ausencias do Refeitorio", fields);
    }

    if (!this._sincronizandoAusenciasPedidos) {
      const semanas = new Set([semanaInicio, semanaFim].filter(Boolean));
      for (const semana of semanas) {
        await this.sincronizarAusenciasPedidosSemana(semana).catch(e => console.warn("[SharePoint] Sincronização após salvar ausência ignorada:", e));
      }
    }

    return result;
  },

  async updateAusencia(id, dados) {
    let antes = null;
    try { antes = await this._getItemFieldsById("Ausencias do Refeitorio", id); }
    catch (e) { console.warn("[SharePoint] Não foi possível ler ausência antes de atualizar.", e); }

    const fields = {};
    if (dados.ativo       !== undefined) fields.Ativo     = dados.ativo;
    if (dados.Ativo       !== undefined) fields.Ativo     = dados.Ativo;
    if (dados.motivo      !== undefined) fields.Motivo    = dados.motivo;
    if (dados.Motivo      !== undefined) fields.Motivo    = dados.Motivo;
    if (dados.observacao  !== undefined) fields.Observacao = dados.observacao;
    if (dados.Observacao  !== undefined) fields.Observacao = dados.Observacao;
    if (dados.dataInicio  !== undefined) fields.Data_Inicio = dados.dataInicio;
    if (dados.Data_Inicio !== undefined) fields.Data_Inicio = dados.Data_Inicio;
    if (dados.dataFim     !== undefined) fields.Data_Fim    = dados.dataFim;
    if (dados.Data_Fim    !== undefined) fields.Data_Fim    = dados.Data_Fim;
    if (dados.Centro_Custo !== undefined) fields.Centro_Custo = dados.Centro_Custo;
    if (dados.centroCusto !== undefined) fields.Centro_Custo = dados.centroCusto;

    const result = await this.updateItem("Ausencias do Refeitorio", id, fields);

    if (!this._sincronizandoAusenciasPedidos) {
      const depois = { ...(antes || {}), ...fields };
      const semanas = new Set();
      const addSemanas = item => {
        const ini = this._ausenciaInicioISO(item);
        const fim = this._ausenciaFimISO(item) || ini;
        if (ini) semanas.add(this.getSemanaId(new Date(`${ini}T12:00:00`)));
        if (fim) semanas.add(this.getSemanaId(new Date(`${fim}T12:00:00`)));
      };
      addSemanas(antes || {});
      addSemanas(depois || {});
      semanas.add(this.getCurrentWeekId ? this.getCurrentWeekId() : this.getSemanaId(new Date()));
      for (const semana of [...semanas].filter(Boolean)) {
        this._repairCache[semana] = 0;
        await this.sincronizarAusenciasPedidosSemana(semana).catch(e => console.warn("[SharePoint] Sincronização após atualizar ausência ignorada:", e));
      }
    }

    return result;
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
    if (this.ausenciaPeriodoEncerrado(a)) return false;
    return this._ausenciaAtivaPorCampo(a);
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
      if (this._isPedidoAdicionalColaborador(p)) continue;
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

  _slugAdicional(v) {
    return this.norm(v || "")
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "") || "colaborador";
  },

  async _buscarPedidoNormalColaboradorDia(semanaId, colaboradorId, colaboradorNome, dia) {
    let pedidos = [];
    try { pedidos = await this.getPedidos(semanaId); }
    catch (e) { console.warn("[SharePoint] Não foi possível buscar pedido normal para transferência:", e); return null; }

    const candidatos = (pedidos || []).filter(p =>
      !this._isPedidoAdicionalColaborador(p) &&
      this._pedidoMesmoColaboradorDia(p, semanaId, colaboradorId, colaboradorNome, dia)
    );
    if (!candidatos.length) return null;
    candidatos.sort((a, b) => this._ordenarPedidoMaisAtual(a, b));
    return candidatos[0];
  },

  async criarRefeicaoAdicionalColaborador(params = {}) {
    const semanaId = params.semanaId || params.Semana_id || this.getCurrentWeekId();
    const pedidoBase = params.pedidoBase || params.pedido || {};
    const tipoAjuste = params.tipoAjuste || params.tipo || "segunda";
    const diaDestino = params.diaDestino || params.dia || this.pick(pedidoBase, "Dia", "dia");
    const opcaoDestino = params.opcaoDestino || params.opcao || "principal";
    const diaOrigem = params.diaOrigem || params.diaRemover || "";
    const observacaoUsuario = params.observacao || params.Observacao || "";

    const pedidoIdBase = String(params.pedidoId || this.pick(pedidoBase, "id", "ID") || "");
    let base = pedidoBase;
    if (pedidoIdBase && !this.pick(base, "Colaborador_nome", "Nome", "Title")) {
      try { base = await this._getItemFieldsById("Pedidos", pedidoIdBase); }
      catch (e) { console.warn("[SharePoint] Não foi possível ler pedido base do adicional:", e); }
    }

    const colaboradorIdBase = String(params.colaboradorId || this.pick(base, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    const colaboradorNome = params.colaboradorNome || this.pick(base, "Colaborador_nome", "Colaborador", "Nome", "Title") || "";
    if (!colaboradorNome) throw new Error("Não foi possível identificar o colaborador para criar refeição adicional.");

    let centroCusto = params.centroCusto || this.pick(base, "Centro_Custo", "CentroCusto", "Setor", "Departamento") || "";
    if (!centroCusto) centroCusto = await this._resolverCentroCustoColaborador(colaboradorIdBase, colaboradorNome);

    const dataDestino = this._pedidoDataRefeicao(semanaId, diaDestino, base);
    const nomePratoDestino = await this._nomePratoCardapioPorOpcao(semanaId, diaDestino, opcaoDestino) || this._pratoPadraoPorOpcao(opcaoDestino);
    const ts = Date.now();
    const baseSlug = colaboradorIdBase || this._slugAdicional(colaboradorNome);
    const idAdicional = `${baseSlug}-adicional-${this._slugAdicional(diaDestino)}-${ts}`;
    const origemAdicional = this.norm(tipoAjuste).includes("transfer")
      ? "Transferência para hoje"
      : "Segunda refeição";

    const obsAdicional = [
      "Refeição adicional vinculada ao colaborador.",
      observacaoUsuario,
      pedidoIdBase ? `PedidoBase:${pedidoIdBase}` : "",
      colaboradorIdBase ? `ColaboradorBaseID:${colaboradorIdBase}` : "",
      `AdicionalID:${ts}`
    ].filter(Boolean).join(" | ");

    const adicional = await this.createItem("Pedidos", {
      Title:            `${semanaId}-${idAdicional}-${diaDestino}`,
      Semana_id:        semanaId,
      Colaborador_id:   String(idAdicional),
      Colaborador_nome: colaboradorNome,
      Dia:              diaDestino,
      Opcao:            opcaoDestino,
      Nome_Prato:       nomePratoDestino,
      Confirmado:       true,
      Data_Hora:        `${dataDestino}T12:00:00`,
      Centro_Custo:     centroCusto || "",
      Status:           "Confirmado",
      Observacao:       obsAdicional,
      Origem:           origemAdicional,
      Alterado_Por:     this.getUserName()
    });

    let pedidoOrigem = null;
    let ausenciaOrigem = null;

    if (this.norm(tipoAjuste).includes("transfer")) {
      if (!diaOrigem) throw new Error("Informe o dia que será removido na transferência.");
      const dataOrigem = this._pedidoDataRefeicao(semanaId, diaOrigem, base);
      const obsOrigem = [
        "Refeição transferida para outro dia.",
        `Transferida para:${diaDestino}`,
        observacaoUsuario,
        adicional?.id ? `PedidoAdicional:${adicional.id}` : ""
      ].filter(Boolean).join(" | ");

      const existenteOrigem = await this._buscarPedidoNormalColaboradorDia(semanaId, colaboradorIdBase, colaboradorNome, diaOrigem);
      const camposOrigem = {
        Semana_id: semanaId,
        Colaborador_id: colaboradorIdBase,
        Colaborador_nome: colaboradorNome,
        Dia: diaOrigem,
        Opcao: "principal",
        Nome_Prato: "Não vai almoçar",
        Confirmado: false,
        Data_Hora: `${dataOrigem}T12:00:00`,
        Centro_Custo: centroCusto || "",
        Status: "Não vai almoçar",
        Observacao: obsOrigem,
        Origem: "Transferência - dia removido",
        Alterado_Por: this.getUserName()
      };

      if (existenteOrigem?.id) pedidoOrigem = await this.updatePedido(existenteOrigem.id, camposOrigem);
      else pedidoOrigem = await this.savePedido(semanaId, colaboradorIdBase, colaboradorNome, diaOrigem, "principal", "Não vai almoçar", camposOrigem);

      await this._garantirAusenciaDoPedidoOperacao({
        ...base,
        ...camposOrigem,
        id: this.pick(pedidoOrigem, "id", "ID") || this.pick(existenteOrigem, "id", "ID") || ""
      }, "Não vai almoçar", semanaId, diaOrigem);

      ausenciaOrigem = true;
    }

    try {
      localStorage.setItem("homy_refeitorio_sync", JSON.stringify({ tipo:"pedido", acao:"adicional", semanaId, ts: Date.now() }));
    } catch (_) {}

    return { adicional, pedidoOrigem, ausenciaOrigem };
  },

  async alterarPedidoOperacao(id, status, pedidoAtual = {}) {
    const isVirtual = !id || String(id).startsWith("ausencia-") || String(id).startsWith("pendente-") || pedidoAtual?._virtualAusencia || pedidoAtual?._virtualPendente;

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
  },
  // ============================================================
  // FECHAMENTO OFICIAL DO DIA — Verdade operacional auditável
  // Listas novas:
  // - Fechamento Diario Refeitorio
  // - Fechamento Itens Refeitorio
  // - Auditoria Refeitorio
  // ============================================================
  _fechamentoListas() {
    return {
      diario: "Fechamento Diario Refeitorio",
      itens: "Fechamento Itens Refeitorio",
      auditoria: "Auditoria Refeitorio"
    };
  },

  _fechamentoKey(semanaId, dia) {
    return `${String(semanaId || "").trim()}-${this.norm(dia || "")}`;
  },

  _jsonSeguro(obj) {
    try { return JSON.stringify(obj ?? null); }
    catch (_) { return JSON.stringify({ erro: "json_invalido" }); }
  },

  _hashTexto(texto) {
    const s = String(texto || "");
    let h = 2166136261;
    for (let i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    return `${s.length}-${(h >>> 0).toString(16)}`;
  },

  _opcaoResumoFechamento(opcao) {
    const op = this.norm(opcao || "");
    if (op === "principal") return "principal";
    if (op === "light") return "light";
    if (op === "carne") return "carne";
    if (op === "massa") return "massa";
    if (op === "lanche") return "lanche";
    if (!op) return "semOpcao";
    return "outros";
  },

  _categoriaPedidoFechamento(p) {
    const origem = this.norm(this.pick(p, "Origem", "origem", "tipo", "Tipo") || "");
    const nome = this.norm(this.pick(p, "Colaborador_nome", "Nome", "Title") || "");
    const id = this.norm(this.pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId") || "");

    if (origem.includes("investigador") || nome.includes("investigador")) return "investigador";
    if (origem.includes("guarda") || nome.includes("guarda")) return "guarda";
    if (origem.includes("prestador")) return "prestador";
    if (origem.includes("visitante")) return "visitante";
    if (origem.includes("terceiro")) return "terceiro";
    if (origem.includes("extra") || nome.includes("refeicao extra") || id.startsWith("extra-")) return "extra";
    if (origem.includes("ausencia") || origem.includes("ausência")) return "ausencia";
    return "colaborador";
  },

  _statusBloqueiaFechamento(p) {
    const status = this.norm(this.pick(p, "Status", "status") || "");
    return [
      "cancelado", "bloqueado", "nao vai almocar", "nao_vai_almocar",
      "não vai almoçar", "ausente", "ferias", "férias", "afastado",
      "atestado", "licenca", "licença", "banco horas", "banco_horas",
      "homy office", "homy_office", "falta", "duplicado inativado"
    ].includes(status);
  },

  _pedidoProdutivoFechamento(p) {
    const status = this.norm(this.pick(p, "Status", "status") || "");
    const origem = this.norm(this.pick(p, "Origem", "origem", "tipo", "Tipo") || "");
    const confirmado = this.isTrue(this.pick(p, "Confirmado", "confirmado"));

    if (status === "travado" && origem.includes("travamento")) return true;
    if (this._statusBloqueiaFechamento(p)) return false;
    return ["confirmado", "aprovado", "extra"].includes(status) || confirmado;
  },

  _pedidoTimestampFechamento(p) {
    const raw = this.pick(p, "Modified", "modified", "Data_Hora", "DataHora", "Created", "created") || "";
    const dt = raw ? new Date(raw) : null;
    if (dt && !isNaN(dt)) return dt.getTime();
    const id = Number(this.pick(p, "id", "ID") || 0);
    return Number.isFinite(id) ? id : 0;
  },

  _pedidoKeyFechamento(p) {
    const categoria = this._categoriaPedidoFechamento(p);
    const dia = this.norm(this.pick(p, "Dia", "dia") || "");
    const opcao = this.norm(this.pick(p, "Opcao", "opcao") || "principal");
    const id = String(this.pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    const nome = this.norm(this.pick(p, "Colaborador_nome", "Nome", "Title") || "");

    if (["extra", "guarda", "investigador", "prestador", "visitante", "terceiro"].includes(categoria)) {
      // Extras/especiais: id de extra vence; se não houver, usa nome/categoria/opção.
      return `especial|${dia}|${categoria}|${opcao}|${id || nome}`;
    }

    if (id) return `colaborador|${dia}|id:${id}`;
    return `colaborador|${dia}|nome:${nome}`;
  },

  _ordenarPreferenciaFechamento(a, b) {
    const ap = this._pedidoProdutivoFechamento(a) ? 1 : 0;
    const bp = this._pedidoProdutivoFechamento(b) ? 1 : 0;
    if (ap !== bp) return bp - ap;

    const ab = this._statusBloqueiaFechamento(a) ? 1 : 0;
    const bb = this._statusBloqueiaFechamento(b) ? 1 : 0;
    if (ab !== bb) return ab - bb;

    return this._pedidoTimestampFechamento(b) - this._pedidoTimestampFechamento(a);
  },

  _snapshotPedidoFechamento(p, contaProducao, motivo, categoria = null) {
    return {
      pedidoId: String(this.pick(p, "id", "ID") || ""),
      colaboradorId: String(this.pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId") || ""),
      colaboradorNome: this.pick(p, "Colaborador_nome", "Nome", "Title") || "",
      dia: this.pick(p, "Dia", "dia") || "",
      opcao: this.pick(p, "Opcao", "opcao") || "",
      status: this.pick(p, "Status", "status") || "",
      confirmado: this.isTrue(this.pick(p, "Confirmado", "confirmado")),
      origem: this.pick(p, "Origem", "origem", "tipo", "Tipo") || "",
      categoria: categoria || this._categoriaPedidoFechamento(p),
      contaProducao: !!contaProducao,
      motivo: motivo || "",
      centroCusto: this.pick(p, "Centro_Custo", "CentroCusto", "Setor", "Departamento") || "",
      dataHora: this.pick(p, "Data_Hora", "DataHora", "Data") || "",
      modified: this.pick(p, "Modified", "modified") || "",
      raw: p
    };
  },

  _calcularFechamentoPorPedidos(semanaId, dia, pedidos = []) {
    const diaNorm = this.norm(dia);
    const candidatos = (pedidos || []).filter(p =>
      String(this.pick(p, "Semana_id", "Semana") || "") === String(semanaId || "") &&
      this.norm(this.pick(p, "Dia", "dia") || "") === diaNorm
    );

    const porKey = new Map();
    for (const p of candidatos) {
      const key = this._pedidoKeyFechamento(p);
      if (!key) continue;
      if (!porKey.has(key)) porKey.set(key, []);
      porKey.get(key).push(p);
    }

    const resumo = {
      total: 0,
      principal: 0,
      light: 0,
      carne: 0,
      massa: 0,
      lanche: 0,
      extras: 0,
      cancelados: 0,
      ausentes: 0,
      duplicidadesIgnoradas: 0,
      outros: 0,
      semOpcao: 0
    };

    const incluidos = [];
    const excluidos = [];

    for (const grupo of porKey.values()) {
      const ordenado = [...grupo].sort((a, b) => this._ordenarPreferenciaFechamento(a, b));
      const preferido = ordenado[0];
      const categoria = this._categoriaPedidoFechamento(preferido);

      if (this._pedidoProdutivoFechamento(preferido)) {
        const opKey = this._opcaoResumoFechamento(this.pick(preferido, "Opcao", "opcao") || "principal");
        resumo[opKey] = (resumo[opKey] || 0) + 1;
        resumo.total++;
        if (["extra", "guarda", "investigador", "prestador", "visitante", "terceiro"].includes(categoria)) resumo.extras++;
        incluidos.push(this._snapshotPedidoFechamento(preferido, true, "Conta no fechamento oficial.", categoria));
      } else {
        const status = this.norm(this.pick(preferido, "Status", "status") || "");
        if (status === "cancelado" || status === "bloqueado" || status === "duplicado inativado") resumo.cancelados++;
        if (this._statusBloqueiaFechamento(preferido) && status !== "cancelado" && status !== "bloqueado" && status !== "duplicado inativado") resumo.ausentes++;
        excluidos.push(this._snapshotPedidoFechamento(preferido, false, `Não conta: ${this.pick(preferido, "Status", "status") || "status não produtivo"}.`, categoria));
      }

      for (const duplicado of ordenado.slice(1)) {
        resumo.duplicidadesIgnoradas++;
        excluidos.push(this._snapshotPedidoFechamento(duplicado, false, `Duplicado ignorado no fechamento. Mantido: ${this.pick(preferido, "id", "ID") || "registro preferido"}.`, this._categoriaPedidoFechamento(duplicado)));
      }
    }

    return { resumo, incluidos, excluidos, totalBruto: candidatos.length };
  },

  async gerarPreviaFechamentoDia(semanaId, dia, options = {}) {
    await this.ensureLogin?.();
    const diaNorm = this.norm(dia || "");
    if (!semanaId || !diaNorm) throw new Error("Informe semana e dia para gerar o fechamento.");

    const [pedidos, checkins, existente] = await Promise.all([
      options.pedidosBase ? Promise.resolve(options.pedidosBase) : this.getPedidos(semanaId, { reparar: false, force: true }),
      this.getCheckIn ? this.getCheckIn(semanaId, diaNorm).catch(() => []) : Promise.resolve([]),
      this.getFechamentoDia(semanaId, diaNorm).catch(() => null)
    ]);

    const calc = this._calcularFechamentoPorPedidos(semanaId, diaNorm, pedidos || []);
    const dataOperacao = this.getDataRefBySemanaDia(semanaId, diaNorm);
    const retiradas = (checkins || []).filter(c => this.isTrue(this.pick(c, "Retirou", "retirou"))).length;
    const key = this._fechamentoKey(semanaId, diaNorm);

    const snapshot = {
      key,
      semanaId,
      dia: diaNorm,
      dataOperacao,
      geradoEm: new Date().toISOString(),
      geradoPor: this.getUserName ? this.getUserName() : "Sistema",
      totais: { ...calc.resumo, checkins: retiradas },
      incluidos: calc.incluidos,
      excluidos: calc.excluidos,
      existente: existente || null
    };
    snapshot.hashResumo = this._hashTexto(this._jsonSeguro({ totais: snapshot.totais, incluidos: snapshot.incluidos.map(i => i.pedidoId), excluidos: snapshot.excluidos.map(i => [i.pedidoId, i.motivo]) }));
    return snapshot;
  },

  _statusFechamentoAtivo(item) {
    const status = this.norm(this.pick(item, "Status_Fechamento", "Status") || "");
    return status === "fechado" || status === "recalculado";
  },

  _dataFechamentoOrdem(item) {
    const raw = this.pick(item, "Fechado_Em", "Modified", "Gerado_Em", "Created") || "";
    const t = raw ? new Date(raw).getTime() : 0;
    return Number.isFinite(t) ? t : Number(this.pick(item, "id", "ID") || 0) || 0;
  },

  _motivoTotalZeroValido(motivo) {
    const texto = String(motivo || "").trim();
    const n = this.norm(texto);
    if (texto.length < 5) return false;
    return ![
      "fechamento conferido pela operacao do dia",
      "fechamento oficial gerado pela operacao do dia",
      "fechamento oficial confirmado",
      "ok",
      "sim"
    ].includes(n);
  },

  async getFechamentosSemana(semanaId) {
    const lista = this._fechamentoListas().diario;
    const items = await this.getItems(lista, { force: true }).catch(() => []);
    return (items || []).filter(i => String(this.pick(i, "Semana_id") || "") === String(semanaId || ""));
  },

  async getFechamentoDia(semanaId, dia, options = {}) {
    const key = this._fechamentoKey(semanaId, dia);
    const diaNorm = this.norm(dia || "");
    const semana = await this.getFechamentosSemana(semanaId);

    let candidatos = (semana || []).filter(i =>
      String(this.pick(i, "Title") || "") === key ||
      (String(this.pick(i, "Semana_id") || "") === String(semanaId || "") && this.norm(this.pick(i, "Dia") || "") === diaNorm)
    );

    if (!options.incluirInativos) {
      candidatos = candidatos.filter(i => this._statusFechamentoAtivo(i));
    }

    candidatos.sort((a, b) => this._dataFechamentoOrdem(b) - this._dataFechamentoOrdem(a));
    return candidatos[0] || null;
  },

  async getItensFechamento(fechamentoKey) {
    const lista = this._fechamentoListas().itens;
    const items = await this.getItems(lista, { force: true }).catch(() => []);
    return (items || []).filter(i => String(this.pick(i, "Fechamento_Key") || "") === String(fechamentoKey || ""));
  },

  async _limparItensFechamento(fechamentoKey) {
    const lista = this._fechamentoListas().itens;
    const itens = await this.getItensFechamento(fechamentoKey);
    for (const item of itens) {
      const id = this.pick(item, "id", "ID");
      if (id) await this.deleteItem(lista, id).catch(e => console.warn("[Fechamento] Falha ao remover item antigo:", e));
    }
    return itens.length;
  },

  async _salvarItemFechamento(fechamentoKey, semanaId, dia, dataOperacao, item) {
    const lista = this._fechamentoListas().itens;
    return this.createItem(lista, {
      Title: `${fechamentoKey}-${item.pedidoId || this.norm(item.colaboradorNome || "item")}`,
      Fechamento_Key: fechamentoKey,
      Semana_id: semanaId,
      Dia: dia,
      Data_Operacao: dataOperacao,
      Pedido_Id: item.pedidoId || "",
      Colaborador_id: item.colaboradorId || "",
      Colaborador_nome: item.colaboradorNome || "",
      Opcao: this.norm(item.opcao || "") || "sem_opcao",
      Status_Pedido: item.status || "",
      Confirmado: !!item.confirmado,
      Origem: item.origem || "",
      Categoria: item.categoria || "outro",
      Conta_Producao: !!item.contaProducao,
      Motivo_Decisao: item.motivo || "",
      Centro_Custo: item.centroCusto || "",
      Data_Hora_Pedido: item.dataHora || null,
      Modified_Pedido: item.modified || null,
      Snapshot_Item_JSON: this._jsonSeguro(item)
    });
  },

  async salvarFechamentoDia(semanaId, dia, options = {}) {
    await this.ensureLogin?.();
    if (options.confirmado !== true && options.confirmacaoExplicita !== true) {
      throw new Error("Fechamento bloqueado: confirmação explícita ausente.");
    }

    const previa = options.previa || await this.gerarPreviaFechamentoDia(semanaId, dia, options);
    const listas = this._fechamentoListas();
    const existente = await this.getFechamentoDia(semanaId, dia).catch(() => null);
    const statusExistente = this.norm(this.pick(existente, "Status_Fechamento", "Status") || "");

    if (existente?.id && statusExistente === "fechado" && !options.recalcular && !options.sobrescrever) {
      throw new Error("Este dia já está fechado. Reabra ou use recalcular para substituir o fechamento.");
    }

    const total = Number(previa?.totais?.total || 0);
    const motivoZero = String(options.motivoTotalZero || options.motivoZero || "").trim();
    if (total === 0 && !this._motivoTotalZeroValido(motivoZero)) {
      throw new Error("Fechamento com Total 0 bloqueado: informe motivo operacional obrigatório, como 'Não terá almoço', 'Feriado' ou 'Empresa sem expediente'.");
    }

    const agora = new Date().toISOString();
    const usuario = this.getUserName ? this.getUserName() : "Sistema";
    const observacaoBase = String(options.observacao || "").trim();
    const observacaoFinal = total === 0
      ? [`Total 0: ${motivoZero}`, observacaoBase].filter(Boolean).join(" | ")
      : (observacaoBase || "Fechamento oficial confirmado pela Operação do Dia.");

    const fields = {
      Title: previa.key,
      Semana_id: semanaId,
      Dia: previa.dia,
      Data_Operacao: previa.dataOperacao,
      Status_Fechamento: options.recalcular ? "Recalculado" : "Fechado",
      Total: total,
      Principal: Number(previa.totais.principal || 0),
      Light: Number(previa.totais.light || 0),
      Carne: Number(previa.totais.carne || 0),
      Massa: Number(previa.totais.massa || 0),
      Lanche: Number(previa.totais.lanche || 0),
      Extras: Number(previa.totais.extras || 0),
      Cancelados: Number(previa.totais.cancelados || 0),
      Ausentes: Number(previa.totais.ausentes || 0),
      Checkins: Number(previa.totais.checkins || 0),
      Duplicidades_Ignoradas: Number(previa.totais.duplicidadesIgnoradas || 0),
      Gerado_Por: previa.geradoPor || usuario,
      Gerado_Em: previa.geradoEm || agora,
      Fechado_Por: usuario,
      Fechado_Em: agora,
      Hash_Resumo: previa.hashResumo,
      Snapshot_JSON: this._jsonSeguro({ ...previa, motivoTotalZero: motivoZero || null }),
      Observacao: observacaoFinal
    };

    let fechamento;
    if (existente?.id) fechamento = await this.updateItem(listas.diario, existente.id, fields);
    else fechamento = await this.createItem(listas.diario, fields);

    await this._limparItensFechamento(previa.key);
    for (const item of [...previa.incluidos, ...previa.excluidos]) {
      await this._salvarItemFechamento(previa.key, semanaId, previa.dia, previa.dataOperacao, item);
    }

    await this.registrarAuditoriaRefeitorio({
      semanaId,
      dia: previa.dia,
      dataOperacao: previa.dataOperacao,
      modulo: "Operação do Dia",
      listaOrigem: listas.diario,
      itemId: this.pick(fechamento, "id", "ID") || this.pick(existente, "id", "ID") || "",
      acao: existente?.id ? "fechamento_recalculado" : "fechamento_criado",
      motivo: observacaoFinal,
      antes: existente || null,
      depois: fields,
      origemAcao: "fechamento-dia"
    }).catch(e => console.warn("[Fechamento] Falha ao registrar auditoria:", e));

    this.clearListCache(listas.diario);
    this.clearListCache(listas.itens);
    this.clearListCache(listas.auditoria);
    this._emitSync?.("fechamento", previa.key);

    return { fechamento, previa };
  },

  async cancelarFechamentoDia(semanaId, dia, motivo = "") {
    await this.ensureLogin?.();
    const listas = this._fechamentoListas();
    const existente = await this.getFechamentoDia(semanaId, dia, { incluirInativos: true });
    if (!existente?.id) throw new Error("Fechamento não encontrado para cancelar.");
    if (!String(motivo || "").trim()) throw new Error("Informe o motivo do cancelamento.");

    const fields = {
      Status_Fechamento: "Cancelado",
      Observacao: [this.pick(existente, "Observacao", "Observação") || "", `Cancelado: ${motivo}`].filter(Boolean).join(" | ")
    };
    const result = await this.updateItem(listas.diario, existente.id, fields);
    await this.registrarAuditoriaRefeitorio({
      semanaId,
      dia,
      dataOperacao: this.pick(existente, "Data_Operacao") || this.getDataRefBySemanaDia(semanaId, dia),
      modulo: "Operação do Dia",
      listaOrigem: listas.diario,
      itemId: existente.id,
      acao: "fechamento_cancelado",
      motivo,
      antes: existente,
      depois: { ...existente, ...fields },
      origemAcao: "cancelamento-fechamento"
    }).catch(e => console.warn("[Fechamento] Falha ao auditar cancelamento:", e));
    this.clearListCache(listas.diario);
    this._emitSync?.("fechamento", this._fechamentoKey(semanaId, dia));
    return result;
  },

  async reabrirFechamentoDia(semanaId, dia, motivo = "") {
    await this.ensureLogin?.();
    const listas = this._fechamentoListas();
    const existente = await this.getFechamentoDia(semanaId, dia);
    if (!existente?.id) throw new Error("Não há fechamento para reabrir.");
    if (!String(motivo || "").trim()) throw new Error("Informe o motivo da reabertura.");

    const fields = {
      Status_Fechamento: "Reaberto",
      Reaberto_Por: this.getUserName ? this.getUserName() : "Sistema",
      Reaberto_Em: new Date().toISOString(),
      Motivo_Reabertura: motivo
    };
    const result = await this.updateItem(listas.diario, existente.id, fields);
    await this.registrarAuditoriaRefeitorio({
      semanaId,
      dia,
      dataOperacao: this.pick(existente, "Data_Operacao") || this.getDataRefBySemanaDia(semanaId, dia),
      modulo: "Operação do Dia",
      listaOrigem: listas.diario,
      itemId: existente.id,
      acao: "fechamento_reaberto",
      motivo,
      antes: existente,
      depois: { ...existente, ...fields },
      origemAcao: "reabertura"
    }).catch(e => console.warn("[Fechamento] Falha ao auditar reabertura:", e));
    return result;
  },

  async registrarAuditoriaRefeitorio(params = {}) {
    const lista = this._fechamentoListas().auditoria;
    const agora = new Date().toISOString();
    const usuario = this.getUserName ? this.getUserName() : "Sistema";
    const email = this.getUserEmail ? this.getUserEmail() : "";
    const antes = params.antes ?? null;
    const depois = params.depois ?? null;
    const hash = this._hashTexto(this._jsonSeguro({ antes, depois, acao: params.acao, itemId: params.itemId, ts: agora }));
    return this.createItem(lista, {
      Title: `${params.semanaId || "semana"}-${params.dia || "dia"}-${params.acao || "acao"}-${params.itemId || Date.now()}`,
      Semana_id: params.semanaId || "",
      Dia: params.dia || "",
      Data_Operacao: params.dataOperacao || null,
      Modulo: params.modulo || "Sistema",
      Lista_Origem: params.listaOrigem || "",
      Item_Id: String(params.itemId || ""),
      Acao: params.acao || "outro",
      Motivo: params.motivo || "",
      Antes_JSON: this._jsonSeguro(antes),
      Depois_JSON: this._jsonSeguro(depois),
      Usuario: usuario,
      Email_Usuario: email,
      Data_Hora: agora,
      Hash_Auditoria: hash,
      Origem_Acao: params.origemAcao || "sistema"
    });
  },
  // ============================================================
  // CORREÇÃO ASSISTIDA DE INTEGRIDADE — v9
  // Objetivo:
  // - limpar bases históricas sujas antes do fechamento oficial;
  // - nunca aplicar ação silenciosa;
  // - só gravar com confirmação do usuário pela tela;
  // - após existir fechamento oficial, auditoria compara contra o fechamento.
  // ============================================================
  _referenciasHistoricasOperacionais() {
    return {
      "2026-W26": {
        fonte: "Luana",
        valores: {
          segunda: { principal: 44, light: 4, carne: 17, massa: 0, lanche: 0, total: 65 },
          terca:   { principal: 49, light: 6, carne: 8,  massa: 6, lanche: 0, total: 69 },
          quarta:  { principal: 45, light: 3, carne: 17, massa: 5, lanche: 0, total: 70 },
          quinta:  { principal: 57, light: 5, carne: 4,  massa: 4, lanche: 0, total: 70 }
        }
      }
    };
  },

  getReferenciaHistoricaOperacional(semanaId) {
    return this._referenciasHistoricasOperacionais()[String(semanaId || "").trim()] || null;
  },

  _resumoCorrecaoVazio() {
    return { total: 0, principal: 0, light: 0, carne: 0, massa: 0, lanche: 0, outros: 0, semOpcao: 0 };
  },

  _normalizarResumoCorrecao(resumo = {}) {
    const out = this._resumoCorrecaoVazio();
    for (const k of Object.keys(out)) out[k] = Number(resumo?.[k] || 0);
    return out;
  },

  _deltaResumoCorrecao(atual = {}, alvo = {}) {
    const a = this._normalizarResumoCorrecao(atual);
    const b = this._normalizarResumoCorrecao(alvo);
    return {
      total: a.total - b.total,
      principal: a.principal - b.principal,
      light: a.light - b.light,
      carne: a.carne - b.carne,
      massa: a.massa - b.massa,
      lanche: a.lanche - b.lanche
    };
  },

  _resumoBateCorrecao(atual = {}, alvo = {}) {
    const d = this._deltaResumoCorrecao(atual, alvo);
    return ["total", "principal", "light", "carne", "massa", "lanche"].every(k => Number(d[k] || 0) === 0);
  },

  _aplicarDeltaResumoCorrecao(resumo, delta = {}) {
    const r = this._normalizarResumoCorrecao(resumo);
    for (const k of ["total", "principal", "light", "carne", "massa", "lanche", "outros", "semOpcao"]) {
      if (delta[k]) r[k] = Number(r[k] || 0) + Number(delta[k] || 0);
    }
    return r;
  },

  _opcaoDeltaCorrecao(opcao) {
    const op = this._opcaoResumoFechamento ? this._opcaoResumoFechamento(opcao) : this.norm(opcao || "principal");
    return ["principal", "light", "carne", "massa", "lanche"].includes(op) ? op : "outros";
  },

  _pedidoByIdCorrecao(pedidos = []) {
    const m = new Map();
    for (const p of (pedidos || [])) {
      const id = String(this.pick(p, "id", "ID") || "");
      if (id) m.set(id, p);
    }
    return m;
  },

  _acaoCancelarPedidoCorrecao(item, motivo, justificativa, manterId = "") {
    const opcao = this._opcaoDeltaCorrecao(item.opcao || "principal");
    const delta = { total: -1 };
    delta[opcao] = -1;
    return {
      acao: "cancelar",
      autoAplicavel: true,
      motivo,
      pedidoId: String(item.pedidoId || ""),
      nome: item.colaboradorNome || item.nome || "",
      dia: item.dia || "",
      opcao,
      statusAtual: item.status || "",
      confirmadoAtual: !!item.confirmado,
      origemAtual: item.origem || "",
      categoria: item.categoria || "outro",
      manterId: manterId || "",
      delta,
      justificativa,
      camposSugeridos: {
        Status: "Cancelado",
        Confirmado: false,
        Origem: "Correção de integridade",
        Observacao: `Correção de integridade: ${justificativa}`
      },
      pedido: item
    };
  },

  _acaoReativarPedidoCorrecao(item, motivo, justificativa) {
    const opcao = this._opcaoDeltaCorrecao(item.opcao || "principal");
    const delta = { total: 1 };
    delta[opcao] = 1;
    return {
      acao: "reativar",
      autoAplicavel: true,
      motivo,
      pedidoId: String(item.pedidoId || ""),
      nome: item.colaboradorNome || item.nome || "",
      dia: item.dia || "",
      opcao,
      statusAtual: item.status || "",
      confirmadoAtual: !!item.confirmado,
      origemAtual: item.origem || "",
      categoria: item.categoria || "outro",
      delta,
      justificativa,
      camposSugeridos: {
        Status: "Confirmado",
        Confirmado: true,
        Opcao: opcao === "outros" ? (item.opcao || "principal") : opcao,
        Origem: "Correção de integridade",
        Observacao: `Correção de integridade: ${justificativa}`
      },
      pedido: item
    };
  },

  _acaoRevisarCorrecao(dia, opcao, quantidade, justificativa, candidatos = []) {
    return {
      acao: "revisar",
      autoAplicavel: false,
      motivo: "revisao-manual",
      dia,
      opcao,
      quantidade: Number(quantidade || 0),
      delta: {},
      justificativa,
      candidatos: candidatos.slice(0, 20)
    };
  },

  _selecionarAcoesQueAproximamCorrecao(atual, alvo, candidatas = []) {
    let simulado = this._normalizarResumoCorrecao(atual);
    const selecionadas = [];

    const aindaSobra = (opcao) => {
      const d = this._deltaResumoCorrecao(simulado, alvo);
      return Number(d.total || 0) > 0 && Number(d[opcao] || 0) > 0;
    };

    const aindaFalta = (opcao) => {
      const d = this._deltaResumoCorrecao(simulado, alvo);
      return Number(d.total || 0) < 0 || Number(d[opcao] || 0) < 0;
    };

    // Primeiro reativa o que está faltando, porque isso pode ser necessário junto com cancelamentos.
    for (const acao of candidatas.filter(a => a.acao === "reativar")) {
      const op = this._opcaoDeltaCorrecao(acao.opcao || "principal");
      if (!aindaFalta(op)) continue;
      selecionadas.push(acao);
      simulado = this._aplicarDeltaResumoCorrecao(simulado, acao.delta);
    }

    // Depois cancela duplicados e retornos retroativos, sem passar do alvo.
    const ordemCancelamento = { "extra-duplicado": 1, "retorno-automatico-retroativo": 2 };
    const cancelamentos = candidatas
      .filter(a => a.acao === "cancelar")
      .sort((a, b) => (ordemCancelamento[a.motivo] || 9) - (ordemCancelamento[b.motivo] || 9));

    for (const acao of cancelamentos) {
      const op = this._opcaoDeltaCorrecao(acao.opcao || "principal");
      if (!aindaSobra(op)) continue;
      selecionadas.push(acao);
      simulado = this._aplicarDeltaResumoCorrecao(simulado, acao.delta);
    }

    return { selecionadas, simulado };
  },

  async _referenciaComparacaoCorrecao(semanaId, dia, options = {}) {
    const diaNorm = this.norm(dia || "");
    const fechamento = await this.getFechamentoDia(semanaId, diaNorm).catch(() => null);
    const statusFechamento = this.norm(this.pick(fechamento, "Status_Fechamento", "Status") || "");
    if (fechamento?.id && statusFechamento !== "reaberto" && statusFechamento !== "cancelado") {
      return {
        tipo: "fechamento-oficial",
        fonte: "Fechamento Oficial",
        valores: {
          total: Number(this.pick(fechamento, "Total") || 0),
          principal: Number(this.pick(fechamento, "Principal") || 0),
          light: Number(this.pick(fechamento, "Light") || 0),
          carne: Number(this.pick(fechamento, "Carne") || 0),
          massa: Number(this.pick(fechamento, "Massa") || 0),
          lanche: Number(this.pick(fechamento, "Lanche") || 0)
        },
        fechamento
      };
    }

    const historica = this.getReferenciaHistoricaOperacional(semanaId);
    const refDia = historica?.valores?.[diaNorm];
    if (refDia) {
      return {
        tipo: "referencia-historica",
        fonte: historica.fonte || "Referência histórica",
        valores: this._normalizarResumoCorrecao(refDia),
        fechamento: null
      };
    }

    return null;
  },

  async validarPreviaFechamentoContraReferencia(semanaId, dia, previa = null, options = {}) {
    const ref = await this._referenciaComparacaoCorrecao(semanaId, dia, options);
    if (!ref) {
      return { temReferencia: false, bloquear: false, status: "sem-referencia", mensagem: "Sem fechamento oficial ou referência histórica para comparar." };
    }

    const atual = this._normalizarResumoCorrecao(previa?.totais || {});
    const alvo = this._normalizarResumoCorrecao(ref.valores || {});
    const delta = this._deltaResumoCorrecao(atual, alvo);
    const bate = this._resumoBateCorrecao(atual, alvo);
    const bloquear = !bate && (ref.tipo === "fechamento-oficial" || ref.tipo === "referencia-historica");

    return {
      temReferencia: true,
      tipoReferencia: ref.tipo,
      fonte: ref.fonte,
      bloquear,
      bate,
      atual,
      alvo,
      delta,
      mensagem: bate
        ? `Prévia bate com ${ref.fonte}.`
        : `Prévia divergente de ${ref.fonte}. Corrija pela Correção Assistida antes de fechar.`
    };
  },

  _candidatasCorrecaoDia(calc, alvo, dia) {
    const incluidos = calc?.incluidos || [];
    const excluidos = calc?.excluidos || [];
    const candidatas = [];

    // Extras duplicados: mesmo dia + categoria especial + nome/colaborador/opção.
    const gruposExtras = new Map();
    for (const item of incluidos) {
      const cat = this.norm(item.categoria || "");
      if (!["guarda", "investigador", "extra", "prestador", "visitante", "terceiro"].includes(cat)) continue;
      const key = [dia, cat, this.norm(item.colaboradorId || item.colaboradorNome || ""), this.norm(item.colaboradorNome || ""), this._opcaoDeltaCorrecao(item.opcao || "principal")].join("|");
      if (!gruposExtras.has(key)) gruposExtras.set(key, []);
      gruposExtras.get(key).push(item);
    }

    for (const grupo of gruposExtras.values()) {
      if (grupo.length <= 1) continue;
      const ordenado = [...grupo].sort((a, b) => Number(a.pedidoId || 0) - Number(b.pedidoId || 0));
      const manter = ordenado[0];
      for (const dup of ordenado.slice(1)) {
        candidatas.push(this._acaoCancelarPedidoCorrecao(
          dup,
          "extra-duplicado",
          `Extra duplicado: manter ID ${manter.pedidoId} e cancelar este registro.`,
          manter.pedidoId
        ));
      }
    }

    // Retornos automáticos retroativos: só são candidatos, não vencem duplicidade.
    for (const item of incluidos) {
      const origem = this.norm(item.origem || "");
      if (!origem.includes("retorno automatico") && !origem.includes("retorno automático")) continue;
      candidatas.push(this._acaoCancelarPedidoCorrecao(
        item,
        "retorno-automatico-retroativo",
        "Retorno automático criado em dia operacional passado."
      ));
    }

    // Pedido cancelado/duplicado que pode preencher opção faltante.
    const deltaInicial = this._deltaResumoCorrecao(calc.resumo, alvo);
    for (const item of excluidos) {
      const status = this.norm(item.status || "");
      const origem = this.norm(item.origem || "");
      const op = this._opcaoDeltaCorrecao(item.opcao || "principal");
      if (Number(deltaInicial[op] || 0) >= 0) continue;
      if (status !== "cancelado" && !origem.includes("duplicado inativado")) continue;
      candidatas.push(this._acaoReativarPedidoCorrecao(
        item,
        "pedido-correto-cancelado",
        `Pedido ${op} estava cancelado/inativado e é necessário para fechar a referência.`
      ));
    }

    return candidatas;
  },

  async gerarPlanoCorrecaoAssistida(semanaId, options = {}) {
    await this.ensureLogin?.();
    const diasBase = options.dia
      ? [this.norm(options.dia)]
      : ["segunda", "terca", "quarta", "quinta", "sexta"];

    const pedidos = options.pedidosBase || await this.getPedidos(semanaId, { reparar: false, force: true });
    const dias = [];

    for (const dia of diasBase) {
      const ref = await this._referenciaComparacaoCorrecao(semanaId, dia, options);
      if (!ref) continue;

      const calc = this._calcularFechamentoPorPedidos(semanaId, dia, pedidos || []);
      const atual = this._normalizarResumoCorrecao(calc.resumo);
      const alvo = this._normalizarResumoCorrecao(ref.valores);
      const deltaInicial = this._deltaResumoCorrecao(atual, alvo);
      const candidatas = this._candidatasCorrecaoDia(calc, alvo, dia);
      const { selecionadas, simulado } = this._selecionarAcoesQueAproximamCorrecao(atual, alvo, candidatas);
      const deltaFinal = this._deltaResumoCorrecao(simulado, alvo);
      const fechaExato = this._resumoBateCorrecao(simulado, alvo);
      const jaBate = this._resumoBateCorrecao(atual, alvo);

      const revisoes = [];
      for (const op of ["principal", "light", "carne", "massa", "lanche"]) {
        const sobra = Number(deltaFinal[op] || 0);
        if (sobra > 0) {
          const candidatos = (calc.incluidos || []).filter(i => this._opcaoDeltaCorrecao(i.opcao || "principal") === op);
          revisoes.push(this._acaoRevisarCorrecao(dia, op, sobra, `Ainda sobram ${sobra} ${op} após aplicar as ações seguras.`, candidatos));
        }
        if (sobra < 0) {
          revisoes.push(this._acaoRevisarCorrecao(dia, op, Math.abs(sobra), `Ainda faltam ${Math.abs(sobra)} ${op} após aplicar as ações seguras.`, []));
        }
      }

      dias.push({
        semanaId,
        dia,
        dataOperacao: this.getDataRefBySemanaDia?.(semanaId, dia) || "",
        referencia: ref,
        atual,
        alvo,
        deltaInicial,
        candidatas,
        acoesSeguras: selecionadas,
        revisoes,
        simulado,
        deltaFinal,
        jaBate,
        fechaExato,
        status: jaBate ? "ok" : (fechaExato ? "corrigivel" : (selecionadas.length ? "parcial" : "revisao")),
        mensagem: jaBate
          ? "Base atual já bate com a referência."
          : (fechaExato ? "Ações seguras fecham exatamente com a referência." : "Há ações seguras, mas ainda fica pendência para revisão.")
      });
    }

    const plano = {
      semanaId,
      geradoEm: new Date().toISOString(),
      status: "somente-leitura",
      dias,
      totais: {
        dias: dias.length,
        acoesSeguras: dias.reduce((n, d) => n + (d.acoesSeguras?.length || 0), 0),
        revisoes: dias.reduce((n, d) => n + (d.revisoes?.length || 0), 0),
        corrigiveis: dias.filter(d => d.fechaExato && !d.jaBate).length,
        parciais: dias.filter(d => d.status === "parcial").length
      }
    };
    plano.hashPlano = this._hashTexto(this._jsonSeguro({ semanaId, dias: dias.map(d => ({ dia: d.dia, acoes: d.acoesSeguras.map(a => [a.acao, a.pedidoId]), simulado: d.simulado })) }));
    return plano;
  },

  async aplicarPlanoCorrecaoAssistida(planoOuDia, options = {}) {
    await this.ensureLogin?.();
    const dias = Array.isArray(planoOuDia?.dias) ? planoOuDia.dias : [planoOuDia];
    const somenteSeFechaExato = options.somenteSeFechaExato === true;
    const aplicarParcial = options.aplicarParcial !== false;
    const resultados = [];

    for (const diaPlano of dias) {
      if (!diaPlano || diaPlano.jaBate) continue;
      if (somenteSeFechaExato && !diaPlano.fechaExato) {
        resultados.push({ dia: diaPlano.dia, pulado: true, motivo: "Plano não fecha exatamente com a referência." });
        continue;
      }
      if (!diaPlano.fechaExato && !aplicarParcial) {
        resultados.push({ dia: diaPlano.dia, pulado: true, motivo: "Aplicação parcial desabilitada." });
        continue;
      }

      for (const acao of (diaPlano.acoesSeguras || [])) {
        if (!acao.autoAplicavel || !acao.pedidoId) continue;
        const pedidoAntes = (await this.getItems("Pedidos", { force: true }).catch(() => []))
          .find(p => String(this.pick(p, "id", "ID") || "") === String(acao.pedidoId));

        const fields = { ...(acao.camposSugeridos || {}) };
        const atualizado = await this.updateItem("Pedidos", acao.pedidoId, fields);

        await this.registrarAuditoriaRefeitorio({
          semanaId: diaPlano.semanaId,
          dia: diaPlano.dia,
          dataOperacao: diaPlano.dataOperacao || this.getDataRefBySemanaDia?.(diaPlano.semanaId, diaPlano.dia) || null,
          modulo: "Correção Assistida",
          listaOrigem: "Pedidos",
          itemId: acao.pedidoId,
          acao: acao.acao === "reativar" ? "pedido_reativado" : (acao.motivo === "extra-duplicado" ? "extra_cancelado" : "pedido_cancelado"),
          motivo: acao.justificativa || acao.motivo || "Correção assistida de integridade.",
          antes: pedidoAntes || null,
          depois: { ...(pedidoAntes || {}), ...fields },
          origemAcao: "correcao-assistida"
        }).catch(e => console.warn("[Correção Assistida] Falha ao auditar ação:", e));

        resultados.push({ dia: diaPlano.dia, pedidoId: acao.pedidoId, acao: acao.acao, motivo: acao.motivo, atualizado });
      }
    }

    this.clearListCache?.("Pedidos");
    this.clearListCache?.(this._fechamentoListas().auditoria);
    this._emitSync?.("pedidos", "correcao-assistida");
    return { aplicadoEm: new Date().toISOString(), total: resultados.filter(r => !r.pulado).length, resultados };
  },



};

// Export explícito para páginas que verificam window.SP antes do login.
if (typeof window !== 'undefined') window.SP = SP;
