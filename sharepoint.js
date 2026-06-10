// ============================================================
// sharepoint.js — Refeitório Homy · Microsoft Graph API
// v: homy-admin-stable-20260610-1
// ============================================================

const SP = window.SP = {
  clientId: "aa37acf9-f3bd-4d1e-968a-fde57f79094c",
  tenantId: "a2850abc-334a-4805-b6b2-420b4aef68a9",
  siteUrl: "homyquimica.sharepoint.com",
  sitePath: "/sites/Refeitrio-Homy",
  scopes: ["https://graph.microsoft.com/Sites.ReadWrite.All", "User.Read"],

  redirectUri: "https://eduardadefiume-homy.github.io/refeitorio-homy/index.html",

  _msalInstance: null,
  _account: null,
  _siteId: null,
  _listIds: {},

  listAliases: {
    Cardapio: ["Cardapio", "Cardápio"],
    Pedidos: ["Pedidos"],
    Colaboradores: ["Colaboradores"],
    Extras: ["Extras"],
    Configuracoes: ["Configuracoes", "Configurações"],
    Ausencias_Refeitorio: [
      "Ausencias_Refeitorio",
      "Ausências_Refeitorio",
      "Ausencias Refeitorio",
      "Ausências Refeitório",
      "Ausencias",
      "Ausências",
      "Ausencia",
      "Ausência"
    ],
    Valores_Refeicao: [
      "Valores_Refeicao",
      "Valores Refeicao",
      "Valores Refeição",
      "Valores de Refeicao",
      "Valores de Refeição",
      "Valor_Refeicao",
      "Valor Refeicao",
      "Valor Refeição",
      "Valores"
    ]
  },

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
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim();
  },

  moneyToNumber(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;
    const s = String(value || "")
      .replace(/[R$\s]/g, "")
      .replace(/\./g, "")
      .replace(",", ".");
    const n = Number(s);
    return Number.isFinite(n) ? n : 0;
  },

  async waitForMsal(timeoutMs = 8000) {
    const start = Date.now();

    while (!window.msal) {
      if (Date.now() - start > timeoutMs) {
        throw new Error("MSAL não carregou. Verifique se msal-browser.min.js está antes de sharepoint.js.");
      }
      await new Promise(resolve => setTimeout(resolve, 80));
    }

    return true;
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
    const [year, week] = String(semanaId || this.getCurrentWeekId()).split("-W").map(Number);

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

  getUserName() {
    return this._account?.name || this._account?.username || "Usuário Homy";
  },

  getUserEmail() {
    return this._account?.username || "";
  },

  isExtraPedido(p) {
    const origem = this.norm(this.pick(p, "Origem", "tipo", "Tipo") || "");
    const tipo = this.norm(this.pick(p, "tipo", "Tipo") || "");
    const nome = this.norm(this.pick(p, "Colaborador_nome", "Title", "Nome") || "");

    return origem.includes("extra") ||
      tipo.includes("extra") ||
      origem.includes("investigador") ||
      tipo.includes("investigador") ||
      origem.includes("guarda") ||
      tipo.includes("guarda") ||
      nome.includes("refeicao extra") ||
      nome.includes("refeição extra");
  },

  isAusenciaPedido(p) {
    const status = this.norm(this.pick(p, "Status") || "");
    const opcao = this.norm(this.pick(p, "Opcao") || "");
    const origem = this.norm(this.pick(p, "Origem") || "");

    return status.includes("nao vai almocar") ||
      status.includes("não vai almoçar") ||
      status.includes("ausente") ||
      status.includes("cancelado") ||
      status.includes("afastado") ||
      status.includes("ferias") ||
      status.includes("férias") ||
      status.includes("bloqueado") ||
      status.includes("travado") ||
      opcao.includes("ausente") ||
      opcao.includes("nao vou almocar") ||
      origem.includes("ausencia") ||
      origem.includes("ausência");
  },

  isPedidoProdutivo(p) {
    const status = this.norm(this.pick(p, "Status") || "");
    const opcao = this.norm(this.pick(p, "Opcao") || "");

    if (this.isAusenciaPedido(p)) return false;
    if (opcao === "ausente") return false;

    return ["confirmado", "extra", "aprovado"].includes(status) ||
      this.isTrue(this.pick(p, "Confirmado"));
  },

  async init() {
    if (this._msalInstance) {
      if (!this._account) this._restoreAccountFromCache();
      return !!this._account;
    }

    await this.waitForMsal();

    this._msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: this.clientId,
        authority: `https://login.microsoftonline.com/${this.tenantId}`,
        redirectUri: this.redirectUri,
        navigateToLoginRequestUrl: false
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true
      },
      system: {
        allowNativeBroker: false
      }
    });

    await this._msalInstance.initialize();

    this._restoreAccountFromCache();

    return !!this._account;
  },

  _restoreAccountFromCache() {
    if (!this._msalInstance) return false;

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

    this._account = null;
    return false;
  },

  async login() {
    await this.init();

    if (this._account) {
      return true;
    }

    const result = await this._msalInstance.loginPopup({
      scopes: this.scopes,
      prompt: "select_account"
    });

    if (!result?.account) {
      throw new Error("Login retornou sem conta.");
    }

    this._account = result.account;
    this._msalInstance.setActiveAccount(this._account);

    return true;
  },

  async loginSilencioso() {
    await this.init();

    if (!this._account) return false;

    try {
      await this._msalInstance.acquireTokenSilent({
        scopes: this.scopes,
        account: this._account
      });

      return true;
    } catch (e) {
      console.warn("[SP] loginSilencioso falhou:", e);
      return false;
    }
  },

  async ensureLogin(options = {}) {
    const interactive = options.interactive !== false;

    await this.init();

    if (this._account) return true;

    if (!interactive) return false;

    return this.login();
  },

  async ensureLoginSilenciosoOuAviso() {
    const ok = await this.ensureLogin({ interactive: false });

    if (ok) return true;

    this.mostrarAvisoLogin();
    return false;
  },

  mostrarAvisoLogin(containerId = "loginAviso") {
    let box = document.getElementById(containerId);

    if (!box) {
      box = document.createElement("div");
      box.id = containerId;
      box.style.cssText = `
        position: relative;
        z-index: 999;
        max-width: 900px;
        margin: 1.2rem auto;
        padding: 1rem 1.2rem;
        border-radius: 14px;
        border: 1px solid rgba(255,200,80,.25);
        background: rgba(255,180,0,.08);
        color: rgba(255,220,150,.95);
        font-family: Barlow, sans-serif;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: .9rem;
        flex-wrap: wrap;
      `;

      const ref = document.querySelector(".main-content, .content, .cards-area, .app, body");

      if (ref && ref !== document.body) {
        ref.prepend(box);
      } else {
        document.body.prepend(box);
      }
    }

    box.innerHTML = `
      <div style="display:flex;align-items:center;gap:.7rem;line-height:1.35">
        <span style="font-size:1.1rem">⚠️</span>
        <span>Não foi possível conectar ao SharePoint. Entre com sua conta Microsoft para continuar.</span>
      </div>
      <button type="button" id="btnLoginSharePointGlobal" style="
        border: none;
        border-radius: 10px;
        background: rgba(192,40,28,.9);
        color: #fff;
        font-weight: 700;
        padding: .7rem 1rem;
        cursor: pointer;
        font-family: Barlow, sans-serif;
      ">Entrar com Microsoft</button>
    `;

    document.getElementById("btnLoginSharePointGlobal")?.addEventListener("click", async () => {
      try {
        await this.login();
        location.reload();
      } catch (e) {
        console.error("[SP] login global:", e);
        alert("Erro ao entrar com Microsoft: " + (e.message || e));
      }
    });
  },

  async logout() {
    await this.init();

    const account = this._account;

    this._account = null;
    this._siteId = null;
    this._listIds = {};

    if (account) {
      await this._msalInstance.logoutPopup({ account });
    }
  },

  async getToken(options = {}) {
    const interactive = options.interactive !== false;

    await this.init();

    if (!this._account) {
      if (!interactive) return null;
      await this.login();
    }

    try {
      const r = await this._msalInstance.acquireTokenSilent({
        scopes: this.scopes,
        account: this._account
      });

      return r.accessToken;
    } catch (e) {
      console.warn("[SP] acquireTokenSilent falhou:", e);

      if (!interactive) return null;

      const r = await this._msalInstance.acquireTokenPopup({
        scopes: this.scopes
      });

      this._account = r.account || this._account;

      if (this._account) {
        this._msalInstance.setActiveAccount(this._account);
      }

      return r.accessToken;
    }
  },

  async graph(method, endpoint, body = null, options = {}) {
    const token = await this.getToken(options);

    if (!token) {
      throw new Error("Login Microsoft necessário.");
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
        Authorization: `Bearer ${token}`,
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

  _cleanFields(fields) {
    const READONLY = [
      "@odata.etag",
      "ComplianceAssetId",
      "AuthorId",
      "EditorId",
      "Created",
      "Modified",
      "id",
      "ID",
      "odata.type",
      "odata.id",
      "odata.editLink",
      "FileSystemObjectType",
      "ServerRedirectedEmbedUri",
      "ServerRedirectedEmbedUrl",
      "ContentTypeId",
      "OData__UIVersionString",
      "Attachments",
      "GUID",
      "_ModerationStatus",
      "_ModerationComments",
      "LinkTitleNoMenu",
      "LinkTitle"
    ];

    const out = {};

    for (const [k, v] of Object.entries(fields || {})) {
      if (!READONLY.includes(k) && !k.startsWith("@odata")) {
        out[k] = v;
      }
    }

    return out;
  },

  async getSiteId() {
    if (this._siteId) return this._siteId;

    const paths = [
      this.sitePath,
      "/sites/Refeitorio-Homy",
      "/sites/refeitorio-homy",
      "/sites/Refeitrio-Homy"
    ].filter(Boolean);

    const unique = [...new Set(paths)];
    let lastErr = null;

    for (const path of unique) {
      try {
        const data = await this.graph("GET", `/sites/${this.siteUrl}:${path}`);

        if (data?.id) {
          this._siteId = data.id;
          return this._siteId;
        }
      } catch (e) {
        lastErr = e;
      }
    }

    throw lastErr || new Error("Site SharePoint não encontrado.");
  },

  async getListId(displayName) {
    if (this._listIds[displayName]) return this._listIds[displayName];

    const siteId = await this.getSiteId();
    const data = await this.graph("GET", `/sites/${siteId}/lists?$select=id,displayName,name`);

    const candidates = this.listAliases[displayName] || [displayName];
    const normalizedCandidates = candidates.map(x => this.norm(x));

    const list = (data.value || []).find(l => {
      const display = this.norm(l.displayName);
      const name = this.norm(l.name);

      return normalizedCandidates.includes(display) ||
        normalizedCandidates.includes(name);
    });

    if (!list) {
      throw new Error(`Lista não encontrada no SharePoint: ${displayName}`);
    }

    this._listIds[displayName] = list.id;

    return list.id;
  },

  async getItems(listName, expandFields = true) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);

    const endpoint = expandFields
      ? `/sites/${siteId}/lists/${listId}/items?expand=fields&top=5000`
      : `/sites/${siteId}/lists/${listId}/items?top=5000`;

    const data = await this.graph("GET", endpoint);

    return (data.value || []).map(item => ({
      id: item.id,
      ...(item.fields || item)
    }));
  },

  async createItem(listName, fields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);

    return this.graph("POST", `/sites/${siteId}/lists/${listId}/items`, { fields });
  },

  async updateItem(listName, itemId, fields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);

    return this.graph("PATCH", `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, fields);
  },

  async deleteItem(listName, itemId) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);

    return this.graph("DELETE", `/sites/${siteId}/lists/${listId}/items/${itemId}`);
  },

  async getConfig(chave) {
    const items = await this.getItems("Configuracoes");
    const item = items.find(i => this.norm(this.pick(i, "Chave", "Title")) === this.norm(chave));
    return item ? this.pick(item, "Valor") : null;
  },

  async setConfig(chave, valor) {
    const items = await this.getItems("Configuracoes");
    const item = items.find(i => this.norm(this.pick(i, "Chave", "Title")) === this.norm(chave));

    const fields = {
      Title: chave,
      Chave: chave,
      Valor: valor
    };

    if (item?.id) return this.updateItem("Configuracoes", item.id, fields);

    return this.createItem("Configuracoes", fields);
  },

  async isCardapioLiberado() {
    const v = await this.getConfig("marcacao_liberada");
    return this.isTrue(v);
  },

  async setMarcacaoLiberada(valor) {
    return this.setConfig("marcacao_liberada", valor ? "sim" : "nao");
  },

  async setCardapioVisivel(valor) {
    return this.setConfig("cardapio_visivel", valor ? "sim" : "nao");
  },

  async getPrazoMarcacao() {
    return this.getConfig("prazo_marcacao");
  },

  async setPrazoMarcacao(valor) {
    return this.setConfig("prazo_marcacao", valor);
  },

  async getCardapio(semanaId) {
    const items = await this.getItems("Cardapio");

    return items
      .filter(i => String(this.pick(i, "Semana_id", "SemanaId", "SemanaID") || "") === String(semanaId))
      .sort((a, b) => {
        const diaOrd = { segunda: 1, terca: 2, terça: 2, quarta: 3, quinta: 4, sexta: 5 };
        const opOrd = { principal: 1, light: 2, carne: 3, massa: 4, lanche: 5 };

        const da = diaOrd[this.norm(this.pick(a, "Dia"))] || 99;
        const db = diaOrd[this.norm(this.pick(b, "Dia"))] || 99;

        if (da !== db) return da - db;

        const oa = opOrd[this.norm(this.pick(a, "Opcao"))] || 99;
        const ob = opOrd[this.norm(this.pick(b, "Opcao"))] || 99;

        return oa - ob;
      });
  },

  async saveCardapio(semanaId, dia, opcao, nomePrato, detalhes = "") {
    const items = await this.getCardapio(semanaId);

    const existente = items.find(i =>
      this.norm(this.pick(i, "Dia")) === this.norm(dia) &&
      this.norm(this.pick(i, "Opcao")) === this.norm(opcao)
    );

    const fields = {
      Title: `${semanaId}-${this.norm(dia)}-${this.norm(opcao)}`,
      Semana_id: semanaId,
      Dia: this.norm(dia),
      Opcao: this.norm(opcao),
      Nome_Prato: nomePrato,
      Detalhes: detalhes || ""
    };

    if (existente?.id) return this.updateItem("Cardapio", existente.id, fields);

    return this.createItem("Cardapio", fields);
  },

  async getTodosColaboradores(incluirInativos = false) {
    const items = await this.getItems("Colaboradores");

    return items
      .filter(c => incluirInativos || this.isTrue(this.pick(c, "Ativo")))
      .sort((a, b) =>
        String(this.pick(a, "Nome", "Title") || "").localeCompare(
          String(this.pick(b, "Nome", "Title") || ""),
          "pt-BR"
        )
      );
  },

  async getColaboradoresAtivos() {
    return this.getTodosColaboradores(false);
  },

  async getColaboradores() {
    return this.getColaboradoresAtivos();
  },

  async createColaborador({ nome, departamento, centroCusto, email, tipo, contaContabil }) {
    return this.createItem("Colaboradores", {
      Title: nome,
      Nome: nome,
      Departamento: departamento || "",
      Centro_Custo: centroCusto || "",
      Conta_Contabil: contaContabil || "",
      Email: email || "",
      tipo: tipo || "colaborador",
      Ativo: true
    });
  },

  async updateColaborador(id, { nome, departamento, centroCusto, email, tipo, contaContabil }) {
    const fields = {
      Title: nome,
      Nome: nome,
      Departamento: departamento || "",
      Centro_Custo: centroCusto || "",
      Email: email || "",
      tipo: tipo || "colaborador"
    };

    if (contaContabil !== undefined) {
      fields.Conta_Contabil = contaContabil || "";
    }

    return this.updateItem("Colaboradores", id, fields);
  },

  async desativarColaborador(id) {
    return this.updateItem("Colaboradores", id, { Ativo: false });
  },

  async deleteColaborador(id) {
    return this.deleteItem("Colaboradores", id);
  },

  async getPedidos(semanaId) {
    const items = await this.getItems("Pedidos");

    return items.filter(i =>
      !semanaId ||
      String(this.pick(i, "Semana_id", "SemanaId", "SemanaID") || "") === String(semanaId)
    );
  },

  async getPedidoColaboradorDia(semanaId, colaboradorId, dia) {
    const pedidos = await this.getPedidos(semanaId);

    return pedidos.find(p =>
      String(this.pick(p, "Colaborador_id")) === String(colaboradorId) &&
      this.norm(this.pick(p, "Dia")) === this.norm(dia)
    ) || null;
  },

  async savePedido(dados) {
    const semanaId = dados.semanaId || dados.Semana_id || this.getCurrentWeekId();
    const colaboradorId = dados.colaboradorId || dados.Colaborador_id;
    const dia = dados.dia || dados.Dia;

    const existente = colaboradorId
      ? await this.getPedidoColaboradorDia(semanaId, colaboradorId, dia)
      : null;

    const nome = dados.colaboradorNome || dados.Colaborador_nome || dados.nome || dados.Nome || "";

    const fields = {
      Title: nome || "Pedido",
      Semana_id: semanaId,
      Colaborador_id: String(colaboradorId || ""),
      Colaborador_nome: nome,
      Centro_Custo: dados.centroCusto || dados.Centro_Custo || "",
      Dia: this.norm(dia),
      Opcao: this.norm(dados.opcao || dados.Opcao || "principal"),
      Nome_Prato: dados.nomePrato || dados.Nome_Prato || "",
      Status: dados.status || dados.Status || "Confirmado",
      Confirmado: dados.confirmado !== undefined ? !!dados.confirmado : true,
      Data_Hora: dados.dataHora || dados.Data_Hora || new Date().toISOString(),
      Origem: dados.origem || dados.Origem || "Refeitório",
      Observacao: dados.observacao || dados.Observacao || ""
    };

    if (existente?.id) return this.updateItem("Pedidos", existente.id, fields);

    return this.createItem("Pedidos", fields);
  },

  async createPedido(dados) {
    return this.savePedido(dados);
  },

  async updatePedido(id, dados) {
    const fields = {};

    if (dados.dia || dados.Dia) fields.Dia = this.norm(dados.dia || dados.Dia);
    if (dados.opcao || dados.Opcao) fields.Opcao = this.norm(dados.opcao || dados.Opcao);
    if (dados.nomePrato || dados.Nome_Prato) fields.Nome_Prato = dados.nomePrato || dados.Nome_Prato;
    if (dados.status || dados.Status) fields.Status = dados.status || dados.Status;
    if (dados.Confirmado !== undefined || dados.confirmado !== undefined) fields.Confirmado = dados.Confirmado ?? dados.confirmado;
    if (dados.Alterado_Por) fields.Alterado_Por = dados.Alterado_Por;
    if (dados.origem || dados.Origem) fields.Origem = dados.origem || dados.Origem;
    if (dados.observacao || dados.Observacao) fields.Observacao = dados.observacao || dados.Observacao;

    return this.updateItem("Pedidos", id, fields);
  },

  async deletePedido(id) {
    return this.deleteItem("Pedidos", id);
  },

  async cancelarPedido(id, motivo = "") {
    return this.updateItem("Pedidos", id, {
      Status: "Cancelado",
      Confirmado: false,
      Observacao: motivo
    });
  },

  async confirmarRetirada(id) {
    return this.updateItem("Pedidos", id, {
      Checkin: true,
      Checkin_Hora: new Date().toISOString()
    });
  },

  async registrarAusenciaAlmoco({ semanaId, colaborador, dia, motivo = "" }) {
    const id = this.pick(colaborador, "id", "ID") || "";
    const nome = this.pick(colaborador, "Nome", "Title") || "";
    const centroCusto = this.pick(colaborador, "Centro_Custo") || "";

    return this.savePedido({
      semanaId,
      colaboradorId: id,
      colaboradorNome: nome,
      centroCusto,
      dia,
      opcao: "ausente",
      nomePrato: "Não irá almoçar",
      status: "Não vai almoçar",
      confirmado: false,
      origem: "Ausência",
      observacao: motivo || "Ausência informada pelo colaborador"
    });
  },

  async getExtras(semanaId, dia = "") {
    const items = await this.getItems("Extras");

    return items.filter(i => {
      const okSemana = !semanaId || String(this.pick(i, "Semana_id", "SemanaId", "SemanaID") || "") === String(semanaId);
      const okDia = !dia || this.norm(this.pick(i, "Dia")) === this.norm(dia);
      return okSemana && okDia;
    });
  },

  async createExtra(dados) {
    return this.createItem("Extras", {
      Title: dados.nome || dados.Nome || "Extra",
      Nome: dados.nome || dados.Nome || "Extra",
      tipo: dados.tipo || dados.Tipo || "extra",
      Semana_id: dados.semanaId || dados.Semana_id || this.getCurrentWeekId(),
      Dia: this.norm(dados.dia || dados.Dia || ""),
      Opcao: this.norm(dados.opcao || dados.Opcao || "principal"),
      Centro_Custo: dados.centroCusto || dados.Centro_Custo || "",
      Observacao: dados.observacao || dados.Observacao || "",
      Status: dados.status || dados.Status || "Confirmado",
      Data_Hora: dados.dataHora || dados.Data_Hora || new Date().toISOString()
    });
  },

  async deleteExtra(id) {
    return this.deleteItem("Extras", id);
  },

  async _addExtraPedidoCC(semanaId, dia, nome, tipo, opcao, observacao, centroCusto, user) {
    await this.createExtra({
      semanaId,
      dia,
      nome,
      tipo,
      opcao,
      observacao,
      centroCusto,
      status: "Confirmado"
    });

    return this.savePedido({
      semanaId,
      colaboradorId: `extra-${dia}`,
      colaboradorNome: nome || "Refeição Extra",
      centroCusto: centroCusto || "",
      dia,
      opcao: opcao || "principal",
      nomePrato: nome || "Refeição Extra",
      status: "Confirmado",
      confirmado: true,
      origem: tipo || "Extra",
      observacao: observacao || user || ""
    });
  },

  async deleteExtraComPedido(extra) {
    if (extra?.id) await this.deleteExtra(extra.id);
    return true;
  },

  async getAusenciasRefeitorio() {
    return this.getItems("Ausencias_Refeitorio");
  },

  async getAusencias(ativasSomente = false) {
    const items = await this.getAusenciasRefeitorio();

    if (!ativasSomente) return items;

    const hoje = new Date().toISOString().slice(0, 10);

    return items.filter(a => {
      const ini = String(this.pick(a, "Data_Inicio") || "").slice(0, 10);
      const fim = String(this.pick(a, "Data_Fim") || "").slice(0, 10);
      if (!ini || !fim) return false;
      return hoje >= ini && hoje <= fim;
    });
  },

  async createAusenciaRefeitorio(dados) {
    return this.createItem("Ausencias_Refeitorio", {
      Title: dados.titulo || dados.colaboradorNome || dados.nome || "Ausência",
      Colaborador_id: dados.colaboradorId || "",
      Colaborador_nome: dados.colaboradorNome || dados.nome || "",
      Data_Inicio: dados.dataInicio || dados.Data_Inicio || "",
      Data_Fim: dados.dataFim || dados.Data_Fim || "",
      Motivo: dados.motivo || dados.Motivo || "",
      Observacao: dados.observacao || dados.Observacao || "",
      Ativo: dados.ativo !== undefined ? !!dados.ativo : true
    });
  },

  async createAusencia(dados) {
    return this.createAusenciaRefeitorio(dados);
  },

  async updateAusenciaRefeitorio(id, dados) {
    return this.updateItem("Ausencias_Refeitorio", id, {
      Data_Inicio: dados.dataInicio || dados.Data_Inicio || "",
      Data_Fim: dados.dataFim || dados.Data_Fim || "",
      Motivo: dados.motivo || dados.Motivo || "",
      Observacao: dados.observacao || dados.Observacao || "",
      Ativo: dados.ativo !== undefined ? !!dados.ativo : true
    });
  },

  async deleteAusencia(id) {
    return this.deleteItem("Ausencias_Refeitorio", id);
  },

  async getValoresRefeicao(ativosSomente = true) {
    const items = await this.getItems("Valores_Refeicao");
    return items.filter(v => !ativosSomente || this.isTrue(this.pick(v, "Ativo")));
  },

  async getValorRefeicaoAtivo() {
    const valores = await this.getValoresRefeicao(true);
    return valores.length ? valores[0] : null;
  },

  async createValorRefeicao(fields) {
    return this.createItem("Valores_Refeicao", fields);
  },

  async updateValorRefeicao(id, fields) {
    return this.updateItem("Valores_Refeicao", id, fields);
  },

  async deleteValorRefeicao(id) {
    return this.deleteItem("Valores_Refeicao", id);
  },

  _resolveColunasValores(item) {
    const valorVascon = this.moneyToNumber(this.pick(
      item,
      "Valor_Vascon",
      "ValorVascon",
      "valor_vascon",
      "valorVascon",
      "Valor"
    ));

    const descontoFuncionario = this.moneyToNumber(this.pick(
      item,
      "Valor_Desconto_Funcionario",
      "Valor_Desconto_Funcion_x00e1_rio",
      "Desconto_Funcionario",
      "DescontoFuncionario",
      "Valor_Desconto",
      "Desconto"
    ));

    return {
      valorVascon,
      descontoFuncionario
    };
  },

  async getCheckIn() {
    return [];
  },

  async getDashboardResumo(semanaId) {
    const hoje = new Date();
    const diaHoje = ["domingo", "segunda", "terca", "quarta", "quinta", "sexta", "sabado"][hoje.getDay()] || "segunda";
    const hojeStr = hoje.toISOString().slice(0, 10);

    const [colabs, pedidos, extras, ausencias] = await Promise.all([
      this.getColaboradoresAtivos().catch(() => []),
      this.getPedidos(semanaId).catch(() => []),
      this.getExtras(semanaId).catch(() => []),
      this.getAusenciasRefeitorio().catch(() => [])
    ]);

    const pedidosHoje = pedidos.filter(p =>
      this.norm(this.pick(p, "Dia")) === this.norm(diaHoje)
    );

    const confirmados = pedidos.filter(p => this.isPedidoProdutivo(p));
    const confirmadosHoje = pedidosHoje.filter(p => this.isPedidoProdutivo(p));

    const ausenciasAtivasHoje = ausencias.filter(a => {
      const ini = String(this.pick(a, "Data_Inicio") || "").slice(0, 10);
      const fim = String(this.pick(a, "Data_Fim") || "").slice(0, 10);
      const ativo = this.isTrue(this.pick(a, "Ativo") ?? true);

      if (!ini || !fim || !ativo) return false;

      return hojeStr >= ini && hojeStr <= fim;
    });

    const ausenciasPedidosHoje = pedidosHoje.filter(p => this.isAusenciaPedido(p));

    const pendentesSemana = Math.max(0, colabs.length * 5 - pedidos.length);

    const porDia = {};

    ["segunda", "terca", "quarta", "quinta", "sexta"].forEach(dia => {
      const lista = pedidos.filter(p => this.norm(this.pick(p, "Dia")) === dia);
      const conf = lista.filter(p => this.isPedidoProdutivo(p));

      porDia[dia] = {
        total: conf.length,
        principal: conf.filter(p => this.norm(this.pick(p, "Opcao")) === "principal").length,
        light: conf.filter(p => this.norm(this.pick(p, "Opcao")) === "light").length,
        carne: conf.filter(p => this.norm(this.pick(p, "Opcao")) === "carne").length,
        massa: conf.filter(p => this.norm(this.pick(p, "Opcao")) === "massa").length,
        lanche: conf.filter(p => this.norm(this.pick(p, "Opcao")) === "lanche").length,
        pendentes: Math.max(0, colabs.length - lista.length)
      };
    });

    const setoresMap = new Map();

    confirmadosHoje.forEach(p => {
      const cc = this.pick(p, "Centro_Custo") || "Sem CC";
      setoresMap.set(cc, (setoresMap.get(cc) || 0) + 1);
    });

    return {
      colaboradoresAtivos: colabs.length,
      pedidosConfirmadosColaboradores: confirmados.length,
      pendentesColaboradores: pendentesSemana,
      checkinsHoje: pedidosHoje.filter(p => this.isTrue(this.pick(p, "Checkin"))).length,
      extrasAtivos: extras.length,
      extrasConfirmados: extras.filter(e => this.norm(this.pick(e, "Status")) === "confirmado").length,
      extrasPendentes: extras.filter(e => this.norm(this.pick(e, "Status")) === "pendente").length,
      totalPedidosHoje: confirmadosHoje.length,
      ausenciasHoje: Math.max(ausenciasPedidosHoje.length, ausenciasAtivasHoje.length),
      ausenciasAtivasHoje: ausenciasAtivasHoje.length,
      totalPedidosSemana: confirmados.length,
      principalHoje: confirmadosHoje.filter(p => this.norm(this.pick(p, "Opcao")) === "principal").length,
      lightHoje: confirmadosHoje.filter(p => this.norm(this.pick(p, "Opcao")) === "light").length,
      outrasHoje: confirmadosHoje.filter(p => !["principal", "light"].includes(this.norm(this.pick(p, "Opcao")))).length,
      setoresHoje: Array.from(setoresMap.entries()).map(([nome, total]) => ({ nome, total })),
      porDia,
      diaHoje
    };
  }
};

window.HOMY_SP_READY = (async () => {
  try {
    await SP.init();
    return true;
  } catch (e) {
    console.error("[HOMY_SP_READY]", e);
    return false;
  }
})();
