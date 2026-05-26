// ============================================================
// sharepoint.js — Camada de conexão com o SharePoint
// Refeitório Homy · Microsoft Graph API
// ============================================================

const SP = {
  // --- CONFIGURAÇÕES ---
  clientId: "2e73e8be-484b-428e-a6c0-c75b1bf70d8a",
  tenantId: "a2850abc-334a-4805-b6b2-420b4aef68a9",
  siteUrl: "homyquimica.sharepoint.com",
  sitePath: "/sites/Refeitrio-Homy",
  scopes: ["https://graph.microsoft.com/Sites.ReadWrite.All","User.Read"],

  // --- ESTADO ---
  _msalInstance: null,
  _account: null,
  _siteId: null,
  _listIds: {},

  // --- INICIALIZAR MSAL ---
  async init() {
    if (this._msalInstance) return true;
    const msalConfig = {
      auth: {
        clientId: this.clientId,
        authority: `https://login.microsoftonline.com/${this.tenantId}`,
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: { cacheLocation: "sessionStorage" }
    };
    this._msalInstance = new msal.PublicClientApplication(msalConfig);
    await this._msalInstance.initialize();
    const accounts = this._msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      this._account = accounts[0];
      return true;
    }
    return false;
  },

  // --- LOGIN ---
  async login() {
    await this.init();
    try {
      const result = await this._msalInstance.loginPopup({ scopes: this.scopes });
      this._account = result.account;
      return true;
    } catch (e) {
      console.error("Login falhou:", e);
      return false;
    }
  },

  // --- TOKEN ---
  async getToken() {
    await this.init();
    if (!this._account) await this.login();
    try {
      const result = await this._msalInstance.acquireTokenSilent({
        scopes: this.scopes,
        account: this._account
      });
      return result.accessToken;
    } catch {
      const result = await this._msalInstance.acquireTokenPopup({ scopes: this.scopes });
      return result.accessToken;
    }
  },

  // --- GRAPH API BASE ---
  async graph(method, endpoint, body = null) {
    const token = await this.getToken();
    const res = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
      method,
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: body ? JSON.stringify(body) : null
    });
    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Graph API error ${res.status}: ${err}`);
    }
    if (res.status === 204) return null;
    return res.json();
  },

  // --- SITE ID ---
  async getSiteId() {
    if (this._siteId) return this._siteId;
    const data = await this.graph("GET", `/sites/${this.siteUrl}:${this.sitePath}`);
    this._siteId = data.id;
    return this._siteId;
  },

  // --- LIST ID ---
  async getListId(listName) {
    if (this._listIds[listName]) return this._listIds[listName];
    const siteId = await this.getSiteId();
    const data = await this.graph("GET", `/sites/${siteId}/lists?$filter=displayName eq '${listName}'`);
    this._listIds[listName] = data.value[0].id;
    return this._listIds[listName];
  },

  // --- CRUD GENÉRICO ---
  async getItems(listName, filter = "") {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const url = `/sites/${siteId}/lists/${listId}/items?expand=fields${filter ? "&$filter=" + filter : ""}`;
    const data = await this.graph("GET", url);
    return data.value.map(i => ({ id: i.id, ...i.fields }));
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

  // ============================================================
  // COLABORADORES
  // ============================================================
  async getColaboradores() {
    return this.getItems("Colaboradores", "fields/ativo eq true");
  },
  async createColaborador(dados) {
    return this.createItem("Colaboradores", {
      Title: dados.nome,
      nome: dados.nome,
      departamento: dados.departamento,
      email: dados.email || "",
      ativo: true,
      tipo: dados.tipo || "colaborador"
    });
  },
  async updateColaborador(id, dados) {
    return this.updateItem("Colaboradores", id, dados);
  },
  async desativarColaborador(id) {
    return this.updateItem("Colaboradores", id, { ativo: false });
  },

  // ============================================================
  // CARDÁPIO
  // ============================================================
  async getCardapio(semanaId) {
    const items = await this.getItems("Card\u00e1pio");
    return items.filter(i => i.semana_id === semanaId);
  },
  async saveCardapio(semanaId, dia, opcao, nomePrato, detalhes) {
    return this.createItem("Card\u00e1pio", {
      Title: `${semanaId}-${dia}-${opcao}`,
      semana_id: semanaId,
      dia,
      opcao,
      Nome_Prato: nomePrato,
      Detalhes: detalhes
    });
  },
  async clearCardapio(semanaId) {
    const items = await this.getCardapio(semanaId);
    for (const item of items) {
      await this.deleteItem("Card\u00e1pio", item.id);
    }
  },

  // ============================================================
  // PEDIDOS
  // ============================================================
  async getPedidos(semanaId) {
    const items = await this.getItems("Pedidos");
    return items.filter(i => i.semana_id === semanaId);
  },
  async getPedidoColaborador(semanaId, colaboradorId) {
    const items = await this.getPedidos(semanaId);
    return items.filter(i => i.colaborador_id === String(colaboradorId));
  },
  async savePedido(semanaId, colaboradorId, colaboradorNome, dia, opcao, nomePrato) {
    return this.createItem("Pedidos", {
      Title: `${semanaId}-${colaboradorId}-${dia}`,
      semana_id: semanaId,
      colaborador_id: String(colaboradorId),
      colaborador_nome: colaboradorNome,
      dia,
      opcao,
      Nome_Prato: nomePrato,
      confirmado: false,
      Data_Hora: new Date().toISOString()
    });
  },
  async confirmarPedidos(semanaId, colaboradorId) {
    const pedidos = await this.getPedidoColaborador(semanaId, colaboradorId);
    for (const p of pedidos) {
      await this.updateItem("Pedidos", p.id, { confirmado: true });
    }
  },

  // ============================================================
  // CONFIGURAÇÕES
  // ============================================================
  async getConfig(chave) {
    const items = await this.getItems("Configura\u00e7\u00f5es");
    const item = items.find(i => i.chave === chave || i.Title === chave);
    return item ? item.valor || item.Valor : null;
  },
  async setConfig(chave, valor, descricao = "") {
    const items = await this.getItems("Configura\u00e7\u00f5es");
    const existing = items.find(i => i.chave === chave || i.Title === chave);
    if (existing) {
      return this.updateItem("Configura\u00e7\u00f5es", existing.id, { Valor: valor });
    }
    return this.createItem("Configura\u00e7\u00f5es", {
      Title: chave,
      chave,
      Valor: valor,
      Descri\u00e7\u00e3o: descricao
    });
  },
  async isCardapioLiberado() {
    const v = await this.getConfig("cardapio_liberado");
    return v === "sim";
  },
  async getPrazoMarcacao() {
    return this.getConfig("prazo_limite");
  },

  // ============================================================
  // CHECK-IN (cozinha)
  // ============================================================
  async getCheckIn(semanaId, dia) {
    const items = await this.getItems("CheckIn");
    return items.filter(i => i.semana_id === semanaId && i.dia === dia);
  },
  async registrarCheckIn(semanaId, colaboradorId, colaboradorNome, dia, confirmadoPor) {
    const existing = await this.getCheckIn(semanaId, dia);
    const found = existing.find(i => i.colaborador_id === String(colaboradorId));
    if (found) {
      return this.updateItem("CheckIn", found.id, {
        Retirou: true,
        Data_Hora_Retirada: new Date().toISOString(),
        Confirmado_Por: confirmadoPor
      });
    }
    return this.createItem("CheckIn", {
      Title: `${semanaId}-${colaboradorId}-${dia}`,
      semana_id: semanaId,
      colaborador_id: String(colaboradorId),
      colaborador_nome: colaboradorNome,
      dia,
      Retirou: true,
      Data_Hora_Retirada: new Date().toISOString(),
      Confirmado_Por: confirmadoPor
    });
  },

  // ============================================================
  // EXTRAS (visitantes, guarda, fixos)
  // ============================================================
  async getExtras(semanaId, dia = null) {
    const items = await this.getItems("Extras");
    return items.filter(i => i.semana_id === semanaId && (!dia || i.dia === dia));
  },
  async addExtra(semanaId, dia, nome, tipo, opcao, observacao, adicionadoPor) {
    return this.createItem("Extras", {
      Title: `${semanaId}-${dia}-${nome}`,
      semana_id: semanaId,
      dia,
      Nome: nome,
      tipo,
      opcao,
      Observacao: observacao,
      Adicionado_Por: adicionadoPor
    });
  },
  async removeExtra(id) {
    return this.deleteItem("Extras", id);
  },

  // ============================================================
  // UTILITÁRIOS
  // ============================================================
  getSemanaId(date = new Date()) {
    const d = new Date(date);
    d.setHours(0,0,0,0);
    d.setDate(d.getDate() + 3 - (d.getDay() + 6) % 7);
    const week1 = new Date(d.getFullYear(), 0, 4);
    const weekNum = 1 + Math.round(((d - week1) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
    return `${d.getFullYear()}-W${String(weekNum).padStart(2,"0")}`;
  },

  getWeekDates(semanaId) {
    const [year, week] = semanaId.split("-W").map(Number);
    const jan4 = new Date(year, 0, 4);
    const startOfWeek = new Date(jan4);
    startOfWeek.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1 + (week - 1) * 7);
    return Array.from({length:5}, (_,i) => {
      const d = new Date(startOfWeek);
      d.setDate(startOfWeek.getDate() + i);
      return d;
    });
  },

  getUserName() {
    return this._account ? this._account.name : "Usuário";
  },

  getUserEmail() {
    return this._account ? this._account.username : "";
  }
};

// Exporta globalmente
window.SP = SP;
