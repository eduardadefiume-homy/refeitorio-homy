// ============================================================
// sharepoint.js — Refeitório Homy · Microsoft Graph API
// ============================================================

const SP = {
  clientId: "2e73e8be-484b-428e-a6c0-c75b1bf70d8a",
  tenantId: "a2850abc-334a-4805-b6b2-420b4aef68a9",
  siteUrl: "homyquimica.sharepoint.com",
  sitePath: "/sites/Refeitrio-Homy",
  scopes: ["https://graph.microsoft.com/Sites.ReadWrite.All", "User.Read"],

  _msalInstance: null,
  _account: null,
  _siteId: null,
  _listIds: {},

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

  async init() {
    if (this._msalInstance) return true;

    const msalConfig = {
      auth: {
        clientId: this.clientId,
        authority: `https://login.microsoftonline.com/${this.tenantId}`,
        redirectUri: window.location.origin + window.location.pathname
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: true
      }
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

  async login() {
    await this.init();

    const result = await this._msalInstance.loginPopup({
      scopes: this.scopes,
      prompt: "select_account"
    });

    this._account = result.account;
    return true;
  },

  async getToken() {
    await this.init();

    if (!this._account) {
      await this.login();
    }

    try {
      const result = await this._msalInstance.acquireTokenSilent({
        scopes: this.scopes,
        account: this._account
      });

      return result.accessToken;
    } catch (e) {
      const result = await this._msalInstance.acquireTokenPopup({
        scopes: this.scopes
      });

      this._account = result.account || this._account;
      return result.accessToken;
    }
  },

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

  async getSiteId() {
    if (this._siteId) return this._siteId;

    const data = await this.graph("GET", `/sites/${this.siteUrl}:${this.sitePath}`);
    this._siteId = data.id;
    return this._siteId;
  },

  async getListId(listName) {
    if (this._listIds[listName]) return this._listIds[listName];

    const siteId = await this.getSiteId();
    const data = await this.graph("GET", `/sites/${siteId}/lists?$filter=displayName eq '${listName}'`);

    if (!data.value || data.value.length === 0) {
      throw new Error(`Lista não encontrada no SharePoint: ${listName}`);
    }

    this._listIds[listName] = data.value[0].id;
    return this._listIds[listName];
  },

  async getItems(listName) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const data = await this.graph("GET", `/sites/${siteId}/lists/${listId}/items?expand=fields`);
    return (data.value || []).map(i => ({ id: i.id, ...i.fields }));
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
  // Colunas: Nome, Departamento, Email, Ativo, tipo
  // ============================================================
  async getColaboradores() {
    const items = await this.getItems("Colaboradores");
    return items.filter(i => this.isTrue(this.pick(i, "Ativo")));
  },

  async createColaborador(dados) {
    return this.createItem("Colaboradores", {
      Title: dados.nome,
      Nome: dados.nome,
      Departamento: dados.departamento || "",
      Email: dados.email || "",
      Ativo: true,
      tipo: dados.tipo || "colaborador"
    });
  },

  async updateColaborador(id, dados) {
    const fields = {};

    if (dados.nome !== undefined) {
      fields.Title = dados.nome;
      fields.Nome = dados.nome;
    }

    if (dados.departamento !== undefined) fields.Departamento = dados.departamento;
    if (dados.email !== undefined) fields.Email = dados.email;
    if (dados.ativo !== undefined) fields.Ativo = dados.ativo;
    if (dados.Ativo !== undefined) fields.Ativo = dados.Ativo;
    if (dados.tipo !== undefined) fields.tipo = dados.tipo;

    return this.updateItem("Colaboradores", id, fields);
  },

  async desativarColaborador(id) {
    return this.updateItem("Colaboradores", id, { Ativo: false });
  },

  // ============================================================
  // CARDAPIO
  // Colunas: Semana_id, Dia, Opcao, Nome_Prato, Detalhes
  // ============================================================
  async getCardapio(semanaId) {
    const items = await this.getItems("Cardapio");
    return items.filter(i => this.pick(i, "Semana_id") === semanaId);
  },

  async saveCardapio(semanaId, dia, opcao, nomePrato, detalhes) {
    return this.createItem("Cardapio", {
      Title: `${semanaId}-${dia}-${opcao}`,
      Semana_id: semanaId,
      Dia: dia,
      Opcao: opcao,
      Nome_Prato: nomePrato,
      Detalhes: detalhes || ""
    });
  },

  async clearCardapio(semanaId) {
    const items = await this.getCardapio(semanaId);
    for (const item of items) {
      await this.deleteItem("Cardapio", item.id);
    }
  },

  // ============================================================
  // PEDIDOS
  // Colunas: Semana_id, Colaborador_id, Colaborador_nome,
  // Dia, Opcao, Nome_Prato, Confirmado, Data_Hora
  // ============================================================
  async getPedidos(semanaId) {
    const items = await this.getItems("Pedidos");
    return items.filter(i => this.pick(i, "Semana_id") === semanaId);
  },

  async getPedidoColaborador(semanaId, colaboradorId) {
    const items = await this.getPedidos(semanaId);
    return items.filter(i => String(this.pick(i, "Colaborador_id")) === String(colaboradorId));
  },

  async savePedido(semanaId, colaboradorId, colaboradorNome, dia, opcao, nomePrato) {
    return this.createItem("Pedidos", {
      Title: `${semanaId}-${colaboradorId}-${dia}`,
      Semana_id: semanaId,
      Colaborador_id: String(colaboradorId),
      Colaborador_nome: colaboradorNome,
      Dia: dia,
      Opcao: opcao,
      Nome_Prato: nomePrato || "",
      Confirmado: false,
      Data_Hora: new Date().toISOString()
    });
  },

  async updatePedido(id, dados) {
    const fields = {};

    if (dados.Semana_id !== undefined) fields.Semana_id = dados.Semana_id;
    if (dados.semanaId !== undefined) fields.Semana_id = dados.semanaId;

    if (dados.Colaborador_id !== undefined) fields.Colaborador_id = String(dados.Colaborador_id);
    if (dados.colaboradorId !== undefined) fields.Colaborador_id = String(dados.colaboradorId);

    if (dados.Colaborador_nome !== undefined) fields.Colaborador_nome = dados.Colaborador_nome;
    if (dados.colaboradorNome !== undefined) fields.Colaborador_nome = dados.colaboradorNome;

    if (dados.Dia !== undefined) fields.Dia = dados.Dia;
    if (dados.dia !== undefined) fields.Dia = dados.dia;

    if (dados.Opcao !== undefined) fields.Opcao = dados.Opcao;
    if (dados.opcao !== undefined) fields.Opcao = dados.opcao;

    if (dados.Nome_Prato !== undefined) fields.Nome_Prato = dados.Nome_Prato;
    if (dados.nomePrato !== undefined) fields.Nome_Prato = dados.nomePrato;

    if (dados.Confirmado !== undefined) fields.Confirmado = dados.Confirmado;
    if (dados.confirmado !== undefined) fields.Confirmado = dados.confirmado;

    if (dados.Data_Hora !== undefined) fields.Data_Hora = dados.Data_Hora;

    return this.updateItem("Pedidos", id, fields);
  },

  async confirmarPedidos(semanaId, colaboradorId) {
    const pedidos = await this.getPedidoColaborador(semanaId, colaboradorId);

    for (const p of pedidos) {
      await this.updateItem("Pedidos", p.id, { Confirmado: true });
    }
  },

  async deletePedido(id) {
    return this.deleteItem("Pedidos", id);
  },

  // ============================================================
  // CONFIGURAÇÕES
  // Colunas: Chave, Valor, Descrição
  // OBS: para evitar erro de nome interno com acento,
  // este código grava somente Title, Chave e Valor.
  // ============================================================
  async getConfig(chave) {
    const items = await this.getItems("Configurações");

    const item = items.find(i =>
      this.pick(i, "Chave") === chave ||
      this.pick(i, "Title") === chave
    );

    return item ? this.pick(item, "Valor") : null;
  },

  async setConfig(chave, valor, descricao = "") {
    const items = await this.getItems("Configurações");

    const existing = items.find(i =>
      this.pick(i, "Chave") === chave ||
      this.pick(i, "Title") === chave
    );

    if (existing) {
      return this.updateItem("Configurações", existing.id, { Valor: valor });
    }

    return this.createItem("Configurações", {
      Title: chave,
      Chave: chave,
      Valor: valor
    });
  },

  async isCardapioLiberado() {
    const possiveisChaves = [
      "cardapio_liberado",
      "marcacao_liberada",
      "pedidos_liberados"
    ];

    for (const chave of possiveisChaves) {
      const valor = await this.getConfig(chave);
      if (this.isTrue(valor)) return true;
    }

    return false;
  },

  async setMarcacaoLiberada(liberado) {
    const valor = liberado ? "sim" : "nao";

    await this.setConfig("cardapio_liberado", valor);
    await this.setConfig("marcacao_liberada", valor);
    await this.setConfig("pedidos_liberados", valor);

    return true;
  },

  async getPrazoMarcacao() {
    return this.getConfig("prazo_limite");
  },

  async setPrazoMarcacao(valor) {
    return this.setConfig("prazo_limite", valor);
  },

  // ============================================================
  // CHECKIN
  // Colunas: Semana_id, Colaborador_id, Colaborador_nome,
  // Dia, Retirou, Data_Hora_Retirada, Confirmado_Por
  // ============================================================
  async getCheckIn(semanaId, dia) {
    const items = await this.getItems("CheckIn");
    return items.filter(i =>
      this.pick(i, "Semana_id") === semanaId &&
      this.pick(i, "Dia") === dia
    );
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

  // ============================================================
  // EXTRAS
  // Colunas: Semana_id, Dia, Nome, tipo, Opcao,
  // Observacao, Adicionado_Por
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

  async removeExtra(id) {
    return this.deleteItem("Extras", id);
  },

  // ============================================================
  // UTILITÁRIOS
  // ============================================================
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

  getWeekDates(semanaId) {
    const [year, week] = semanaId.split("-W").map(Number);

    const jan4 = new Date(year, 0, 4);
    const startOfWeek = new Date(jan4);

    startOfWeek.setDate(
      jan4.getDate() - (jan4.getDay() || 7) + 1 + (week - 1) * 7
    );

    return Array.from({ length: 5 }, (_, i) => {
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

window.SP = SP;
