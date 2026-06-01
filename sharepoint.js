// ============================================================
// sharepoint.js — Refeitório Homy · Microsoft Graph API
// ============================================================

const SP = {
  clientId: "aa37acf9-f3bd-4d1e-968a-fde57f79094c",
  clientSecret: "",
  appOnly: false,
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
    // Modo app-only: evita pedir acesso individual aos colaboradores.
    if (this.appOnly && this.clientSecret && this.clientSecret !== "COLE_AQUI_O_CLIENT_SECRET") {
      if (this._appToken && this._appTokenExpiresAt && Date.now() < this._appTokenExpiresAt) {
        return this._appToken;
      }

      const res = await fetch(`https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type: "client_credentials",
          client_id: this.clientId,
          client_secret: this.clientSecret,
          scope: "https://graph.microsoft.com/.default"
        })
      });

      if (!res.ok) {
        const err = await res.text();
        throw new Error(`Erro ao obter token app-only ${res.status}: ${err}`);
      }

      const data = await res.json();
      this._appToken = data.access_token;
      this._appTokenExpiresAt = Date.now() + ((data.expires_in || 3600) - 120) * 1000;
      return this._appToken;
    }

    // Fallback antigo com MSAL/login, caso appOnly esteja desligado.
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
  // Colunas: Nome, Departamento, Email, Ativo, tipo, Centro_Custo
  // ============================================================

  async getTodosColaboradores() {
    return this.getItems("Colaboradores");
  },

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

    if (dados.departamento !== undefined) fields.Departamento = dados.departamento;
    if (dados.email !== undefined) fields.Email = dados.email;
    if (dados.ativo !== undefined) fields.Ativo = dados.ativo;
    if (dados.Ativo !== undefined) fields.Ativo = dados.Ativo;
    if (dados.tipo !== undefined) fields.tipo = dados.tipo;
    if (dados.Centro_Custo !== undefined) fields.Centro_Custo = dados.Centro_Custo;
    if (dados.centroCusto !== undefined) fields.Centro_Custo = dados.centroCusto;

    return this.updateItem("Colaboradores", id, fields);
  },

  async desativarColaborador(id) {
    return this.updateItem("Colaboradores", id, { Ativo: false });
  },

  async deleteColaborador(id) {
    return this.deleteItem("Colaboradores", id);
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
  // Dia, Opcao, Nome_Prato, Confirmado, Data_Hora,
  // Centro_Custo, Status, Observacao, Origem, Alterado_Por
  // ============================================================
  async getPedidos(semanaId) {
    const items = await this.getItems("Pedidos");
    return items.filter(i => this.pick(i, "Semana_id") === semanaId);
  },

  async getPedidoColaborador(semanaId, colaboradorId) {
    const items = await this.getPedidos(semanaId);
    return items.filter(i => String(this.pick(i, "Colaborador_id")) === String(colaboradorId));
  },

  async savePedido(semanaId, colaboradorId, colaboradorNome, dia, opcao, nomePrato, dadosExtras = {}) {
    return this.createItem("Pedidos", {
      Title: `${semanaId}-${colaboradorId}-${dia}`,
      Semana_id: semanaId,
      Colaborador_id: String(colaboradorId),
      Colaborador_nome: colaboradorNome,
      Dia: dia,
      Opcao: opcao,
      Nome_Prato: nomePrato || "",
      Confirmado: dadosExtras.confirmado ?? dadosExtras.Confirmado ?? false,
      Data_Hora: dadosExtras.dataHora || dadosExtras.Data_Hora || new Date().toISOString(),
      Centro_Custo: dadosExtras.centroCusto || dadosExtras.Centro_Custo || "",
      Status: dadosExtras.status || dadosExtras.Status || "Confirmado",
      Observacao: dadosExtras.observacao || dadosExtras.Observacao || "",
      Origem: dadosExtras.origem || dadosExtras.Origem || "Refeitório",
      Alterado_Por: dadosExtras.alteradoPor || dadosExtras.Alterado_Por || this.getUserName()
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
    if (dados.dataHora !== undefined) fields.Data_Hora = dados.dataHora;

    if (dados.Centro_Custo !== undefined) fields.Centro_Custo = dados.Centro_Custo;
    if (dados.centroCusto !== undefined) fields.Centro_Custo = dados.centroCusto;

    if (dados.Status !== undefined) fields.Status = dados.Status;
    if (dados.status !== undefined) fields.Status = dados.status;

    if (dados.Observacao !== undefined) fields.Observacao = dados.Observacao;
    if (dados.observacao !== undefined) fields.Observacao = dados.observacao;

    if (dados.Origem !== undefined) fields.Origem = dados.Origem;
    if (dados.origem !== undefined) fields.Origem = dados.origem;

    if (dados.Alterado_Por !== undefined) fields.Alterado_Por = dados.Alterado_Por;
    if (dados.alteradoPor !== undefined) fields.Alterado_Por = dados.alteradoPor;

    return this.updateItem("Pedidos", id, fields);
  },

  async confirmarPedidos(semanaId, colaboradorId) {
    const pedidos = await this.getPedidoColaborador(semanaId, colaboradorId);

    for (const p of pedidos) {
      await this.updateItem("Pedidos", p.id, {
        Confirmado: true,
        Status: this.pick(p, "Status") || "Confirmado",
        Origem: this.pick(p, "Origem") || "Refeitório",
        Alterado_Por: this.getUserName()
      });
    }
  },

  async deletePedido(id) {
    return this.deleteItem("Pedidos", id);
  },

  // ============================================================
  // VALORES_REFEICAO
  // Lista: Valores_Refeicao
  // Colunas: Title, Data_Inicio, Data_Fim, Valor_Vascon,
  // Valor_Desconto_Funcionario, Observacao, Ativo
  // ============================================================
  async getValoresRefeicao(apenasAtivos = true) {
    const items = await this.getItems("Valores de Refeição");

    if (!apenasAtivos) return items;

    return items.filter(i => this.isTrue(this.pick(i, "Ativo")));
  },

  async getValorRefeicaoVigente(dataReferencia = new Date()) {
    const data = new Date(dataReferencia);
    const items = await this.getValoresRefeicao(true);

    return items.find(i => {
      const inicio = this.pick(i, "Data_Inicio") ? new Date(this.pick(i, "Data_Inicio")) : null;
      const fim = this.pick(i, "Data_Fim") ? new Date(this.pick(i, "Data_Fim")) : null;

      if (!inicio || !fim) return false;
      return data >= inicio && data <= fim;
    }) || null;
  },

  async createValorRefeicao(dados) {
    return this.createItem("Valores de Refeição", {
      Title: dados.title || dados.Title || "Valor refeição",
      Data_Inicio: dados.dataInicio || dados.Data_Inicio,
      Data_Fim: dados.dataFim || dados.Data_Fim,
      Valor_Vascon: Number(dados.valorVascon ?? dados.Valor_Vascon ?? 0),
      Valor_Desconto_Funcionario: Number(dados.valorDescontoFuncionario ?? dados.Valor_Desconto_Funcionario ?? 0),
      Observacao: dados.observacao || dados.Observacao || "",
      Ativo: dados.ativo ?? dados.Ativo ?? true
    });
  },

  async updateValorRefeicao(id, dados) {
    const fields = {};

    if (dados.title !== undefined) fields.Title = dados.title;
    if (dados.Title !== undefined) fields.Title = dados.Title;

    if (dados.dataInicio !== undefined) fields.Data_Inicio = dados.dataInicio;
    if (dados.Data_Inicio !== undefined) fields.Data_Inicio = dados.Data_Inicio;

    if (dados.dataFim !== undefined) fields.Data_Fim = dados.dataFim;
    if (dados.Data_Fim !== undefined) fields.Data_Fim = dados.Data_Fim;

    if (dados.valorVascon !== undefined) fields.Valor_Vascon = Number(dados.valorVascon);
    if (dados.Valor_Vascon !== undefined) fields.Valor_Vascon = Number(dados.Valor_Vascon);

    if (dados.valorDescontoFuncionario !== undefined) fields.Valor_Desconto_Funcionario = Number(dados.valorDescontoFuncionario);
    if (dados.Valor_Desconto_Funcionario !== undefined) fields.Valor_Desconto_Funcionario = Number(dados.Valor_Desconto_Funcionario);

    if (dados.observacao !== undefined) fields.Observacao = dados.observacao;
    if (dados.Observacao !== undefined) fields.Observacao = dados.Observacao;

    if (dados.ativo !== undefined) fields.Ativo = dados.ativo;
    if (dados.Ativo !== undefined) fields.Ativo = dados.Ativo;

    return this.updateItem("Valores de Refeição", id, fields);
  },

  // ============================================================
  // AUSENCIAS DO REFEITORIO
  // Lista: Ausencias_Refeitorio
  // Colunas: Title, Colaborador_id, Colaborador_nome, Data_Inicio,
  // Data_Fim, Motivo, Observacao, Ativo, Criado_Por
  // ============================================================
  async getAusencias(apenasAtivas = true) {
    const items = await this.getItems("Ausencias do Refeitorio");

    if (!apenasAtivas) return items;

    return items.filter(i => this.isTrue(this.pick(i, "Ativo")));
  },

  async getAusenciasColaborador(colaboradorId, dataReferencia = null) {
    const items = await this.getAusencias(true);

    return items.filter(i => {
      const mesmoColaborador = String(this.pick(i, "Colaborador_id")) === String(colaboradorId);
      if (!mesmoColaborador) return false;

      if (!dataReferencia) return true;

      const data = new Date(dataReferencia);
      const inicio = this.pick(i, "Data_Inicio") ? new Date(this.pick(i, "Data_Inicio")) : null;
      const fim = this.pick(i, "Data_Fim") ? new Date(this.pick(i, "Data_Fim")) : null;

      if (!inicio || !fim) return false;
      return data >= inicio && data <= fim;
    });
  },

  async colaboradorEstaAusente(colaboradorId, dataReferencia = new Date()) {
    const ausencias = await this.getAusenciasColaborador(colaboradorId, dataReferencia);
    return ausencias.length > 0 ? ausencias[0] : null;
  },

  async createAusencia(dados) {
    const colaboradorNome = dados.colaboradorNome || dados.Colaborador_nome || "";
    const motivo = dados.motivo || dados.Motivo || "Ausência";

    return this.createItem("Ausencias do Refeitorio", {
      Title: dados.title || dados.Title || `${colaboradorNome} - ${motivo}`,
      Colaborador_id: String(dados.colaboradorId || dados.Colaborador_id || ""),
      Colaborador_nome: colaboradorNome,
      Data_Inicio: dados.dataInicio || dados.Data_Inicio,
      Data_Fim: dados.dataFim || dados.Data_Fim,
      Motivo: motivo,
      Observacao: dados.observacao || dados.Observacao || "",
      Ativo: dados.ativo ?? dados.Ativo ?? true,
      Criado_Por: dados.criadoPor || dados.Criado_Por || this.getUserName()
    });
  },

  async updateAusencia(id, dados) {
    const fields = {};

    if (dados.title !== undefined) fields.Title = dados.title;
    if (dados.Title !== undefined) fields.Title = dados.Title;

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

    if (dados.criadoPor !== undefined) fields.Criado_Por = dados.criadoPor;
    if (dados.Criado_Por !== undefined) fields.Criado_Por = dados.Criado_Por;

    return this.updateItem("Ausencias do Refeitorio", id, fields);
  },

  async deleteAusencia(id) {
    return this.deleteItem("Ausencias do Refeitorio", id);
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


  async saveCheckIn(semanaId, colaboradorId, colaboradorNome, dia, confirmadoPor) {
    return this.registrarCheckIn(semanaId, colaboradorId, colaboradorNome, dia, confirmadoPor);
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
  // ALIASES DE COMPATIBILIDADE
  // Mantém compatibilidade com funções antigas usadas no admin/index.html
  // ============================================================
  async addItem(listName, fields) {
    return this.createItem(listName, fields);
  },

  async patchListItem(listName, itemId, fields) {
    return this.updateItem(listName, itemId, fields);
  },

  async removeItem(listName, itemId) {
    return this.deleteItem(listName, itemId);
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
  },

  // ============================================================
  // VALORES DE REFEIÇÃO
  // ============================================================

  async getValoresRefeicao() {
    const items = await this.getItems("Valores de Refeição");
    return items;
  },

  async createValorRefeicao(dados) {
    return this.createItem("Valores de Refeição", {
      Title: dados.title || "",
      Data_Inicio: dados.dataInicio,
      Data_Fim: dados.dataFim,
      Valor_Vascon: Number(dados.valorVascon || 0),
      Valor_Desconto_Funcionario: Number(dados.valorDesconto || 0),
      Observacao: dados.observacao || "",
      Ativo: dados.ativo !== false
    });
  },

  async updateValorRefeicao(id, dados) {
    const fields = {};

    if (dados.title !== undefined) fields.Title = dados.title;
    if (dados.dataInicio !== undefined) fields.Data_Inicio = dados.dataInicio;
    if (dados.dataFim !== undefined) fields.Data_Fim = dados.dataFim;

    if (dados.valorVascon !== undefined)
      fields.Valor_Vascon = Number(dados.valorVascon);

    if (dados.valorDesconto !== undefined)
      fields.Valor_Desconto_Funcionario = Number(dados.valorDesconto);

    if (dados.observacao !== undefined)
      fields.Observacao = dados.observacao;

    if (dados.ativo !== undefined)
      fields.Ativo = dados.ativo;

    return this.updateItem("Valores de Refeição", id, fields);
  }
,

  // ============================================================
  // CHECK-IN COZINHA
  // Lista: CheckIn
  // Colunas sugeridas: Title, PedidoId, NomeColaborador,
  // OpcaoEscolhida, Dia, DataHora, Confirmado
  // ============================================================
  async getCheckIns() {
    return this.getItems("CheckIn");
  },

  async createCheckIn(dados) {
    return this.createItem("CheckIn", {
      Title: String(dados.pedidoId || dados.PedidoId || ""),
      PedidoId: String(dados.pedidoId || dados.PedidoId || ""),
      NomeColaborador: dados.nomeColaborador || dados.NomeColaborador || "",
      OpcaoEscolhida: dados.opcaoEscolhida || dados.OpcaoEscolhida || "",
      Dia: dados.dia || dados.Dia || "",
      DataHora: dados.dataHora || dados.DataHora || new Date().toISOString(),
      Confirmado: dados.confirmado ?? dados.Confirmado ?? true
    });
  }


};

window.SP = SP;
