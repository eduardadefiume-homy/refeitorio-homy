// ============================================================
// sharepoint.js — Refeitório Homy · Microsoft Graph API
// ============================================================

var SP = window.SP = {
  clientId: "aa37acf9-f3bd-4d1e-968a-fde57f79094c",
  clientSecret: "",
  appOnly: false,
  tenantId: "a2850abc-334a-4805-b6b2-420b4aef68a9",
  siteUrl: "homyquimica.sharepoint.com",
  sitePath: "/sites/Refeitorio-Homy",
  scopes: ["Sites.ReadWrite.All", "User.Read"],

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

    if (!window.msal) {
      throw new Error("Biblioteca MSAL não carregou. Verifique se o script msal-browser.min.js está antes do sharepoint.js.");
    }

    const msalConfig = {
      auth: {
        clientId: this.clientId,
        authority: `https://login.microsoftonline.com/${this.tenantId}`,
        redirectUri: window.location.origin + window.location.pathname,
        navigateToLoginRequestUrl: false
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: true
      }
    };

    this._msalInstance = new msal.PublicClientApplication(msalConfig);
    await this._msalInstance.initialize();

    // Igual ao Ramais Homy: trata o retorno do login e reaproveita sessão.
    const redirectResult = await this._msalInstance.handleRedirectPromise();
    if (redirectResult && redirectResult.account) {
      this._account = redirectResult.account;
      this._msalInstance.setActiveAccount(this._account);
      return true;
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
  },

  async login() {
    await this.init();

    // Login por redirect, igual Ramais: mais estável no GitHub Pages que popup.
    await this._msalInstance.loginRedirect({
      scopes: this.scopes
    });

    return false;
  },

  async getToken() {
    await this.init();

    if (!this._account) {
      await this.login();
      return null;
    }

    try {
      const result = await this._msalInstance.acquireTokenSilent({
        scopes: this.scopes,
        account: this._account
      });

      return result.accessToken;
    } catch (e) {
      console.warn("[SP] Token silencioso falhou; redirecionando para login.", e);

      await this._msalInstance.acquireTokenRedirect({
        scopes: this.scopes,
        account: this._account
      });

      return null;
    }
  },

  async ensureLogin() {
    await this.init();

    if (!this._account) {
      await this.login();
      return false;
    }

    return true;
  },

  getUserName() {
    return this._account?.name || this._account?.username || "Usuário Homy";
  },

  async logout() {
    await this.init();
    const account = this._account;
    this._account = null;
    if (account) {
      await this._msalInstance.logoutRedirect({ account });
    }
  },

  async graph(method, endpoint, body = null) {
    const token = await this.getToken();
    if (!token) return null;

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

    const paths = [
      this.sitePath,
      "/sites/Refeitorio-Homy",
      "/sites/Refeitrio-Homy"
    ].filter((value, index, array) => value && array.indexOf(value) === index);

    let lastError = null;

    for (const path of paths) {
      try {
        const data = await this.graph("GET", `/sites/${this.siteUrl}:${path}`);
        if (data && data.id) {
          this._siteId = data.id;
          this.sitePath = path;
          return this._siteId;
        }
      } catch (e) {
        lastError = e;
      }
    }

    throw lastError || new Error("Site do SharePoint não encontrado.");
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
    let endpoint = `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=999`;
    const items = [];

    while (endpoint) {
      const data = await this.graph("GET", endpoint);
      items.push(...(data.value || []));
      endpoint = data["@odata.nextLink"] ? data["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "") : null;
    }

    return items.map(i => ({ id: i.id, ...i.fields }));
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
// ============================================================
// HOTFIX HOMY 2026-06-01 — regras finais de dados
// ============================================================
(function(){
  if(!window.SP) return;
  const SP = window.SP;

  SP.normalizeText = function(v){
    return String(v ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toLowerCase();
  };

  SP.getCurrentWeekId = function(date = new Date()){
    return this.getSemanaId ? this.getSemanaId(date) : (()=>{
      const d = new Date(date); d.setHours(0,0,0,0);
      d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
      const w1 = new Date(d.getFullYear(),0,4);
      const week = 1 + Math.round(((d-w1)/86400000 - 3 + ((w1.getDay()+6)%7))/7);
      return `${d.getFullYear()}-W${String(week).padStart(2,"0")}`;
    })();
  };

  SP.dateOnly = function(date){
    const d = new Date(date);
    if(isNaN(d.getTime())) return "";
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
  };

  SP.getDataRefBySemanaDia = function(semanaId, dia){
    const mapa = {segunda:0,terca:1,terça:1,quarta:2,quinta:3,sexta:4};
    const idx = mapa[this.normalizeText(dia)];
    if(idx === undefined || !semanaId) return this.dateOnly(new Date());
    const datas = this.getWeekDates ? this.getWeekDates(semanaId) : [];
    return datas[idx] ? this.dateOnly(datas[idx]) : this.dateOnly(new Date());
  };

  const _savePedidoBase = SP.savePedido;
  SP.savePedido = async function(semanaId, colaboradorId, colaboradorNome, dia, opcao, nomePrato, dadosExtras = {}){
    const extra = Object.assign({}, dadosExtras);
    extra.Data_Ref = extra.Data_Ref || extra.dataRef || this.getDataRefBySemanaDia(semanaId, dia);
    const item = await this.createItem("Pedidos", {
      Title: `${semanaId}-${colaboradorId}-${dia}`,
      Semana_id: semanaId,
      Colaborador_id: String(colaboradorId),
      Colaborador_nome: colaboradorNome,
      Dia: dia,
      Opcao: opcao,
      Nome_Prato: nomePrato || "",
      Confirmado: extra.confirmado ?? extra.Confirmado ?? false,
      Data_Hora: extra.dataHora || extra.Data_Hora || new Date().toISOString(),
      Data_Ref: extra.Data_Ref,
      Centro_Custo: extra.centroCusto || extra.Centro_Custo || "",
      Status: extra.status || extra.Status || "Confirmado",
      Observacao: extra.observacao || extra.Observacao || "",
      Origem: extra.origem || extra.Origem || "Refeitório",
      Alterado_Por: extra.alteradoPor || extra.Alterado_Por || this.getUserName()
    });
    return item;
  };

  SP.getPedidosPorPeriodo = async function(dataInicio, dataFim){
    const items = await this.getItems("Pedidos");
    return items.filter(p=>{
      let d = this.pick(p,"Data_Ref","DataRef","data_ref");
      if(!d){
        const semana = this.pick(p,"Semana_id","semana_id");
        const dia = this.pick(p,"Dia","dia");
        d = this.getDataRefBySemanaDia(semana,dia);
      }
      d = String(d || "").slice(0,10);
      return d && d >= dataInicio && d <= dataFim;
    });
  };

  SP.isConfigOn = function(valor){
    const v = this.normalizeText(valor);
    return ["sim","true","1","yes","ativo","liberado"].includes(v);
  };

  SP.isConfigOff = function(valor){
    const v = this.normalizeText(valor);
    return ["nao","não","false","0","no","inativo","bloqueado"].includes(v);
  };

  SP.getMarcacaoLiberada = async function(){
    const chaves = ["marcacao_liberada","pedidos_liberados","cardapio_liberado"];
    const vals = [];
    for(const c of chaves){ vals.push(await this.getConfig(c)); }
    if(vals.some(v=>this.isConfigOff(v))) return false;
    if(vals.some(v=>this.isConfigOn(v))) return true;
    return true;
  };

  SP.isCardapioLiberado = async function(){
    return this.getMarcacaoLiberada();
  };

  SP.getCardapioVisivel = async function(){
    const vals = [await this.getConfig("cardapio_visivel"), await this.getConfig("cardapio_liberado")];
    if(vals.some(v=>this.isConfigOff(v))) return false;
    if(vals.some(v=>this.isConfigOn(v))) return true;
    return true;
  };

  SP.setCardapioVisivel = async function(visivel){
    const valor = visivel ? "sim" : "nao";
    await this.setConfig("cardapio_visivel", valor);
    return true;
  };

  SP.setMarcacaoLiberada = async function(liberado){
    const valor = liberado ? "sim" : "nao";
    await this.setConfig("marcacao_liberada", valor);
    await this.setConfig("pedidos_liberados", valor);
    await this.setConfig("cardapio_liberado", valor);
    await this.setConfig("sync_timestamp", new Date().toISOString());
    return true;
  };

  SP.getPrazoMarcacao = async function(){
    const chaves = ["prazo_limite","prazo_marcacao","data_limite_marcacao","horario_limite_marcacao"];
    for(const c of chaves){
      const v = await this.getConfig(c);
      if(v) return v;
    }
    return null;
  };

  SP.setPrazoMarcacao = async function(valor){
    await this.setConfig("prazo_limite", valor);
    await this.setConfig("prazo_marcacao", valor);
    await this.setConfig("sync_timestamp", new Date().toISOString());
    return true;
  };

  SP.addExtraPedido = async function(semanaId, dia, nome, tipo, opcao="principal", observacao="", adicionadoPor=null){
    const normalDia = this.normalizeText(dia);
    const normalNome = this.normalizeText(nome);
    const pedidos = await this.getPedidos(semanaId);
    const existente = pedidos.find(p =>
      this.normalizeText(this.pick(p,"Colaborador_nome","Title")) === normalNome &&
      this.normalizeText(this.pick(p,"Dia")) === normalDia &&
      ["extra","extra automatica","extra automática","investigador","guarda","prestador","visitante","motorista","marmita","outro"].includes(this.normalizeText(this.pick(p,"Origem","Tipo","tipo")))
    );
    if(existente) return existente;

    let extraItem = null;
    try{
      extraItem = await this.createItem("Extras", {
        Title: `${semanaId}-${dia}-${nome}`,
        Semana_id: semanaId,
        Dia: dia,
        Nome: nome,
        tipo: tipo,
        Opcao: opcao,
        Observacao: observacao || "",
        Adicionado_Por: adicionadoPor || this.getUserName()
      });
    }catch(e){ console.warn("Não salvou em Extras, seguirá criando pedido:", e); }

    return this.savePedido(
      semanaId,
      `EXTRA-${extraItem?.id || Date.now()}`,
      nome,
      dia,
      opcao,
      observacao || nome,
      {
        confirmado:true,
        status:"Confirmado",
        origem: tipo || "Extra",
        observacao,
        dataRef:this.getDataRefBySemanaDia(semanaId,dia),
        alteradoPor: adicionadoPor || this.getUserName()
      }
    );
  };

  SP.ensureExtraAutomaticoSemana = async function(semanaId){
    const dias = ["segunda","terca","quarta","quinta","sexta"];
    for(const dia of dias){
      await this.addExtraPedido(semanaId,dia,"Refeição extra","extra automatica","principal","Refeição extra automática");
    }
    await this.setConfig("sync_timestamp", new Date().toISOString());
    return true;
  };
})();


// ============================================================
// HOTFIX OPERAÇÃO HOMY 2026-06-01
// Corrige: valores, extras automáticos, extras separados, sync.
// ============================================================
(function(){
  if(!window.SP) return;
  const SP = window.SP;

  function money(v){
    const n = Number(String(v ?? 0).replace(/[R$\s.]/g,'').replace(',','.'));
    return Number.isFinite(n) ? n : 0;
  }
  function norm(v){
    return String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase();
  }

  SP.isExtraPedido = function(p){
    const origem = norm(this.pick(p,'Origem','origem','Tipo','tipo'));
    const nome = norm(this.pick(p,'Colaborador_nome','Nome','Title'));
    return origem.includes('extra') || ['investigador','guarda','prestador','visitante','motorista','marmita','outro'].includes(origem) || nome.includes('refeicao extra') || nome.includes('investigador') || nome.includes('guarda');
  };

  SP.createValorRefeicao = async function(dados){
    const fields = {
      Title: dados.title || dados.Title || 'Valor refeição',
      Data_Inicio: dados.dataInicio || dados.Data_Inicio,
      Data_Fim: dados.dataFim || dados.Data_Fim,
      Valor_Vascon: money(dados.valorVascon ?? dados.Valor_Vascon),
      Valor_Desconto: money(dados.valorDesconto ?? dados.valorDescontoFuncionario ?? dados.Valor_Desconto ?? dados.Valor_Desconto_Funcionario),
      Observacao: dados.observacao || dados.Observacao || '',
      Ativo: dados.ativo ?? dados.Ativo ?? true
    };
    try{
      return await this.createItem('Valores de Refeição', fields);
    }catch(e){
      // Compatibilidade caso o ambiente tenha a coluna antiga.
      if(String(e.message||e).includes('Valor_Desconto')){
        delete fields.Valor_Desconto;
        fields.Valor_Desconto_Funcionario = money(dados.valorDesconto ?? dados.valorDescontoFuncionario ?? dados.Valor_Desconto ?? dados.Valor_Desconto_Funcionario);
        return await this.createItem('Valores de Refeição', fields);
      }
      throw e;
    }
  };

  SP.updateValorRefeicao = async function(id,dados){
    const fields = {};
    if(dados.title!==undefined) fields.Title=dados.title;
    if(dados.dataInicio!==undefined) fields.Data_Inicio=dados.dataInicio;
    if(dados.dataFim!==undefined) fields.Data_Fim=dados.dataFim;
    if(dados.valorVascon!==undefined) fields.Valor_Vascon=money(dados.valorVascon);
    if(dados.valorDesconto!==undefined) fields.Valor_Desconto=money(dados.valorDesconto);
    if(dados.observacao!==undefined) fields.Observacao=dados.observacao;
    if(dados.ativo!==undefined) fields.Ativo=dados.ativo;
    try{return await this.updateItem('Valores de Refeição', id, fields);}catch(e){
      if(String(e.message||e).includes('Valor_Desconto')){ fields.Valor_Desconto_Funcionario=fields.Valor_Desconto; delete fields.Valor_Desconto; return await this.updateItem('Valores de Refeição', id, fields); }
      throw e;
    }
  };

  SP.addExtraPedido = async function(semanaId, dia, nome, tipo, opcao='principal', observacao='', adicionadoPor=null){
    const normalDia=norm(dia), normalNome=norm(nome), normalTipo=norm(tipo);
    const pedidos = await this.getPedidos(semanaId);
    const existente = pedidos.find(p => norm(this.pick(p,'Dia'))===normalDia && norm(this.pick(p,'Colaborador_nome','Title','Nome'))===normalNome && norm(this.pick(p,'Origem','Tipo','tipo'))===normalTipo);
    if(existente) return existente;

    let extraItem=null;
    try{
      extraItem = await this.createItem('Extras',{
        Title:`${semanaId}-${dia}-${nome}-${Date.now()}`,
        Semana_id:semanaId, Dia:dia, Nome:nome, tipo:tipo, Opcao:opcao,
        Observacao:observacao||'', Adicionado_Por:adicionadoPor||this.getUserName()
      });
    }catch(e){console.warn('[SP] Extra não salvo em Extras:',e);}

    return await this.savePedido(semanaId, `EXTRA-${extraItem?.id || Date.now()}`, nome, dia, opcao, observacao || nome, {
      confirmado:true, status:'Confirmado', origem:tipo||'Extra', observacao,
      dataRef:this.getDataRefBySemanaDia ? this.getDataRefBySemanaDia(semanaId,dia) : null,
      alteradoPor:adicionadoPor||this.getUserName()
    });
  };

  SP.ensureExtraAutomaticoSemana = async function(semanaId){
    const dias=['segunda','terca','quarta','quinta','sexta'];
    const pedidos=await this.getPedidos(semanaId);
    for(const dia of dias){
      const existe=pedidos.some(p => norm(this.pick(p,'Dia'))===dia && norm(this.pick(p,'Colaborador_nome','Title','Nome')).includes('refeicao extra') && norm(this.pick(p,'Origem','Tipo','tipo')).includes('extra'));
      if(!existe) await this.addExtraPedido(semanaId,dia,'Refeição extra','extra automatica','principal','Refeição extra automática');
    }
    await this.setConfig('sync_timestamp', new Date().toISOString());
    return true;
  };

  SP.getDashboardResumo = async function(semanaId){
    const [colabs,pedidos,extras,checkins] = await Promise.all([
      this.getColaboradores(), this.getPedidos(semanaId), this.getExtras ? this.getExtras(semanaId) : Promise.resolve([]), this.getCheckIn ? this.getCheckIn(semanaId) : Promise.resolve([])
    ]);
    const pedidosColab = pedidos.filter(p=>!this.isExtraPedido(p));
    const pedidosExtra = pedidos.filter(p=>this.isExtraPedido(p));
    return {
      colaboradoresAtivos: colabs.length,
      pedidosConfirmadosColaboradores: pedidosColab.filter(p=>this.isTrue(this.pick(p,'Confirmado')) || norm(this.pick(p,'Status'))==='confirmado').length,
      pendentesColaboradores: Math.max(0, colabs.length - pedidosColab.length),
      extrasAtivos: extras.length,
      extrasConfirmados: pedidosExtra.length,
      totalPedidosHoje: pedidos.length,
      checkinsHoje: checkins.length
    };
  };
})();


// ============================================================
// HOTFIX FINAL OPERAÇÃO — valores, extras e sincronização
// ============================================================
(function(){
  if(!window.SP) return;
  const SP = window.SP;
  const norm = v => String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase();
  const money = v => { const n=Number(String(v ?? 0).replace(/[R$\s.]/g,'').replace(',','.')); return Number.isFinite(n)?n:0; };
  const pick = (obj,...keys)=>{ for(const k of keys){ if(obj && obj[k]!==undefined && obj[k]!==null) return obj[k]; } return ''; };

  // Corrige a coluna real da lista Valores de Refeição: Valor_Desconto.
  SP.createValorRefeicao = async function(dados){
    const fields = {
      Title: dados.title || dados.Title || 'Valor refeição',
      Data_Inicio: dados.dataInicio || dados.Data_Inicio,
      Data_Fim: dados.dataFim || dados.Data_Fim,
      Valor_Vascon: money(dados.valorVascon ?? dados.Valor_Vascon),
      Valor_Desconto: money(dados.valorDesconto ?? dados.valorDescontoFuncionario ?? dados.Valor_Desconto),
      Observacao: dados.observacao || dados.Observacao || '',
      Ativo: dados.ativo ?? dados.Ativo ?? true
    };
    return await this.createItem('Valores de Refeição', fields);
  };

  SP.updateValorRefeicao = async function(id,dados){
    const fields = {};
    if(dados.title!==undefined) fields.Title=dados.title;
    if(dados.Title!==undefined) fields.Title=dados.Title;
    if(dados.dataInicio!==undefined) fields.Data_Inicio=dados.dataInicio;
    if(dados.Data_Inicio!==undefined) fields.Data_Inicio=dados.Data_Inicio;
    if(dados.dataFim!==undefined) fields.Data_Fim=dados.dataFim;
    if(dados.Data_Fim!==undefined) fields.Data_Fim=dados.Data_Fim;
    if(dados.valorVascon!==undefined) fields.Valor_Vascon=money(dados.valorVascon);
    if(dados.Valor_Vascon!==undefined) fields.Valor_Vascon=money(dados.Valor_Vascon);
    if(dados.valorDesconto!==undefined) fields.Valor_Desconto=money(dados.valorDesconto);
    if(dados.Valor_Desconto!==undefined) fields.Valor_Desconto=money(dados.Valor_Desconto);
    if(dados.observacao!==undefined) fields.Observacao=dados.observacao;
    if(dados.Observacao!==undefined) fields.Observacao=dados.Observacao;
    if(dados.ativo!==undefined) fields.Ativo=dados.ativo;
    if(dados.Ativo!==undefined) fields.Ativo=dados.Ativo;
    return await this.updateItem('Valores de Refeição', id, fields);
  };

  SP.isExtraPedido = function(p){
    const origem = norm(pick(p,'Origem','origem','Tipo','tipo'));
    const nome = norm(pick(p,'Colaborador_nome','Nome','Title'));
    return origem.includes('extra') || ['investigador','guarda','prestador','visitante','motorista','marmita','outro'].includes(origem) || nome.includes('refeicao extra') || nome.includes('investigador') || nome.includes('guarda');
  };

  SP.ensureExtraAutomaticoSemana = async function(semanaId){
    const dias = ['segunda','terca','quarta','quinta','sexta'];
    const pedidos = await this.getPedidos(semanaId);
    for(const dia of dias){
      const jaTem = pedidos.some(p => norm(pick(p,'Dia'))===dia && norm(pick(p,'Colaborador_nome','Title','Nome'))==='refeicao extra' && norm(pick(p,'Origem','Tipo','tipo')).includes('extra'));
      if(!jaTem){
        await this.addExtraPedido(semanaId,dia,'Refeição extra','extra automatica','principal','Refeição extra automática');
      }
    }
    await this.setConfig('sync_timestamp', new Date().toISOString());
    return true;
  };

  SP.addExtraPedido = async function(semanaId, dia, nome, tipo, opcao='principal', observacao='', adicionadoPor=null){
    const normalDia=norm(dia), normalNome=norm(nome), normalTipo=norm(tipo);
    const pedidos = await this.getPedidos(semanaId);
    const existente = pedidos.find(p => norm(pick(p,'Dia'))===normalDia && norm(pick(p,'Colaborador_nome','Title','Nome'))===normalNome && norm(pick(p,'Origem','Tipo','tipo'))===normalTipo);
    if(existente) return existente;
    const cid = 'EXTRA-' + Date.now() + '-' + Math.random().toString(16).slice(2,6);
    if(this.addExtra){
      try{ await this.addExtra(semanaId,dia,nome,tipo,opcao,observacao,adicionadoPor || this.getUserName()); }catch(e){ console.warn('Extras list:',e); }
    }
    return await this.savePedido(semanaId,cid,nome,dia,opcao,nome,{confirmado:true,status:'Confirmado',origem:tipo,observacao:observacao,dataRef:this.getDataRefBySemanaDia?this.getDataRefBySemanaDia(semanaId,dia):undefined});
  };
})();


// ============================================================
// HOTFIX DEFINITIVO — campos SharePoint dinâmicos + extras limpos
// ============================================================
(function(){
  if(!window.SP) return;
  const SP = window.SP;
  const norm = v => String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/gi,'').toLowerCase();
  const money = v => { const n=Number(String(v ?? 0).replace(/[R$\s.]/g,'').replace(',','.')); return Number.isFinite(n)?n:0; };
  const pick = (obj,...keys)=>{ for(const k of keys){ if(obj && obj[k]!==undefined && obj[k]!==null) return obj[k]; } return ''; };

  SP._columnsCache = SP._columnsCache || {};

  SP.getListColumns = async function(listName){
    if(this._columnsCache[listName]) return this._columnsCache[listName];
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const data = await this.graph('GET', `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName,hidden,readOnly`);
    const cols = data.value || [];
    this._columnsCache[listName] = cols;
    return cols;
  };

  SP.findColumnName = async function(listName, candidates){
    const cols = await this.getListColumns(listName);
    const wanted = candidates.map(norm);
    let col = cols.find(c => wanted.includes(norm(c.name)) || wanted.includes(norm(c.displayName)));
    if(col) return col.name;
    col = cols.find(c => wanted.some(w => norm(c.name).includes(w) || norm(c.displayName).includes(w)));
    return col ? col.name : null;
  };

  SP.createItemMapped = async function(listName, mapping){
    const fields = {};
    for(const entry of mapping){
      const value = entry.value;
      if(value === undefined || value === null || value === '') continue;
      const internal = await this.findColumnName(listName, entry.names || [entry.name]);
      if(internal) fields[internal] = value;
      else if(entry.required) throw new Error(`Coluna não encontrada em ${listName}: ${(entry.names||[entry.name]).join(' / ')}`);
    }
    return this.createItem(listName, fields);
  };

  SP.updateItemMapped = async function(listName, itemId, mapping){
    const fields = {};
    for(const entry of mapping){
      const value = entry.value;
      if(value === undefined || value === null) continue;
      const internal = await this.findColumnName(listName, entry.names || [entry.name]);
      if(internal) fields[internal] = value;
    }
    return this.updateItem(listName, itemId, fields);
  };

  SP.createValorRefeicao = async function(dados){
    return this.createItemMapped('Valores de Refeição', [
      {names:['Title','Título','Titulo'], value:dados.title || dados.Title || 'Valor refeição', required:true},
      {names:['Data_Inicio','Data Início','Data Inicio','DataInicio','Inicio'], value:dados.dataInicio || dados.Data_Inicio, required:true},
      {names:['Data_Fim','Data Fim','DataFim','Fim'], value:dados.dataFim || dados.Data_Fim, required:true},
      {names:['Valor_Vascon','Valor Vascon','ValorVascon','Vascon'], value:money(dados.valorVascon ?? dados.Valor_Vascon), required:true},
      {names:['Valor_Desconto','Valor Desconto','ValorDesconto','Valor descontado funcionário','Valor descontado funcionario','Desconto funcionário','Desconto funcionario','Valor_Desconto_Funcionario','Valor Desconto Funcionario','Valor Desconto Funcionário'], value:money(dados.valorDesconto ?? dados.valorDescontoFuncionario ?? dados.Valor_Desconto ?? dados.Valor_Desconto_Funcionario)},
      {names:['Observacao','Observação','Obs'], value:dados.observacao || dados.Observacao || ''},
      {names:['Ativo','Status'], value:dados.ativo ?? dados.Ativo ?? true}
    ]);
  };

  SP.updateValorRefeicao = async function(id,dados){
    return this.updateItemMapped('Valores de Refeição', id, [
      {names:['Title','Título','Titulo'], value:dados.title ?? dados.Title},
      {names:['Data_Inicio','Data Início','Data Inicio','DataInicio','Inicio'], value:dados.dataInicio ?? dados.Data_Inicio},
      {names:['Data_Fim','Data Fim','DataFim','Fim'], value:dados.dataFim ?? dados.Data_Fim},
      {names:['Valor_Vascon','Valor Vascon','ValorVascon','Vascon'], value:dados.valorVascon!==undefined?money(dados.valorVascon):(dados.Valor_Vascon!==undefined?money(dados.Valor_Vascon):undefined)},
      {names:['Valor_Desconto','Valor Desconto','ValorDesconto','Valor descontado funcionário','Valor descontado funcionario','Desconto funcionário','Desconto funcionario','Valor_Desconto_Funcionario','Valor Desconto Funcionario','Valor Desconto Funcionário'], value:dados.valorDesconto!==undefined?money(dados.valorDesconto):(dados.Valor_Desconto!==undefined?money(dados.Valor_Desconto):undefined)},
      {names:['Observacao','Observação','Obs'], value:dados.observacao ?? dados.Observacao},
      {names:['Ativo','Status'], value:dados.ativo ?? dados.Ativo}
    ]);
  };

  SP.isExtraPedido = function(p){
    const origem = norm(pick(p,'Origem','origem','Tipo','tipo'));
    const nome = norm(pick(p,'Colaborador_nome','Nome','Title'));
    return origem.includes('extra') || ['investigador','guarda','prestador','visitante','motorista','marmita','outro'].includes(origem) || nome.includes('refeicaoextra') || nome.includes('investigador') || nome.includes('guarda');
  };

  SP.addExtraPedido = async function(semanaId, dia, nome, tipo, opcao='principal', observacao='', adicionadoPor=null){
    const nDia=norm(dia), nNome=norm(nome), nTipo=norm(tipo);
    const pedidos = await this.getPedidos(semanaId);
    const permitirDuplicado = nTipo==='investigador';
    if(!permitirDuplicado){
      const existente = pedidos.find(p => norm(pick(p,'Dia'))===nDia && norm(pick(p,'Colaborador_nome','Title','Nome'))===nNome && norm(pick(p,'Origem','Tipo','tipo'))===nTipo);
      if(existente) return existente;
    }
    let extraItem=null;
    try{
      extraItem = await this.createItemMapped('Extras', [
        {names:['Title','Título','Titulo'], value:`${semanaId}-${dia}-${nome}-${Date.now()}`},
        {names:['Semana_id','SemanaId','Semana','Semana ID'], value:semanaId},
        {names:['Dia'], value:dia},
        {names:['Nome','Title'], value:nome},
        {names:['tipo','Tipo'], value:tipo},
        {names:['Opcao','Opção','Opcao'], value:opcao},
        {names:['Observacao','Observação','Obs'], value:observacao||''},
        {names:['Adicionado_Por','Adicionado Por','Criado Por'], value:adicionadoPor||this.getUserName()}
      ]);
    }catch(e){ console.warn('[SP] Não salvou na lista Extras:',e); }

    return this.savePedido(semanaId, `EXTRA-${extraItem?.id || Date.now()}-${Math.random().toString(16).slice(2,5)}`, nome, dia, opcao, observacao || nome, {
      confirmado:true,
      status:'Confirmado',
      origem:tipo || 'extra',
      observacao:observacao||'',
      dataRef:this.getDataRefBySemanaDia ? this.getDataRefBySemanaDia(semanaId,dia) : undefined,
      alteradoPor:adicionadoPor||this.getUserName()
    });
  };

  SP.ensureExtraAutomaticoSemana = async function(semanaId){
    const dias=['segunda','terca','quarta','quinta','sexta'];
    const pedidos=await this.getPedidos(semanaId);
    for(const dia of dias){
      const existe = pedidos.some(p => norm(pick(p,'Dia'))===dia && norm(pick(p,'Colaborador_nome','Title','Nome'))==='refeicaoextra' && norm(pick(p,'Origem','Tipo','tipo')).includes('extra'));
      if(!existe) await this.addExtraPedido(semanaId,dia,'Refeição extra','extra automatica','principal','Refeição extra automática');
    }
    await this.setConfig('sync_timestamp', new Date().toISOString());
    return true;
  };

  SP.cleanupExtraAutomaticoSemana = async function(semanaId){
    const pedidos=await this.getPedidos(semanaId);
    const seen={};
    for(const p of pedidos){
      const dia=norm(pick(p,'Dia'));
      const isAuto = norm(pick(p,'Colaborador_nome','Title','Nome'))==='refeicaoextra' && norm(pick(p,'Origem','Tipo','tipo')).includes('extra');
      if(!isAuto) continue;
      if(seen[dia]){ try{ await this.deletePedido(p.id); }catch(e){ console.warn('Não removeu extra automático duplicado:',e); } }
      else seen[dia]=true;
    }
  };

  SP.deletePedido = async function(id){
    return this.deleteItem('Pedidos', id);
  };

  SP.deleteExtra = async function(id){
    return this.deleteItem('Extras', id);
  };

  SP.deleteExtraComPedido = async function(extra){
    if(extra && extra.id) { try{ await this.deleteExtra(extra.id); }catch(e){ console.warn(e); } }
    const semana = pick(extra,'Semana_id','Semana','semana');
    const dia = pick(extra,'Dia','dia');
    const nome = pick(extra,'Nome','Title','Colaborador_nome');
    const tipo = pick(extra,'tipo','Tipo','Origem');
    if(semana && dia && nome){
      const pedidos = await this.getPedidos(semana);
      for(const p of pedidos){
        if(norm(pick(p,'Dia'))===norm(dia) && norm(pick(p,'Colaborador_nome','Title','Nome'))===norm(nome) && (!tipo || norm(pick(p,'Origem','Tipo','tipo')).includes(norm(tipo)))){
          try{ await this.deletePedido(p.id); }catch(e){ console.warn(e); }
        }
      }
    }
  };
})();
