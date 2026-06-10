// admin-core.js — Login MSAL e navegação entre módulos

const AdminCore = window.AdminCore = {

  MODULOS: {
    dashboard:     { title: "Dashboard",           sub: "Visão geral do refeitório" },
    cardapio:      { title: "Cardápio",             sub: "Gerenciar cardápio da semana" },
    pedidos:       { title: "Pedidos",              sub: "Ver e editar pedidos" },
    operacao:      { title: "Operação do Dia",      sub: "Acompanhamento e ajustes da produção" },
    colaboradores: { title: "Colaboradores",        sub: "Cadastro e gestão de colaboradores" },
    extras:        { title: "Extras / Visitantes",  sub: "Refeições fora do quadro" },
    ausencias:     { title: "Ausências",            sub: "Férias, atestados e afastamentos" },
    valores:       { title: "Valores",              sub: "Tabela de valores de refeição" },
    relatorios:    { title: "Relatórios",           sub: "Totais e estatísticas por semana" },
    configuracoes: { title: "Configurações",        sub: "Parâmetros gerais do sistema" }
  },

  _iniciado: false,

  async init() {
    if (this._iniciado) return;
    this._iniciado = true;

    AdminUtils.bindModalClose();
    this._bindNav();
    this._bindSemana();

    this._mostrarLogin();

    try {
      if (!window.SP) {
        throw new Error("SP não carregou. Verifique o caminho ../sharepoint.js no admin/index.html.");
      }

      await SP.waitForMsal?.();

      const logado = await SP.init();

      if (logado && SP._account) {
        this._mostrarApp();
        return;
      }

      this._mostrarLogin();

    } catch (e) {
      console.warn("[AdminCore] init:", e);

      const status = document.getElementById("loginStatus");
      if (status) status.textContent = "Erro: " + (e.message || e);

      this._mostrarLogin();
    }
  },

  async login() {
    const btn = document.getElementById("btnLogin");
    const status = document.getElementById("loginStatus");

    try {
      if (btn) {
        btn.disabled = true;
        btn.textContent = "⏳ Abrindo Microsoft...";
      }

      if (status) status.textContent = "Preparando autenticação...";

      if (!window.SP) {
        throw new Error("SP não carregou. Verifique o sharepoint.js.");
      }

      await SP.waitForMsal?.();
      await SP.init();

      const ok = await SP.login();

      if (!ok || !SP._account) {
        throw new Error("Login Microsoft não foi concluído.");
      }

      this._mostrarApp();

    } catch (e) {
      console.error("[AdminCore] login:", e);

      if (status) status.textContent = "Erro: " + (e.message || e.errorCode || e);

      if (btn) {
        btn.disabled = false;
        btn.textContent = "🔐 Entrar com conta Homy";
      }

      AdminUtils.toast("Erro no login: " + (e.message || e), "error");
    }
  },

  async logout() {
    try {
      await SP.logout();
    } catch (e) {
      console.warn("[AdminCore] logout:", e);
    }

    document.getElementById("loginScreen")?.classList.remove("hide");
    document.getElementById("app")?.classList.remove("show");
  },

  _mostrarApp() {
    document.getElementById("loginScreen")?.classList.add("hide");

    const app = document.getElementById("app");
    if (app) app.classList.add("show");

    const userInfo = document.getElementById("userInfo");
    if (userInfo) userInfo.textContent = SP.getUserName();

    const badge = document.getElementById("semanaBadge");
    if (badge) badge.textContent = AdminUtils.formatSemana(AdminState.getSemanaId());

    const label = document.getElementById("semanaLabel");
    if (label) label.textContent = AdminState.getSemanaLabel();

    this.loadModule(AdminState.moduloAtivo || "dashboard");
  },

  _mostrarLogin() {
    document.getElementById("loginScreen")?.classList.remove("hide");
    document.getElementById("app")?.classList.remove("show");
  },

  _bindNav() {
    document.querySelectorAll(".nav-item[data-module]").forEach(item => {
      if (item.dataset.boundCore) return;
      item.dataset.boundCore = "1";
      item.addEventListener("click", () => this.loadModule(item.dataset.module));
    });

    const btnLogin = document.getElementById("btnLogin");
    if (btnLogin && !btnLogin.dataset.boundCore) {
      btnLogin.dataset.boundCore = "1";
      btnLogin.addEventListener("click", () => this.login());
    }

    const btnLogout = document.getElementById("btnLogout");
    if (btnLogout && !btnLogout.dataset.boundCore) {
      btnLogout.dataset.boundCore = "1";
      btnLogout.addEventListener("click", () => this.logout());
    }
  },

  _bindSemana() {
    const ant = document.getElementById("btnSemanaAnterior");
    if (ant && !ant.dataset.boundCore) {
      ant.dataset.boundCore = "1";
      ant.addEventListener("click", () => AdminState.semanaAnterior());
    }

    const prox = document.getElementById("btnSemanaProxima");
    if (prox && !prox.dataset.boundCore) {
      prox.dataset.boundCore = "1";
      prox.addEventListener("click", () => AdminState.semanaProxima());
    }
  },

  async loadModule(mod) {
    if (!this.MODULOS[mod]) return;

    if (!SP?._account) {
      this._mostrarLogin();
      return;
    }

    AdminState.moduloAtivo = mod;

    document.querySelectorAll(".nav-item[data-module]").forEach(el => {
      el.classList.toggle("active", el.dataset.module === mod);
    });

    const info = this.MODULOS[mod];
    AdminUtils.setTxt("topbarTitle", info.title);
    AdminUtils.setTxt("topbarSub", info.sub);

    document.querySelectorAll(".module").forEach(el => el.classList.remove("active"));
    document.getElementById(`mod-${mod}`)?.classList.add("active");

    AdminUtils.setTxt("semanaLabel", AdminState.getSemanaLabel());

    const semanaId = AdminState.getSemanaId();

    const loaders = {
      dashboard:     () => AdminDashboard?.load(semanaId),
      cardapio:      () => AdminCardapio?.load(semanaId),
      pedidos:       () => AdminPedidos?.load(semanaId),
      operacao:      () => AdminOperacao?.load(semanaId),
      colaboradores: () => AdminColaboradores?.load(),
      extras:        () => AdminExtras?.load(semanaId),
      ausencias:     () => AdminAusencias?.load(),
      valores:       () => AdminValores?.load(),
      relatorios:    () => AdminRelatorios?.load(semanaId),
      configuracoes: () => AdminConfiguracoes?.load()
    };

    try {
      await loaders[mod]?.();
    } catch (e) {
      console.error(`[AdminCore] loadModule ${mod}:`, e);
      AdminUtils.toast(`Erro ao carregar ${info.title}: ` + (e.message || e), "error");
    }
  }
};

document.addEventListener("DOMContentLoaded", () => {
  if (!window.__ADMIN_BOOT_EXTERNAL__) {
    AdminCore.init();
  }
});
