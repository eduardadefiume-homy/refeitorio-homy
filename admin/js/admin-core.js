// admin-core.js — Login MSAL e navegação entre módulos

const AdminCore = window.AdminCore = {

  MODULOS: {
    dashboard:     { title: "Dashboard",        sub: "Visão geral do refeitório" },
    cardapio:      { title: "Cardápio",          sub: "Gerenciar cardápio da semana" },
    pedidos:       { title: "Pedidos",           sub: "Ver e editar pedidos" },
    operacao:      { title: "Operação do Dia",   sub: "Acompanhamento e ajustes da produção" },
    colaboradores: { title: "Colaboradores",     sub: "Cadastro e gestão de colaboradores" },
    extras:        { title: "Extras / Visitantes", sub: "Refeições fora do quadro" },
    valores:       { title: "Valores",           sub: "Tabela de valores de refeição" },
    relatorios:    { title: "Relatórios",        sub: "Totais e estatísticas por semana" },
    configuracoes: { title: "Configurações",     sub: "Parâmetros gerais do sistema" }
  },

  _navBound: false,
  _semanaBound: false,
  _appCarregado: false,

  async init() {
    AdminUtils.bindModalClose();
    this._bindNav();
    this._bindSemana();

    try {
      const logado = await SP.init();

      if (logado && SP._account) {
        this._mostrarApp();
        return;
      }

    } catch (e) {
      console.warn("[AdminCore] init SP:", e);
    }

    this._mostrarLogin();
  },

  async login() {
    const btn = document.getElementById("btnLogin");
    const status = document.getElementById("loginStatus");

    try {
      if (btn) {
        btn.disabled = true;
        btn.textContent = "Redirecionando Microsoft...";
      }

      if (status) {
        status.textContent = "Abrindo login Microsoft...";
      }

      await SP.login();

    } catch (e) {
      console.error("[AdminCore] login:", e);

      const msg = e.message || e.errorCode || String(e);

      if (status) {
        status.textContent = "Erro: " + msg;
      }

      if (btn) {
        btn.disabled = false;
        btn.textContent = "Entrar com conta Homy";
      }

      AdminUtils.toast("Erro no login: " + msg, "error");
    }
  },

  async logout() {
    try {
      await SP.logout();
    } catch (e) {
      console.warn("[AdminCore] logout:", e);
    }

    this._appCarregado = false;
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

    if (!this._appCarregado) {
      this._appCarregado = true;
      this.loadModule("dashboard");
    }
  },

  _mostrarLogin() {
    document.getElementById("loginScreen")?.classList.remove("hide");
    document.getElementById("app")?.classList.remove("show");

    const btn = document.getElementById("btnLogin");
    if (btn) {
      btn.disabled = false;
      btn.textContent = "Entrar com conta Homy";
    }
  },

  _bindNav() {
    if (this._navBound) return;
    this._navBound = true;

    document.querySelectorAll(".nav-item[data-module]").forEach(item => {
      item.addEventListener("click", () => {
        this.loadModule(item.dataset.module);
      });
    });

    document.getElementById("btnLogin")?.addEventListener("click", () => this.login());
    document.getElementById("btnLogout")?.addEventListener("click", () => this.logout());
  },

  _bindSemana() {
    if (this._semanaBound) return;
    this._semanaBound = true;

    document.getElementById("btnSemanaAnterior")?.addEventListener("click", () => AdminState.semanaAnterior());
    document.getElementById("btnSemanaProxima")?.addEventListener("click", () => AdminState.semanaProxima());
  },

  loadModule(mod) {
    if (!this.MODULOS[mod]) return;

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
      valores:       () => AdminValores?.load(),
      relatorios:    () => AdminRelatorios?.load(semanaId),
      configuracoes: () => AdminConfiguracoes?.load()
    };

    try {
      loaders[mod]?.();
    } catch (e) {
      console.error(`[AdminCore] erro ao carregar módulo ${mod}:`, e);
      AdminUtils.toast("Erro ao carregar módulo: " + (e.message || e), "error");
    }
  }
};

document.addEventListener("DOMContentLoaded", () => AdminCore.init());
