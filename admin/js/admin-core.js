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

  // ── Inicialização ───────────────────────────────────────────
  async init() {
    AdminUtils.bindModalClose();
    this._bindNav();
    this._bindSemana();

    // Tenta reaproveitar sessão existente antes de mostrar login
    try {
      const logado = await SP.init();
      if (logado) {
        this._mostrarApp();
        return;
      }
    } catch (e) {
      console.warn("[AdminCore] init SP:", e);
    }

    this._mostrarLogin();
  },

  // ── Login ───────────────────────────────────────────────────
  async login() {
    const btn = document.getElementById("btnLogin");
    const status = document.getElementById("loginStatus");

    try {
      if (btn) { btn.disabled = true; btn.textContent = "⏳ Abrindo Microsoft..."; }
      if (status) status.textContent = "Preparando autenticação...";

      await SP.init();

      const result = await SP._msalInstance.loginPopup({
        scopes: SP.scopes,
        prompt: "select_account"
      });

      if (!result?.account) throw new Error("Login sem conta autenticada.");

      SP._account = result.account;
      SP._msalInstance.setActiveAccount(result.account);

      this._mostrarApp();

    } catch (e) {
      console.error("[AdminCore] login:", e);
      if (status) status.textContent = "Erro: " + (e.message || e.errorCode || e);
      if (btn) { btn.disabled = false; btn.textContent = "🔐 Entrar com conta Homy"; }
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

  // ── Tela ────────────────────────────────────────────────────
  _mostrarApp() {
    document.getElementById("loginScreen")?.classList.add("hide");
    const app = document.getElementById("app");
    if (app) app.classList.add("show");

    const userInfo = document.getElementById("userInfo");
    if (userInfo) userInfo.textContent = SP.getUserName();

    // Atualiza badge de semana
    const badge = document.getElementById("semanaBadge");
    if (badge) badge.textContent = AdminUtils.formatSemana(AdminState.getSemanaId());

    this.loadModule("dashboard");
  },

  _mostrarLogin() {
    document.getElementById("loginScreen")?.classList.remove("hide");
    document.getElementById("app")?.classList.remove("show");
  },

  // ── Navegação ───────────────────────────────────────────────
  _bindNav() {
    document.querySelectorAll(".nav-item[data-module]").forEach(item => {
      item.addEventListener("click", () => {
        this.loadModule(item.dataset.module);
      });
    });

    document.getElementById("btnLogin")?.addEventListener("click", () => this.login());
    document.getElementById("btnLogout")?.addEventListener("click", () => this.logout());
  },

  _bindSemana() {
    document.getElementById("btnSemanaAnterior")?.addEventListener("click", () => AdminState.semanaAnterior());
    document.getElementById("btnSemanaProxima")?.addEventListener("click",  () => AdminState.semanaProxima());
  },

  loadModule(mod) {
    if (!this.MODULOS[mod]) return;

    AdminState.moduloAtivo = mod;

    // Atualiza nav
    document.querySelectorAll(".nav-item[data-module]").forEach(el => {
      el.classList.toggle("active", el.dataset.module === mod);
    });

    // Atualiza topbar
    const info = this.MODULOS[mod];
    AdminUtils.setTxt("topbarTitle", info.title);
    AdminUtils.setTxt("topbarSub",   info.sub);

    // Mostra módulo correto
    document.querySelectorAll(".module").forEach(el => el.classList.remove("active"));
    document.getElementById(`mod-${mod}`)?.classList.add("active");

    // Atualiza label de semana
    AdminUtils.setTxt("semanaLabel", AdminState.getSemanaLabel());

    // Carrega dados do módulo
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

    loaders[mod]?.();
  }
};

// Inicializa quando o DOM estiver pronto
document.addEventListener("DOMContentLoaded", () => AdminCore.init());
