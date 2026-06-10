// admin-configuracoes.js — Configurações do Admin Homy

const AdminConfiguracoes = window.AdminConfiguracoes = {

  async load() {
    await this._carregarToggles();
    this._bindBotoes();
    await this._garantirExtrasSeAtivo();
  },

  async _garantirExtrasSeAtivo() {
    try {
      await SP.init();

      const ativo = SP.isTrue(await SP.getConfig("refeicao_extra_automatica"));
      if (!ativo) return;

      const semanaId = AdminState.getSemanaId();
      const criados = await this._criarExtrasAutomaticos(semanaId);

      if (criados > 0) {
        AdminUtils.toast(`${criados} refeição(ões) extra automática(s) criada(s) para esta semana.`, "success");
      }

    } catch (e) {
      console.warn("[extras auto] _garantirExtrasSeAtivo:", e);
    }
  },

  async _carregarToggles() {
    try {
      await SP.init();

      const [emailNotif, extraAuto, marcacaoLiberada, cardapioVisivel] = await Promise.all([
        SP.getConfig("notificar_email").catch(() => "nao"),
        SP.getConfig("refeicao_extra_automatica").catch(() => "nao"),
        SP.getConfig("marcacao_liberada").catch(() => "sim"),
        SP.getConfig("cardapio_visivel").catch(() => "sim")
      ]);

      this._setToggle("toggleEmail", SP.isTrue(emailNotif));
      this._setToggle("toggleExtra", SP.isTrue(extraAuto));
      this._setToggle("toggleMarcacao", SP.isTrue(marcacaoLiberada));
      this._setToggle("toggleCardapio", SP.isTrue(cardapioVisivel));

      this._setToggle("dashToggleMarcacao", SP.isTrue(marcacaoLiberada));
      this._setToggle("dashToggleCardapio", SP.isTrue(cardapioVisivel));

    } catch (e) {
      console.warn("[Configurações] carregarToggles:", e);
    }
  },

  _setToggle(id, value) {
    const el = document.getElementById(id);
    if (el) el.checked = !!value;
  },

  async _criarExtrasAutomaticos(semanaId) {
    const diasUteis = ["segunda", "terca", "quarta", "quinta", "sexta"];
    const user = SP.getUserName ? SP.getUserName() : "";
    let criados = 0;

    let extrasExistentes = [];
    try {
      extrasExistentes = await SP.getExtras(semanaId);
    } catch (e) {
      extrasExistentes = [];
    }

    for (const dia of diasUteis) {
      const normDia = AdminUtils.norm(dia);

      const jaExiste = extrasExistentes.some(e => {
        const eDia = AdminUtils.norm(SP.pick(e, "Dia") || "");
        const eTipo = AdminUtils.norm(SP.pick(e, "tipo", "Tipo") || "");
        const eNome = AdminUtils.norm(SP.pick(e, "Nome", "Title") || "");

        return eDia === normDia &&
          (
            eTipo.includes("extra") ||
            eNome.includes("refeicao extra") ||
            eNome.includes("refeição extra")
          );
      });

      if (jaExiste) continue;

      await SP._addExtraPedidoCC(
        semanaId,
        dia,
        "Refeição Extra",
        "extra automatica",
        "principal",
        "Refeição extra automática do dia",
        "120101 - ADM GERAL",
        user
      );

      criados++;
    }

    return criados;
  },

  async _removerExtrasAutomaticos(semanaId) {
    let extrasExistentes = [];

    try {
      extrasExistentes = await SP.getExtras(semanaId);
    } catch (e) {
      return 0;
    }

    let removidos = 0;

    for (const e of extrasExistentes) {
      const eTipo = AdminUtils.norm(SP.pick(e, "tipo", "Tipo") || "");
      const eNome = AdminUtils.norm(SP.pick(e, "Nome", "Title") || "");

      if (
        eTipo.includes("extra automatica") ||
        eTipo.includes("extra auto") ||
        eNome.includes("refeicao extra") ||
        eNome.includes("refeição extra")
      ) {
        try {
          await SP.deleteExtraComPedido(e);
          removidos++;
        } catch (err) {
          console.warn("[extra auto] falha ao remover:", err);
        }
      }
    }

    return removidos;
  },

  async salvarToggleConfig(chave, ativo, mensagens = {}) {
    await SP.init();
    await SP.setConfig(chave, ativo ? "sim" : "nao");

    AdminUtils.toast(
      ativo ? (mensagens.on || "Ativado.") : (mensagens.off || "Desativado."),
      "success"
    );
  },

  _bindBotoes() {
    const bindToggle = (id, chave, mensagens, depois) => {
      const el = document.getElementById(id);
      if (!el || el.dataset.boundCfg) return;

      el.dataset.boundCfg = "1";

      el.addEventListener("change", async function () {
        const ativo = this.checked;
        this.disabled = true;

        try {
          await AdminConfiguracoes.salvarToggleConfig(chave, ativo, mensagens);

          if (typeof depois === "function") {
            await depois(ativo);
          }

          if (id === "toggleMarcacao") {
            AdminConfiguracoes._setToggle("dashToggleMarcacao", ativo);
          }

          if (id === "toggleCardapio") {
            AdminConfiguracoes._setToggle("dashToggleCardapio", ativo);
          }

          if (id === "dashToggleMarcacao") {
            AdminConfiguracoes._setToggle("toggleMarcacao", ativo);
          }

          if (id === "dashToggleCardapio") {
            AdminConfiguracoes._setToggle("toggleCardapio", ativo);
          }

        } catch (e) {
          console.error("[toggle config]", e);
          AdminUtils.toast("Erro: " + (e.message || e), "error");
          this.checked = !ativo;
        } finally {
          this.disabled = false;
        }
      });
    };

    bindToggle("toggleEmail", "notificar_email", {
      on: "Notificação por email ativada.",
      off: "Notificação por email desativada."
    });

    bindToggle("toggleMarcacao", "marcacao_liberada", {
      on: "Marcação de refeição liberada.",
      off: "Marcação de refeição bloqueada."
    });

    bindToggle("toggleCardapio", "cardapio_visivel", {
      on: "Cardápio da semana visível.",
      off: "Cardápio da semana oculto."
    });

    bindToggle("dashToggleMarcacao", "marcacao_liberada", {
      on: "Marcação de refeição liberada.",
      off: "Marcação de refeição bloqueada."
    });

    bindToggle("dashToggleCardapio", "cardapio_visivel", {
      on: "Cardápio da semana visível.",
      off: "Cardápio da semana oculto."
    });

    const tExtra = document.getElementById("toggleExtra");
    if (tExtra && !tExtra.dataset.boundCfg) {
      tExtra.dataset.boundCfg = "1";

      tExtra.addEventListener("change", async function () {
        const ativo = this.checked;
        this.disabled = true;

        try {
          await SP.setConfig("refeicao_extra_automatica", ativo ? "sim" : "nao");

          const semanaId = AdminState.getSemanaId();

          if (ativo) {
            AdminUtils.toast("Criando extras automáticos...", "info");
            const criados = await AdminConfiguracoes._criarExtrasAutomaticos(semanaId);

            AdminUtils.toast(
              criados > 0
                ? `Ativado. ${criados} extras criados para esta semana.`
                : "Ativado. Extras já existiam para esta semana.",
              "success"
            );
          } else {
            AdminUtils.toast("Removendo extras automáticos...", "info");
            const removidos = await AdminConfiguracoes._removerExtrasAutomaticos(semanaId);
            AdminUtils.toast(`Desativado. ${removidos} extras removidos desta semana.`, "success");
          }

          if (typeof AdminExtras !== "undefined" && AdminState.moduloAtivo === "extras") {
            AdminExtras.load(semanaId);
          }

        } catch (e) {
          console.error("[toggleExtra]", e);
          AdminUtils.toast("Erro: " + (e.message || e), "error");
          tExtra.checked = !ativo;
        } finally {
          tExtra.disabled = false;
        }
      });
    }

    const btnSalvarPrazo = document.getElementById("btnSalvarPrazo");
    if (btnSalvarPrazo && !btnSalvarPrazo.dataset.boundCfg) {
      btnSalvarPrazo.dataset.boundCfg = "1";

      btnSalvarPrazo.addEventListener("click", async () => {
        const data = AdminUtils.getVal("prazoData") || AdminUtils.getVal("dashPrazoData");
        const hora = AdminUtils.getVal("prazoHora") || AdminUtils.getVal("dashPrazoHora") || "10:00";

        if (!data) {
          AdminUtils.toast("Informe a data limite.", "error");
          return;
        }

        try {
          const iso = `${data}T${hora}:00`;
          await SP.setPrazoMarcacao(iso);
          AdminUtils.toast("Prazo salvo.", "success");
        } catch (e) {
          AdminUtils.toast("Erro ao salvar prazo: " + (e.message || e), "error");
        }
      });
    }
  }
};
