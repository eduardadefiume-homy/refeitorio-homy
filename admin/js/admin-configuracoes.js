// admin-configuracoes.js — Configurações do Admin Homy

const AdminConfiguracoes = window.AdminConfiguracoes = {

  async load() {
    // Remove bindings anteriores para re-vincular corretamente
    ["toggleEmail", "toggleExtra"].forEach(id => {
      const el = document.getElementById(id);
      if (el) delete el.dataset.boundCfg;
    });
    await this._carregarToggles();
    this._bindBotoes();
  },

  async _carregarToggles() {
    try {
      await SP.init();
      const [emailNotif, extraAuto] = await Promise.all([
        SP.getConfig("notificar_email"),
        SP.getConfig("refeicao_extra_automatica")
      ]);
      this._setToggle("toggleEmail", SP.isTrue(emailNotif));
      this._setToggle("toggleExtra", SP.isTrue(extraAuto));
    } catch (e) {
      console.warn("[Configurações] carregarToggles:", e);
    }
  },

  _setToggle(id, value) {
    const el = document.getElementById(id);
    if (el) el.checked = !!value;
  },

  // Cria 1 extra automático por dia útil da semana (CC = ADM GERAL)
  async _criarExtrasAutomaticos(semanaId) {
    const diasUteis = ["segunda", "terca", "quarta", "quinta", "sexta"];
    const user = SP.getUserName ? SP.getUserName() : "";
    let criados = 0;

    // Carrega todos os extras da semana de uma vez (1 chamada só)
    let extrasExistentes = [];
    try {
      extrasExistentes = await SP.getExtras(semanaId);
    } catch (e) {
      extrasExistentes = [];
    }

    for (const dia of diasUteis) {
      // Verifica se já existe extra automático nesse dia
      const normDia = dia.toLowerCase().trim();
      const jaExiste = extrasExistentes.some(e => {
        const eDia  = AdminUtils.norm(SP.pick(e, "Dia")           || "");
        const eTipo = AdminUtils.norm(SP.pick(e, "tipo", "Tipo")  || "");
        const eNome = AdminUtils.norm(SP.pick(e, "Nome", "Title") || "");
        return eDia === normDia &&
               (eTipo.includes("extra") || eNome.includes("refeicao extra") || eNome.includes("refeição extra"));
      });

      if (jaExiste) continue;

      try {
        await SP._addExtraPedidoCC(
          semanaId, dia,
          "Refeição Extra",
          "extra automatica",
          "principal",
          "Refeição extra automática do dia",
          "ADM GERAL - 120101",
          user
        );
        criados++;
      } catch (e) {
        console.warn("[extra auto] falha em " + dia + ":", e);
      }
    }

    return criados;
  },

  // Remove extras automáticos da semana
  async _removerExtrasAutomaticos(semanaId) {
    let extrasExistentes = [];
    try {
      extrasExistentes = await SP.getExtras(semanaId);
    } catch (e) {
      return 0;
    }

    let removidos = 0;
    for (const e of extrasExistentes) {
      const eTipo = AdminUtils.norm(SP.pick(e, "tipo", "Tipo")  || "");
      const eNome = AdminUtils.norm(SP.pick(e, "Nome", "Title") || "");
      if (eTipo.includes("extra automatica") || eTipo.includes("extra auto") ||
          eNome.includes("refeicao extra") || eNome.includes("refeição extra")) {
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

  _bindBotoes() {
    // Toggle notificação email
    const tEmail = document.getElementById("toggleEmail");
    if (tEmail && !tEmail.dataset.boundCfg) {
      tEmail.dataset.boundCfg = "1";
      tEmail.addEventListener("change", async function () {
        try {
          await SP.setConfig("notificar_email", this.checked ? "sim" : "nao");
          AdminUtils.toast(
            this.checked
              ? "✅ Notificação por email ativada."
              : "🔕 Notificação por email desativada.",
            "success"
          );
        } catch (e) {
          AdminUtils.toast("Erro ao salvar: " + e.message, "error");
        }
      });
    }

    // Toggle extra automático
    const tExtra = document.getElementById("toggleExtra");
    if (tExtra && !tExtra.dataset.boundCfg) {
      tExtra.dataset.boundCfg = "1";
      tExtra.addEventListener("change", async function () {
        const ativo = this.checked;
        tExtra.disabled = true;

        try {
          await SP.setConfig("refeicao_extra_automatica", ativo ? "sim" : "nao");

          const semanaId = AdminState.getSemanaId();

          if (ativo) {
            AdminUtils.toast("⏳ Criando extras automáticos...", "info");
            const criados = await AdminConfiguracoes._criarExtrasAutomaticos(semanaId);
            AdminUtils.toast(
              criados > 0
                ? "✅ Ativado. " + criados + " extras criados para esta semana."
                : "✅ Ativado. Extras já existiam para esta semana.",
              "success"
            );
          } else {
            AdminUtils.toast("⏳ Removendo extras automáticos...", "info");
            const removidos = await AdminConfiguracoes._removerExtrasAutomaticos(semanaId);
            AdminUtils.toast(
              "🔕 Desativado. " + removidos + " extras removidos desta semana.",
              "success"
            );
          }

          // Recarrega módulo de extras se estiver visível
          if (typeof AdminExtras !== "undefined" && AdminState.moduloAtivo === "extras") {
            AdminExtras.load(semanaId);
          }

        } catch (e) {
          console.error("[toggleExtra]", e);
          AdminUtils.toast("Erro: " + (e.message || e), "error");
          // Reverte o toggle visualmente em caso de erro
          tExtra.checked = !ativo;
        } finally {
          tExtra.disabled = false;
        }
      });
    }
  }
};
