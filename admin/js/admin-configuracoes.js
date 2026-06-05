// admin-configuracoes.js — Configurações do Admin Homy

const AdminConfiguracoes = window.AdminConfiguracoes = {

  async load() {
    await this._carregarToggles();
    this._bindBotoes();
  },

  async _carregarToggles() {
    try {
      await SP.init();
      const [marcacao, cardapio, emailNotif, extraAuto] = await Promise.all([
        SP.isCardapioLiberado(),
        SP.getConfig("cardapio_visivel"),
        SP.getConfig("notificar_email"),
        SP.getConfig("refeicao_extra_automatica")
      ]);

      this._setToggle("toggleEmail",  SP.isTrue(emailNotif));
      this._setToggle("toggleExtra",  SP.isTrue(extraAuto));
    } catch (e) {
      console.warn("[Configurações]", e);
    }
  },

  _setToggle(id, value) {
    const el = document.getElementById(id);
    if (el) el.checked = !!value;
  },

  _bindBotoes() {
    // Toggle notificação email
    const tEmail = document.getElementById("toggleEmail");
    if (tEmail && !tEmail.dataset.boundCfg) {
      tEmail.dataset.boundCfg = "1";
      tEmail.addEventListener("change", async function () {
        await SP.setConfig("notificar_email", this.checked ? "sim" : "nao");
        AdminUtils.toast(this.checked ? "Notificação por email ativada." : "Notificação desativada.", "success");
      });
    }

    // Toggle extra automático
    const tExtra = document.getElementById("toggleExtra");
    if (tExtra && !tExtra.dataset.boundCfg) {
      tExtra.dataset.boundCfg = "1";
      tExtra.addEventListener("change", async function () {
        await SP.setConfig("refeicao_extra_automatica", this.checked ? "sim" : "nao");
        AdminUtils.toast(this.checked ? "Refeição extra automática ativada." : "Desativada.", "success");
      });
    }
  }
};
