// admin-state.js — Estado global do Admin Homy
// Fonte única de verdade para offset de semana e módulo ativo.

const AdminState = window.AdminState = {
  semanaOffset: 0,
  moduloAtivo:  "dashboard",

  getSemanaId() {
    return AdminUtils.getSemanaIdByOffset(this.semanaOffset);
  },

  getSemanaLabel() {
    const id  = this.getSemanaId();
    const mon = AdminUtils.getMondayByOffset(this.semanaOffset);
    const fri = new Date(mon); fri.setDate(mon.getDate() + 4);
    return `${id} · ${AdminUtils.formatDateBR(mon)} a ${AdminUtils.formatDateBR(fri)}`;
  },

  semanaAnterior() { this.semanaOffset--; this._onSemanaChange(); },
  semanaProxima()  { this.semanaOffset++; this._onSemanaChange(); },

  _onSemanaChange() {
    // Atualiza badge de semana na topbar
    const badge = document.getElementById("semanaBadge");
    if (badge) badge.textContent = AdminUtils.formatSemana(this.getSemanaId());
    const label = document.getElementById("semanaLabel");
    if (label) label.textContent = this.getSemanaLabel();
    // Recarrega o módulo atual com a nova semana
    AdminCore.loadModule(this.moduloAtivo);
  }
};
