// ============================================================
// admin-state.js — Estado global controlado do Admin
// ============================================================

window.AdminState = {
  semanaAtual: null,
  usuarioAtual: null,
  colaboradores: [],
  cardapio: [],
  pedidos: [],
  extras: [],
  valoresRefeicao: [],
  ausencias: [],
  filtros: {
    dia: null,
    status: "Todos os status",
    busca: "",
    dataInicio: null,
    dataFim: null,
    centroCusto: null
  }
};
