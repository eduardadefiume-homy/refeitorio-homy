// PATCH — PRIORIDADE 2
// Adicionar no sharepoint.js existente

// ============================================================
// VALORES REFEIÇÃO
// ============================================================

async getValoresRefeicao() {
  const items = await this.getItems("Valores_Refeicao");
  return items;
},

async createValorRefeicao(dados) {
  return this.createItem("Valores_Refeicao", {
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

  return this.updateItem("Valores_Refeicao", id, fields);
}
