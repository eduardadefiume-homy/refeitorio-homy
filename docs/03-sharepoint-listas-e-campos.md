# SharePoint - Listas e Campos

Sistema: Refeitório Homy  
Tipo: Dicionário de dados  
Responsável: TI Homy  
Última revisão: 18/06/2026  
Versão: 1.1  
Status: Pronto para revisão interna

## 1. Regra crítica

O sistema localiza listas usando `displayName`. Portanto, os nomes das listas devem ser exatamente iguais aos documentados.

## 2. Site

```text
homyquimica.sharepoint.com/sites/Refeitrio-Homy
```

## 3. Listas obrigatórias

| Lista | Finalidade | Observação |
|---|---|---|
| Cardapio | opções de refeição por semana/dia/opção | sem acento |
| Pedidos | pedidos dos colaboradores e extras | lista central da operação |
| Colaboradores | cadastro base de pessoas | filtra ativos |
| Extras | visitantes, guardas, investigadores e refeições extras | pode refletir em Pedidos |
| Configurações | toggles e parâmetros do sistema | chave/valor |
| Valores de Refeição | valores Vascon e desconto | nome com acento |
| Ausencias do Refeitorio | férias, afastamentos e não vai almoçar | sem acento em Ausencias |
| CheckIn | retirada de refeições pela cozinha | confirmação diária |

## 4. Cardapio

| Campo interno | Tipo esperado | Obrigatório | Uso |
|---|---|---|---|
| Title | Texto | Sim | chave composta `semana-dia-opcao` |
| Semana_id | Texto | Sim | identifica semana ISO |
| Dia | Texto | Sim | segunda, terca, quarta, quinta, sexta |
| Opcao | Texto | Sim | principal, light, carne, massa, lanche |
| Nome_Prato | Texto | Sim | nome do prato |
| Detalhes | Texto múltiplo | Não | acompanhamentos/detalhes |

Regra: chave única lógica = `Semana_id + Dia + Opcao`.

## 5. Pedidos

| Campo interno | Tipo esperado | Obrigatório | Uso |
|---|---|---|---|
| Title | Texto | Sim | chave composta do pedido |
| Semana_id | Texto | Sim | semana ISO |
| Colaborador_id | Texto | Sim | id do colaborador ou id sintético de extra |
| Colaborador_nome | Texto | Sim | nome exibido |
| Dia | Texto | Sim | dia do pedido |
| Opcao | Texto | Sim | opção escolhida |
| Nome_Prato | Texto | Não | prato vinculado |
| Confirmado | Sim/Não | Sim | indica se conta para produção |
| Data_Hora | Data/Hora | Sim | data de criação/alteração |
| Centro_Custo | Texto | Não | CC do colaborador/extra |
| Status | Texto | Sim | Confirmado, Cancelado, Não vai almoçar etc. |
| Observacao | Texto múltiplo | Não | auditoria/observações |
| Origem | Texto | Sim | Refeitório, Admin, Extra, Automático |
| Alterado_Por | Texto | Não | usuário que alterou |

## 6. Colaboradores

| Campo interno | Tipo esperado | Obrigatório | Uso |
|---|---|---|---|
| Title | Texto | Sim | título do item |
| Nome | Texto | Sim | nome do colaborador |
| Departamento | Texto | Não | departamento/setor |
| Centro_Custo | Texto | Não | centro de custo |
| Email | Texto | Não | email corporativo |
| Ativo | Sim/Não | Sim | define se aparece para marcação |
| tipo | Texto | Não | colaborador/admin/outro |

## 7. Extras

| Campo interno | Tipo esperado | Obrigatório | Uso |
|---|---|---|---|
| Title | Texto | Sim | chave do extra |
| Semana_id | Texto | Sim | semana ISO |
| Dia | Texto | Sim | dia do extra |
| Nome | Texto | Sim | visitante/guarda/investigador |
| tipo | Texto | Sim | visitante, guarda, investigador, prestador, extra |
| Opcao | Texto | Sim | opção de refeição |
| Observacao | Texto | Não | observação |
| Adicionado_Por | Texto | Não | usuário que adicionou |
| Status | Texto | Não | pendente/confirmado se aplicável |

## 8. Configurações

| Campo interno | Tipo esperado | Obrigatório | Uso |
|---|---|---|---|
| Title | Texto | Sim | chave |
| Chave | Texto | Sim | nome do parâmetro |
| Valor | Texto | Sim | valor gravado |

Chaves conhecidas:

- `cardapio_liberado`
- `marcacao_liberada`
- `pedidos_liberados`
- `cardapio_visivel`
- `prazo_limite`
- `notificar_email`
- `refeicao_extra_automatica`

## 9. Valores de Refeição

| Campo interno | Tipo esperado | Obrigatório | Uso |
|---|---|---|---|
| Title / Título | Texto | Sim | título do período |
| Data_Inicio | Data | Sim | início da vigência |
| Data_Fim | Data | Sim | fim da vigência |
| Valor_Vascon | Número/Moeda | Sim | valor cobrado pela Vascon |
| Valor_Desconto_Funcionário | Número/Moeda | Sim | desconto do funcionário |
| Observacao | Texto | Não | observação |
| Ativo | Sim/Não | Sim | registro vigente |

Observação: o código possui resolução dinâmica de colunas para essa lista. Mesmo assim, o nome da lista `Valores de Refeição` não pode ser alterado.

## 10. Ausencias do Refeitorio

| Campo interno | Tipo esperado | Obrigatório | Uso |
|---|---|---|---|
| Title | Texto | Sim | título da ausência |
| Colaborador_id | Texto | Sim | id do colaborador |
| Colaborador_nome | Texto | Sim | nome do colaborador |
| Centro_Custo | Texto | Não | centro de custo |
| Data_Inicio | Data | Sim | início |
| Data_Fim | Data | Sim | fim |
| Motivo | Texto | Sim | ferias, atestado, falta, licenca, afastamento, nao_vai_almocar, homy_office, banco_horas, outro |
| Observacao | Texto | Não | observação |
| Ativo | Sim/Não | Sim | se ainda vale |
| Criado_Por | Texto | Não | usuário que criou |

## 11. CheckIn

| Campo interno | Tipo esperado | Obrigatório | Uso |
|---|---|---|---|
| Title | Texto | Sim | chave do check-in |
| Semana_id | Texto | Sim | semana ISO |
| Colaborador_id | Texto | Sim | id do colaborador |
| Colaborador_nome | Texto | Sim | nome do colaborador |
| Dia | Texto | Sim | dia do check-in |
| Retirou | Sim/Não | Sim | retirada confirmada |
| Data_Hora_Retirada | Data/Hora | Sim | horário da retirada |
| Confirmado_Por | Texto | Sim | usuário que confirmou |

## 12. Pontos a validar

- Exportar do SharePoint os nomes internos reais das colunas.
- Confirmar se todas as colunas existem em produção.
- Confirmar se funções chamadas por módulos admin existem em `sharepoint.js` atual.
- Confirmar se `Extras` integrados criam registro correspondente em `Pedidos`.
