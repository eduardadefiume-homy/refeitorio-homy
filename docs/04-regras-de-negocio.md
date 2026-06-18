# Regras de Negócio

Sistema: Refeitório Homy  
Tipo: Regras de negócio  
Responsável: TI Homy  
Última revisão: 18/06/2026  
Versão: 1.1  
Status: Pronto para revisão interna

## RN-001 - SharePoint é fonte única de verdade

Todos os dados de negócio devem ser lidos e gravados no SharePoint. A interface só pode atualizar o estado visual após confirmação da gravação/leitura no SharePoint.

## RN-002 - Não usar armazenamento local como banco

É proibido usar `localStorage`, `sessionStorage` ou variáveis locais como base de dados de negócio.

Exceção: `localStorage` pode ser usado como cache técnico do MSAL para autenticação entre páginas.

## RN-003 - Semana ISO

As operações semanais usam `Semana_id` no formato `YYYY-Www`, calculado por semana ISO.

## RN-004 - Cardápio único por semana, dia e opção

A chave lógica do cardápio é:

```text
Semana_id + Dia + Opcao
```

Editar um prato deve atualizar o registro existente, preservar o ID e não criar duplicata.

## RN-005 - Modais não salvam ao cancelar ou fechar

Todo modal deve ter ações explícitas:

- Fechar: apenas fecha.
- Cancelar: apenas fecha.
- Salvar: grava no SharePoint.

É proibido auto-save, salvar ao fechar ou salvar ao cancelar.

## RN-006 - Marcação de refeição

A marcação depende do toggle administrativo. Quando liberada, colaboradores ativos podem escolher refeições da semana. Cada pedido deve ser gravado em `Pedidos` com colaborador, dia, opção, semana, status e origem.

## RN-007 - Prazo de marcação

O prazo é informativo. O bloqueio/travamento ocorre por ação administrativa explícita, salvo se funcionalidade futura automatizar isso e for documentada.

## RN-008 - Travamento de pendentes

Quando o admin trava pendentes, colaboradores ativos sem pedido para o dia selecionado recebem pedido automático como Principal, com status Confirmado, origem Admin/Automático e observação de auditoria.

## RN-009 - Operação do Dia

A Operação do Dia deve refletir os pedidos existentes no SharePoint para a semana e dia selecionados.

Status produtivos:

- Confirmado
- Extra
- Aprovado

Status cancelados/não produtivos:

- Cancelado
- Férias/Ferias
- Afastado
- Não vai almoçar
- Bloqueado
- Travado

## RN-010 - Ações da Operação

A Operação do Dia pode alterar status de pedido para Confirmado, Cancelado ou Não vai almoçar. Alterações devem gravar `Alterado_Por` e `Origem` quando aplicável.

## RN-011 - Extras

Extras representam visitantes, guardas, investigadores, prestadores ou refeições fora do quadro normal. Quando a integração operacional estiver ativa, todo extra precisa gerar ou estar refletido em um pedido para aparecer na Operação do Dia e Cozinha.

## RN-012 - Ausências

Ausências ativas no período impedem confirmação normal na Cozinha e devem aparecer com status visual de ausente/travado.

Motivos aceitos:

- ferias
- atestado
- falta
- licenca
- afastamento
- nao_vai_almocar
- homy_office
- banco_horas
- outro

## RN-013 - Valores de refeição

Apenas um registro de valor deve estar ativo por vez. Períodos de vigência não devem se sobrepor. Alterações em valores impactam relatórios e desconto de funcionário.

## RN-014 - Check-in da cozinha

Confirmar retirada grava registro em `CheckIn` com semana, dia, colaborador, horário, usuário confirmador e flag `Retirou`.

## RN-015 - Relatórios

Relatórios devem considerar pedidos confirmados/produtivos no período selecionado. Pedidos automáticos sem `Data_Hora` adequada devem ser tratados por `Semana_id` quando aplicável.

## RN-016 - Segurança

Acesso administrativo exige autenticação Microsoft Homy. Alterações em permissões, escopos, tenant, client ID ou redirect URI devem ser registradas em ADR.
