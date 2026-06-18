# Visão Geral do Refeitório Homy

Sistema: Refeitório Homy  
Tipo: Visão de produto  
Responsável: TI Homy  
Última revisão: 18/06/2026  
Versão: 1.1  
Status: Pronto para revisão interna

## 1. Objetivo

O Refeitório Homy é um PWA interno criado para digitalizar e controlar o processo de refeição da Homy Química.

O sistema atende:

- Marcação semanal de refeições pelos colaboradores.
- Exibição de cardápio da semana e do dia.
- Operação diária da cozinha.
- Gestão administrativa de cardápio, pedidos, extras, ausências, valores e relatórios.

## 2. Escopo

Está dentro do escopo:

- Frontend estático em GitHub Pages.
- Autenticação Microsoft via MSAL.
- Armazenamento em SharePoint Lists.
- Integração com Microsoft Graph API.
- Relatórios administrativos.
- Operação diária da cozinha.

Fora do escopo atual:

- Backend próprio.
- Banco SQL, Firebase ou Dataverse.
- Multiempresa/multitenancy.
- Aplicativo nativo Android/iOS.
- Integração automática com folha de pagamento.

## 3. Perfis de usuário

| Perfil | Quem usa | Principais ações |
|---|---|---|
| Colaborador | Funcionários Homy | consultar cardápio, marcar refeição e consultar pedido |
| Cozinha | Equipe da cozinha | ver lista do dia e confirmar retirada |
| Admin | Luana/TI | gerenciar cardápio, pedidos, operação, colaboradores, extras, ausências, valores, relatórios e configurações |
| TI | Suporte/desenvolvimento | manter código, SharePoint, autenticação e documentação |

## 4. Módulos

| Módulo | Arquivo | Finalidade |
|---|---|---|
| Hub | `index.html` | página inicial com navegação |
| Cardápio Semana | `cardapio-semana.html` | consulta do cardápio de segunda a sexta |
| Cardápio Dia | `cardapio-dia.html` | consulta do cardápio do dia e pedido do colaborador |
| Marcar Refeição | `marcar-refeicao.html` | seleção de colaborador e refeições da semana |
| Cozinha | `cozinha.html` | operação de retirada diária |
| Admin | `admin/index.html` | painel administrativo completo |

## 5. Resultado esperado

Ao final de cada ciclo semanal:

- Cardápio cadastrado.
- Colaboradores com pedidos confirmados ou status operacional definido.
- Pendentes tratados por travamento ou ação manual.
- Cozinha com lista confiável do dia.
- Extras, ausências e relatórios refletidos no SharePoint.
- Gestão com dados rastreáveis.
