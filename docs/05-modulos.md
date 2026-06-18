# Módulos do Sistema

Sistema: Refeitório Homy  
Tipo: Referência técnica  
Responsável: TI Homy  
Última revisão: 18/06/2026  
Versão: 1.1  
Status: Pronto para revisão interna

## 1. Mapa geral de arquivos

| Arquivo | Tipo | Função |
|---|---|---|
| `sharepoint.js` | JS central | autenticação, Graph API e CRUD SharePoint |
| `index.html` | página | hub principal |
| `cardapio-semana.html` | página | cardápio semanal |
| `cardapio-dia.html` | página | cardápio diário e consulta de pedido |
| `marcar-refeicao.html` | página | marcação semanal por colaborador |
| `cozinha.html` | página | check-in da cozinha |
| `admin/index.html` | página | painel administrativo |

## 2. Módulos Admin

| Módulo | Arquivo | Lê | Grava |
|---|---|---|---|
| Core | `admin-core.js` | autenticação/sessão | logout/login via MSAL |
| State | `admin-state.js` | semana atual | estado local de navegação |
| Utils | `admin-utils.js` | DOM | DOM/toasts/modais |
| Dashboard | `admin-dashboard.js` | resumo, pedidos, extras, configs | configs/status/prazo |
| Cardápio | `admin-cardapio.js` | Cardapio | Cardapio |
| Pedidos | `admin-pedidos.js` | Pedidos | Pedidos |
| Operação do Dia | `admin-operacao-dia.js` | Pedidos | Pedidos/Status |
| Colaboradores | `admin-colaboradores.js` | Colaboradores | Colaboradores |
| Extras | `admin-extras.js` | Extras/Pedidos | Extras/Pedidos |
| Ausências | `admin-ausencias.js` | Ausencias do Refeitorio | Ausencias do Refeitorio |
| Valores | `admin-valores.js` | Valores de Refeição | Valores de Refeição |
| Relatórios | `admin-relatorios.js` | Pedidos | exportação local Excel |
| Configurações | `admin-configuracoes.js` | Configurações | Configurações |

## 3. Dependências críticas

- `sharepoint.js` deve carregar antes das páginas que usam `SP`.
- No Admin, a ordem dos scripts importa: utils, state, core e depois módulos.
- IDs HTML usados pelos módulos não devem ser removidos sem ajustar JS.
- Alterações em `AdminCore.MODULOS` precisam corresponder às seções do `admin/index.html`.

## 4. Pontos de atenção identificados

Os módulos admin atuais chamam algumas funções que devem ser confirmadas no `sharepoint.js` de produção:

- `SP.getDashboardResumo`
- `SP.setCardapioVisivel`
- `SP.addExtraPedido`
- `SP.deleteExtraComPedido`

Se essas funções não estiverem presentes, o módulo correspondente pode falhar em tempo de execução.
