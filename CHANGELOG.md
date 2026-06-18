# Changelog

Todas as alterações relevantes deste projeto devem ser registradas neste arquivo.

O formato segue a ideia de versionamento semântico:

- `PATCH`: correções pequenas sem mudança de regra.
- `MINOR`: novas funcionalidades ou alterações operacionais compatíveis.
- `MAJOR`: mudança estrutural, quebra de fluxo ou alteração crítica de regra.

## v1.1.0 - 18/06/2026

### Adicionado

- Criação do pacote de documentação técnica viva para Git.
- Criação do padrão HDS - Homy Documentation Standard.
- Criação dos ADRs iniciais do projeto.
- Criação dos manuais internos separados por público: Técnico, Admin, Cozinha e Colaborador.
- Criação dos anexos internos de SharePoint, histórico, suporte e matriz de regras.

### Alterado

- Separação entre documentação viva no Git e documentação formal interna da Homy.
- Reclassificação das regras de negócio críticas em documento próprio.
- Correção conceitual sobre `localStorage`: proibido como banco de dados, permitido apenas como cache técnico do MSAL quando previsto no projeto.

### Regra de negócio

- SharePoint permanece como fonte única de verdade.
- Modais não podem salvar ao fechar ou cancelar.
- Edição de cardápio deve atualizar registro existente e preservar ID.
- Operação do Dia passa a tratar explicitamente status Confirmado, Cancelado, Não vai almoçar, Férias, Afastado, Bloqueado e Travado.
- Extras devem estar refletidos operacionalmente em Pedidos quando a funcionalidade integrada estiver ativa.

### Documentação

- Criada documentação Git em `/docs`.
- Criada documentação interna Homy em arquivos Word separados.

### Pontos a validar

- Confirmar nomes internos reais de todos os campos no SharePoint.
- Validar funções esperadas em `sharepoint.js` que são chamadas por módulos admin: `getDashboardResumo`, `setCardapioVisivel`, `addExtraPedido` e `deleteExtraComPedido`.

## v1.0.0 - Junho/2026

### Adicionado

- Manual Técnico inicial do Sistema Refeitório Homy.
- Registro da arquitetura: GitHub Pages, MSAL, Graph API e SharePoint Lists.
- Registro dos módulos: Hub, Cardápio Semana, Cardápio Dia, Marcar Refeição, Cozinha e Admin.
- Registro inicial das listas SharePoint e regras críticas.
