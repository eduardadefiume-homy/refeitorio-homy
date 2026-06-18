# Segurança e Permissões

Sistema: Refeitório Homy  
Tipo: Segurança  
Responsável: TI Homy  
Última revisão: 18/06/2026  
Versão: 1.1  
Status: Pronto para revisão interna

## 1. Autenticação

O sistema usa Microsoft MSAL.js para autenticação com contas Microsoft da Homy.

Configuração atual conhecida:

| Campo | Valor |
|---|---|
| Client ID | `aa37acf9-f3bd-4d1e-968a-fde57f79094c` |
| Tenant ID | `a2850abc-334a-4805-b6b2-420b4aef68a9` |
| Redirect URI | `https://eduardadefiume-homy.github.io/refeitorio-homy/index.html` |
| Scopes | `Sites.ReadWrite.All`, `User.Read` |
| Fluxo | `loginPopup` |

## 2. Regras de autenticação

- Usar `loginPopup`, não `loginRedirect`, para evitar problemas no GitHub Pages.
- `redirectUri` deve ser fixo na raiz do app.
- Não usar `window.location.pathname` como redirect dinâmico.
- O cache técnico do MSAL pode usar `localStorage` para manter sessão entre páginas.

## 3. Permissões

Permissões Graph conhecidas:

- `User.Read`
- `Sites.ReadWrite.All`

Ponto de atenção: `Sites.ReadWrite.All` é uma permissão ampla. A TI deve validar se é possível reduzir o escopo em evolução futura sem quebrar o app.

## 4. Dados sensíveis

O sistema pode manipular:

- Nome de colaborador.
- Centro de custo.
- Email.
- Status de refeição.
- Ausências/férias/afastamentos.
- Informações de valores e relatórios.

## 5. Cuidados

- Não expor tokens em console.
- Não salvar dados de negócio em armazenamento local.
- Não publicar segredos no Git.
- Não dar permissão SharePoint desnecessária.
- Não abrir tela administrativa para usuários sem necessidade.

## 6. Alterações que exigem ADR

- Troca de fluxo MSAL.
- Mudança de escopos Graph.
- Mudança de tenant/client ID.
- Mudança de backend/dados.
- Migração de SharePoint para outra base.
- Criação de autenticação alternativa.
