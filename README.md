# Refeitório Homy

Sistema interno da Homy Química para marcação de refeições, gestão de cardápio, operação da cozinha, extras, ausências, valores e relatórios gerenciais.

## Produção

| Módulo | URL |
|---|---|
| Hub principal | `https://eduardadefiume-homy.github.io/refeitorio-homy/` |
| Admin | `https://eduardadefiume-homy.github.io/refeitorio-homy/admin/index.html` |
| Cardápio semana | `https://eduardadefiume-homy.github.io/refeitorio-homy/cardapio-semana.html` |
| Cardápio dia | `https://eduardadefiume-homy.github.io/refeitorio-homy/cardapio-dia.html` |
| Marcar refeição | `https://eduardadefiume-homy.github.io/refeitorio-homy/marcar-refeicao.html` |
| Cozinha | `https://eduardadefiume-homy.github.io/refeitorio-homy/cozinha.html` |

## Stack

- HTML5, CSS3 e JavaScript puro.
- GitHub Pages para hospedagem estática.
- Microsoft MSAL.js para autenticação com Microsoft Entra ID.
- Microsoft Graph API para leitura e gravação no SharePoint.
- SharePoint Lists como base de dados do sistema.
- SheetJS/ExcelJS para exportação de relatórios.

## Regra principal

SharePoint é a única fonte de verdade para dados de negócio. O fluxo correto é sempre:

```text
Interface -> sharepoint.js -> Microsoft Graph API -> SharePoint -> Atualizar tela com dados reais
```

É proibido usar `localStorage`, `sessionStorage` ou variáveis locais como banco de dados. Exceção: `localStorage` pode ser usado apenas como cache técnico do MSAL para autenticação entre páginas.

## Documentação

A documentação técnica viva está na pasta `/docs`.

| Documento | Finalidade |
|---|---|
| `docs/00-visao-geral.md` | visão do produto, perfis e escopo |
| `docs/00-padrao-documentacao.md` | padrão HDS de documentação Homy |
| `docs/01-arquitetura.md` | arquitetura técnica e fluxos de dados |
| `docs/02-instalacao-do-zero.md` | guia para replicar o sistema |
| `docs/03-sharepoint-listas-e-campos.md` | listas, campos e dicionário de dados |
| `docs/04-regras-de-negocio.md` | regras oficiais do sistema |
| `docs/05-modulos.md` | módulos e dependências |
| `docs/06-fluxos-operacionais.md` | rotinas da Luana, cozinha, colaboradores e TI |
| `docs/07-troubleshooting.md` | diagnóstico e erros conhecidos |
| `docs/08-checklist-testes.md` | checklists antes de publicar |
| `docs/09-seguranca-e-permissoes.md` | autenticação, permissões e riscos |
| `docs/10-governanca-git.md` | branches, commits, PRs e versionamento |
| `docs/adr/` | Architecture Decision Records |

## Publicação

1. Alterar os arquivos no Git.
2. Atualizar a documentação impactada.
3. Registrar no `CHANGELOG.md`.
4. Fazer commit seguindo Conventional Commits.
5. Aguardar publicação do GitHub Pages.
6. Validar em navegador com cache limpo.

## Status documental

Versão documental: `1.1.0`  
Data base: 18/06/2026  
Status: Pronto para revisão interna
