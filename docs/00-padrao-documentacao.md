# HDS - Homy Documentation Standard

Sistema: Refeitório Homy  
Tipo: Governança documental  
Responsável: TI Homy  
Última revisão: 18/06/2026  
Versão: 1.0  
Status: Pronto para revisão interna

## 1. Objetivo

Definir o padrão oficial de documentação do Refeitório Homy para que o sistema seja mantido com rastreabilidade, segurança operacional e facilidade de evolução.

## 2. Princípios

1. Código completo fica no Git.
2. Documentação técnica viva fica no Git, em Markdown.
3. Documentação formal interna fica em Word/PDF, por versão.
4. Toda alteração relevante precisa atualizar documentação, changelog e checklist de validação.
5. Nenhuma regra de negócio deve existir apenas em conversa, memória ou comentário solto no código.

## 3. Tipos de documentação

| Tipo | Onde fica | Finalidade |
|---|---|---|
| Técnica viva | Git `/docs` | manutenção diária, rastreabilidade e revisão por commit |
| Formal interna | Word/PDF Homy | referência institucional, auditoria e repasse interno |
| Operacional | Manual Admin, Cozinha, Colaborador | uso prático por perfil |
| Decisão arquitetural | Git `/docs/adr` | registrar por que uma decisão foi tomada |
| Histórico | `CHANGELOG.md` e Anexo B | saber o que mudou por versão |

## 4. Cabeçalho padrão

Todo documento deve iniciar com:

```md
# Nome do Documento

Sistema: Refeitório Homy  
Tipo: Técnico | Usuário | Operacional | Arquitetura | Requisito  
Responsável: TI Homy  
Última revisão: DD/MM/AAAA  
Versão: X.Y  
Status: Rascunho | Pronto para revisão | Validado | Obsoleto  
Público-alvo: TI | Admin | Cozinha | Colaborador | Gestão
```

## 5. Matriz de impacto documental

| Mudança | Atualizar no Git | Atualizar internamente |
|---|---|---|
| Regra de negócio | `04-regras-de-negocio.md`, `CHANGELOG.md` | Manual Técnico, Manual Operacional impactado |
| Fluxo Admin/Luana | `06-fluxos-operacionais.md`, módulo específico | Manual Operacional Admin |
| Tela Cozinha | `05-modulos.md`, `06-fluxos-operacionais.md` | Manual da Cozinha |
| Tela Colaborador | `05-modulos.md`, `06-fluxos-operacionais.md` | Manual do Colaborador |
| SharePoint | `03-sharepoint-listas-e-campos.md` | Anexo A, Manual Técnico |
| Autenticação/Segurança | `09-seguranca-e-permissoes.md`, ADR | Manual Técnico |
| Arquitetura | `01-arquitetura.md`, ADR | Manual Técnico |
| Bug recorrente | `07-troubleshooting.md`, `CHANGELOG.md` | Anexo C, erros conhecidos |
| Visual simples | `CHANGELOG.md` | Manual apenas se mudar o uso |

## 6. Código no manual

O manual interno não deve conter arquivos inteiros de código. Deve conter apenas trechos críticos e referências ao Git.

Permitido no manual:

- Snippets curtos de configuração.
- Nome de funções críticas.
- Nome de arquivos.
- Fluxos e regras.
- Prints e checklists.

Não recomendado no manual:

- Código completo de páginas HTML.
- Código completo de `sharepoint.js`.
- CSS completo.
- Cópias longas de módulos admin.

## 7. Definição de pronto

Uma alteração só está pronta quando:

- O código foi alterado no arquivo correto.
- A regra de negócio foi revisada.
- O SharePoint impactado foi identificado.
- O checklist de teste foi executado.
- O `CHANGELOG.md` foi atualizado.
- A documentação Git impactada foi atualizada.
- O manual interno impactado foi sinalizado para revisão.
- O commit segue Conventional Commits.
