# Governança Git, Commits e PRs

Sistema: Refeitório Homy  
Tipo: Governança técnica  
Responsável: TI Homy  
Última revisão: 18/06/2026  
Versão: 1.0  
Status: Pronto para revisão interna

## 1. Branches

Usar nomes claros, minúsculos e com hífen:

```text
feat/nome-da-funcionalidade
fix/nome-do-problema
docs/nome-da-documentacao
refactor/nome-do-modulo
security/nome-do-ajuste
data/nome-lista-sharepoint
```

Exemplos:

- `fix/pedidos-sumindo-operacao-dia`
- `feat/extras-integrados-pedidos`
- `docs/manual-tecnico-v1-1`
- `data/campos-internos-valores-refeicao`

## 2. Commits

Usar Conventional Commits:

```text
tipo(escopo): descrição curta
```

Tipos principais:

- `feat`
- `fix`
- `docs`
- `refactor`
- `style`
- `perf`
- `test`
- `chore`
- `security`
- `data`

Exemplos:

```text
fix(operacao-dia): corrigir exibição de pedidos confirmados
feat(extras): integrar extras à operação do dia
docs(manual): atualizar regras de negócio do travamento
data(sharepoint): documentar campos internos da lista Pedidos
security(msal): documentar cache de autenticação entre páginas
```

## 3. Corpo de commit para mudanças relevantes

```text
Contexto:
- ...

Alterações:
- ...

Regra de negócio:
- ...

Documentação:
- ...

Testes:
- ...
```

## 4. Template de PR

```md
## Resumo

## Motivo

## Arquivos alterados

## Módulos impactados

## SharePoint impactado
Lista:
Campos:

## Regras de negócio impactadas

## Documentação atualizada
Git:
- [ ] ...

Interna Homy:
- [ ] ...

## Testes realizados
- [ ] Login
- [ ] Leitura SharePoint
- [ ] Gravação SharePoint
- [ ] Admin
- [ ] Cozinha
- [ ] Marcar Refeição
- [ ] Regressão do módulo relacionado

## Commit sugerido
```

## 5. Regra de ouro

Não subir alteração sem entender:

- Causa raiz.
- Arquivo correto.
- Lista SharePoint impactada.
- Regra de negócio envolvida.
- Teste necessário.
- Documentação que precisa ser atualizada.
