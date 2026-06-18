# ADR-0003 - MSAL com loginPopup

Data: 18/06/2026  
Status: Aprovado

## Contexto

O GitHub Pages pode ter problemas com redirects dinâmicos em múltiplas páginas.

## Decisão

Usar MSAL com `loginPopup` e `redirectUri` fixo na raiz do app.

## Consequências

- Evita erros de redirect URI em páginas diferentes.
- Login deve ser iniciado por ação do usuário quando popup for necessário.
- Alterações nesse fluxo exigem teste completo.
