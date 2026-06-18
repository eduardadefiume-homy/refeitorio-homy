# ADR-0002 - GitHub Pages sem backend próprio

Data: 18/06/2026  
Status: Aprovado

## Contexto

O sistema precisava ser simples, barato e fácil de publicar internamente.

## Decisão

Hospedar o frontend no GitHub Pages, sem backend próprio.

## Consequências

- Publicação simples por commit.
- Toda lógica de dados fica em `sharepoint.js` via Graph API.
- Não há processamento server-side.
- Autenticação e permissões dependem do Microsoft Entra ID.
