# ADR-0004 - localStorage apenas para cache MSAL

Data: 18/06/2026  
Status: Aprovado

## Contexto

O sistema possui várias páginas estáticas na mesma origem. Sem cache compartilhado, o usuário pode precisar autenticar novamente ao navegar.

## Decisão

Permitir `localStorage` apenas como cache técnico do MSAL, nunca como banco de dados de negócio.

## Consequências

- Sessão pode ser compartilhada entre páginas.
- Dados de negócio continuam 100% no SharePoint.
- Qualquer uso de `localStorage` fora de autenticação deve ser recusado ou documentado como exceção técnica.
