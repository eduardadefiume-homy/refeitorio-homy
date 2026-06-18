# ADR-0001 - SharePoint como fonte única de verdade

Data: 18/06/2026  
Status: Aprovado

## Contexto

O Refeitório Homy não possui backend próprio. O sistema precisa armazenar cardápio, pedidos, colaboradores, extras, ausências, valores e check-ins.

## Decisão

Usar SharePoint Lists como fonte única de verdade para todos os dados de negócio.

## Consequências

- Toda escrita deve ocorrer no SharePoint.
- A interface só atualiza após confirmação.
- É proibido usar armazenamento local como banco.
- Mudanças em listas/campos precisam ser documentadas.
