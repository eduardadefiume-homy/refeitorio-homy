# Checklist de Testes

Sistema: Refeitório Homy  
Tipo: Qualidade  
Responsável: TI Homy  
Última revisão: 18/06/2026  
Versão: 1.0  
Status: Pronto para revisão interna

## 1. Antes de qualquer commit

- [ ] Identifique o arquivo alterado.
- [ ] Identifique o módulo afetado.
- [ ] Identifique listas SharePoint impactadas.
- [ ] Identifique campos SharePoint impactados.
- [ ] Verifique se mudou regra de negócio.
- [ ] Atualize documentação impactada.
- [ ] Atualize `CHANGELOG.md`.

## 2. JavaScript

- [ ] Rodar `node --check arquivo.js`.
- [ ] Verificar funções duplicadas.
- [ ] Verificar chamadas a funções inexistentes.
- [ ] Testar no navegador com F12 aberto.
- [ ] Console sem erros vermelhos.

## 3. Admin

- [ ] Login abre e autentica.
- [ ] Dashboard carrega.
- [ ] Troca de semana funciona.
- [ ] Cardápio carrega e salva.
- [ ] Pedidos carrega.
- [ ] Operação do Dia carrega por dia.
- [ ] Colaboradores carrega.
- [ ] Extras carrega.
- [ ] Valores carrega.
- [ ] Relatórios carrega.
- [ ] Configurações carrega.

## 4. Fluxo completo de pedido

- [ ] Colaborador ativo aparece na marcação.
- [ ] Pedido é salvo em `Pedidos`.
- [ ] Pedido aparece em Admin > Pedidos.
- [ ] Pedido aparece em Operação do Dia.
- [ ] Pedido aparece na Cozinha.
- [ ] CheckIn grava retirada.

## 5. Regressão de status

Testar pelo menos um pedido com:

- [ ] Confirmado.
- [ ] Cancelado.
- [ ] Não vai almoçar.
- [ ] Férias/Ferias.
- [ ] Afastado.
- [ ] Travado/Bloqueado.
- [ ] Extra.

## 6. SharePoint

- [ ] Lista existe com nome exato.
- [ ] Campos internos existem.
- [ ] Gravação cria item correto.
- [ ] Edição atualiza item correto.
- [ ] Exclusão remove ou desativa conforme regra.

## 7. Pós-publicação

- [ ] Aguardar GitHub Pages propagar.
- [ ] Fazer Ctrl+Shift+R.
- [ ] Testar em aba anônima.
- [ ] Testar em outro navegador se possível.
- [ ] Registrar evidência/print quando for alteração crítica.
