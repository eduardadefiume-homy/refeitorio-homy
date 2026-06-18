# Troubleshooting

Sistema: Refeitório Homy  
Tipo: Suporte técnico  
Responsável: TI Homy  
Última revisão: 18/06/2026  
Versão: 1.1  
Status: Pronto para revisão interna

## 1. Diagnóstico padrão

Sempre iniciar por:

1. Abrir F12 no navegador.
2. Ver aba Console.
3. Ver aba Network.
4. Confirmar se arquivos JS carregaram.
5. Confirmar se Graph API retornou 200/201/204.
6. Validar item diretamente no SharePoint.

## 2. Erros comuns

| Sintoma | Causa provável | Ação |
|---|---|---|
| `AADSTS50011` | redirect URI divergente | validar App Registration e `redirectUri` fixo |
| `popup_window_error` | popup aberto sem clique do usuário | login deve ser acionado por clique |
| `ReferenceError: SP is not defined` | `sharepoint.js` não carregou | validar ordem dos scripts |
| `ReferenceError: AdminX is not defined` | módulo admin não carregou | validar script e SyntaxError anterior |
| `SyntaxError` | erro no JS | rodar `node --check arquivo.js` |
| Lista não encontrada | nome SharePoint divergente | validar displayName exato |
| 401/403 Graph | token/permissão | validar login, escopos e consentimento |
| Tela mostra dado antigo | cache GitHub Pages/navegador | Ctrl+Shift+R ou versionar script |
| Pedido sumiu da operação | filtro semana/dia/status ou registro ausente | validar lista Pedidos no SharePoint |
| Cardápio duplicado | edição criou novo item | corrigir para atualizar por Semana_id+Dia+Opcao |

## 3. Pedido não aparece na Operação do Dia

Validar no SharePoint:

- Existe na lista `Pedidos`?
- `Semana_id` é a semana selecionada?
- `Dia` corresponde ao filtro?
- `Status` está preenchido?
- `Colaborador_id` existe?
- Pedido foi excluído?
- Está com status cancelado ou ausência?

## 4. Colaborador não aparece na Marcação

Validar:

- Existe em `Colaboradores`?
- `Ativo` está verdadeiro?
- Nome está preenchido em `Nome` ou `Title`?
- Filtro de busca está escondendo?
- Erro no console impede carregamento?

## 5. Cozinha não atualiza

Validar:

- Pedido existe e está confirmado/produtivo?
- Dia atual está correto?
- Semana ISO está correta?
- Há ausência ativa para o colaborador?
- CheckIn já foi registrado?

## 6. Erros conhecidos que não devem retornar

| Erro antigo | Correção oficial |
|---|---|
| Editar cardápio criava novo prato | editar sempre atualiza registro existente |
| Cancelar modal salvava alterações | cancelar apenas fecha |
| Lista errada para valores | usar `Valores de Refeição` |
| localStorage usado como banco | proibido; SharePoint é fonte de verdade |
