# Colaboradores

## Objetivo

Adicionar a base correta de colaboradores para todos os relatórios e operações futuras.

## Incluído nesta etapa

- Campo `Centro_Custo` no cadastro de colaborador
- Suporte a `Centro_Custo` no `sharepoint.js`
- Criação/edição de colaborador com Centro de Custo
- Botão Editar
- Botão Desativar
- Exibição de Centro de Custo na tabela
- Arquivo modular `admin/js/admin-colaboradores.js`

## Arquivos alterados

- `sharepoint.js`
- `admin/index.html`
- `admin/js/admin-colaboradores.js`

## Observação importante

A exclusão foi tratada como **desativação** (`Ativo = false`) para não perder histórico de pedidos/refeições.

Isso é necessário porque relatórios por colaborador, folha e centro de custo dependem do histórico.
