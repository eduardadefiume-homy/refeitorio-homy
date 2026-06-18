# Fluxos Operacionais

Sistema: Refeitório Homy  
Tipo: Operação  
Responsável: TI Homy  
Última revisão: 18/06/2026  
Versão: 1.1  
Status: Pronto para revisão interna

## 1. Fluxo semanal da Luana/Admin

1. Acessar Admin.
2. Cadastrar ou revisar cardápio da semana.
3. Verificar valores e configurações.
4. Liberar marcação no Dashboard.
5. Acompanhar pendentes.
6. No prazo definido, travar pendentes se necessário.
7. Acompanhar Operação do Dia.
8. Ajustar status manualmente quando necessário.
9. Conferir extras e ausências.
10. Gerar relatórios.

## 2. Fluxo do colaborador

1. Acessar Marcar Refeição.
2. Buscar nome.
3. Selecionar colaborador.
4. Escolher opção por dia da semana.
5. Confirmar marcação.
6. Sistema grava pedido no SharePoint.

## 3. Fluxo da cozinha

1. Acessar Cozinha.
2. Conferir cardápio do dia.
3. Buscar colaborador na lista.
4. Confirmar retirada.
5. Sistema grava CheckIn.
6. Colaborador confirmado fica marcado como retirado.

## 4. Fluxo de extras

1. Admin acessa Extras.
2. Seleciona dia e tipo.
3. Informa nome/opção/observação ou usa pré-definido.
4. Sistema grava extra.
5. Quando integração estiver ativa, sistema cria/relaciona pedido correspondente.
6. Extra aparece na Operação do Dia e na Cozinha.

## 5. Fluxo de ausências

1. Admin cadastra ausência com colaborador, motivo, início e fim.
2. Sistema grava ausência ativa no SharePoint.
3. Cozinha consulta ausências do dia.
4. Colaborador ausente aparece travado/indisponível.

## 6. Fluxo de valores e NF

1. Admin cadastra período de valor.
2. Mantém apenas um valor ativo.
3. Upload da NF Vascon quando aplicável.
4. Relatório/reconciliação cruza pedidos confirmados com valor unitário.
5. Divergências devem ser analisadas antes de fechamento.

## 7. Fluxo de incidente

1. Usuário reporta erro com print.
2. TI identifica módulo/tela.
3. TI abre F12 e coleta Console/Network.
4. TI valida dados no SharePoint.
5. TI corrige código/documentação.
6. TI testa, registra changelog e faz commit.
7. TI orienta usuário a recarregar com cache limpo.
