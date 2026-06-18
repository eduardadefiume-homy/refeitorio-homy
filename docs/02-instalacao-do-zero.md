# Instalação e Replicação do Zero

Sistema: Refeitório Homy  
Tipo: Guia técnico  
Responsável: TI Homy  
Última revisão: 18/06/2026  
Versão: 1.0  
Status: Pronto para revisão interna

## 1. Objetivo

Permitir que outra pessoa da TI consiga recriar o ambiente do Refeitório Homy com base no Git, no Microsoft Entra ID e no SharePoint.

## 2. Pré-requisitos

- Conta Microsoft 365 da Homy.
- Permissão para criar/editar App Registration no Microsoft Entra ID.
- Permissão no SharePoint do site do Refeitório.
- Acesso ao repositório GitHub `eduardadefiume-homy/refeitorio-homy`.
- Acesso ao GitHub Pages do repositório.

## 3. Criar ou validar App Registration

| Campo | Valor atual |
|---|---|
| App Name | Refeitorio Homy |
| Client ID | `aa37acf9-f3bd-4d1e-968a-fde57f79094c` |
| Tenant ID | `a2850abc-334a-4805-b6b2-420b4aef68a9` |
| Redirect URI | `https://eduardadefiume-homy.github.io/refeitorio-homy/index.html` |
| Auth Flow | `loginPopup` |
| Scopes | `Sites.ReadWrite.All`, `User.Read` |

Atenção: o `redirectUri` deve ser fixo na raiz do app. Não usar `window.location.pathname`.

## 4. Criar ou validar SharePoint

Site atual:

```text
homyquimica.sharepoint.com/sites/Refeitrio-Homy
```

Listas obrigatórias:

- Cardapio
- Pedidos
- Colaboradores
- Extras
- Configurações
- Valores de Refeição
- Ausencias do Refeitorio
- CheckIn

Os nomes devem ser exatos, inclusive acentos quando existirem.

## 5. Publicar no GitHub Pages

1. Subir arquivos no repositório.
2. Ativar GitHub Pages na branch principal.
3. Garantir que `index.html` esteja na raiz.
4. Validar acesso ao hub.
5. Validar cada tela individualmente.

## 6. Configuração inicial mínima

1. Cadastrar colaboradores ativos na lista `Colaboradores`.
2. Criar registros básicos na lista `Configurações`:
   - `cardapio_liberado`
   - `marcacao_liberada`
   - `pedidos_liberados`
   - `cardapio_visivel`
   - `prazo_limite`
3. Cadastrar cardápio da semana pelo Admin.
4. Liberar marcação pelo Dashboard.
5. Fazer teste real de pedido.
6. Confirmar que pedido aparece em Pedidos, Operação do Dia e Cozinha.

## 7. Checklist de validação do ambiente

- [ ] Login Microsoft abre em popup.
- [ ] Conta Homy autentica corretamente.
- [ ] `sharepoint.js` obtém token.
- [ ] Graph API lê listas do SharePoint.
- [ ] Cardápio salva e atualiza registro existente.
- [ ] Pedido é salvo com Semana_id, Dia, Opção e Colaborador.
- [ ] Cozinha exibe pedido confirmado do dia.
- [ ] CheckIn grava retirada.
- [ ] Admin carrega módulos sem erro no console.
- [ ] Relatórios conseguem ler Pedidos.

## 8. Pontos a validar manualmente

- Nomes internos reais dos campos no SharePoint.
- Permissões exatas concedidas ao App Registration.
- Funções chamadas pelos módulos admin e presentes em `sharepoint.js`.
- Se há necessidade de versionar scripts no `admin/index.html` com `?v=YYYYMMDD` para evitar cache.
