---
tags: [referência, sap, navegação]
---

# SAP: navegação

## `SAP(window=0)`

Cria a conexão com uma sessão SAP GUI já aberta. `window` é o índice da sessão a usar; `0` usa a primeira. A biblioteca confirma que há um usuário autenticado e, se estiver em `S000`, navega para a tela principal. Falha se não houver uma conexão ou login SAP disponível.

```python
sap = SAP()          # primeira sessão
outra_sessao = SAP(1)
```

## `select_transaction(transaction)`

Abre uma transação. O código é convertido para maiúsculas e a biblioteca confirma que a transação ativa corresponde ao solicitado. Se não corresponder, lança `Exception` com a mensagem do rodapé SAP.

```python
sap.select_transaction("coois")
```

## `select_main_screen(skip_error=False)`

Vai para `SESSION_MANAGER`, a tela principal. Com `skip_error=True`, uma falha é ignorada; sem ele, lança `Select main screen failed.`.

## `change_active_tab(selected_tab, skip_error=False)`

Ativa uma aba. `selected_tab` pode ser o índice (`0` é a padrão) ou parte/nome do texto da aba. É chamado automaticamente pelos métodos de formulário quando `selected_tab` é informado.

```python
sap.change_active_tab(1)
sap.change_active_tab("Seleção")
```

## `run_actual_transaction(skip_error=False)`

Executa a tela atual. Primeiro envia Enter; se o título da janela não mudar, envia F8. A intenção é funcionar tanto em telas que validam com Enter quanto em telas de relatório que executam com F8.

## `navigate_into_menu_header(*nested_path)`

Percorre o menu superior pelo texto, na ordem recebida, e seleciona o último item.

```python
sap.navigate_into_menu_header("Lista", "Exportar", "Arquivo local")
```

Os textos precisam existir na interface SAP atual; esta função não tem `skip_error`.

## `insert_variant(variant_name, skip_error=False)`

Abre o diálogo de variante, limpa o campo de criador, informa `variant_name` e confirma. É uma conveniência baseada no fluxo padrão de variantes e pode não ser compatível com todas as transações.

Próximo: [[Referência/SAP/Formulários|preencher campos e controles]].
