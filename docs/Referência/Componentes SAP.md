---
tags:
  - referência
  - componentes
---

# Componentes SAP

## Grid

Obtenha com `sap.get_grid()`. Use `get_content()` para coletar a grade inteira, `get_columns("Nome")` para colunas específicas, `get_row(indice)` para uma linha e `get_cell_value(indice, coluna_id)` para uma célula. `get_column_id("Título")` encontra o ID técnico necessário para algumas chamadas. `count_rows()` conta as linhas.

Para interagir: `select_all_content()`, `select_column(column_id)`, `click_cell(indice, column_id)`, `open_cell_modal(indice, column_id)`, `press_button("Tooltip")` e `press_nested_button("Menu", "Item")`. Se a janela de layouts já estiver aberta, `select_layout("Nome do layout")` seleciona o layout.

## Table

Obtenha com `sap.get_table()`. A tabela trabalha com linha visível para `get_cell_value(row, column)`, `write_cell_value(row, column, texto)`, `click_cell(row, column)` e `flag_cell(row, column, True/False)`. `count_visible_rows()` informa quantas linhas cabem na tela. Para seleção de linha inteira, `select_entire_row(absolute_row)` e `unselect_entire_row(absolute_row)` recebem a linha absoluta. Use `get_content()` ou `get_columns("Título")` para extrair dados.

## Tree

Obtenha com `sap.get_tree()`. `get_content()` retorna cabeçalho e conteúdo; `get_columns("Título")` limita a extração às colunas desejadas. `get_hierarchy_content()` também inclui identificadores de linha e de pai.

## Node

Obtenha com `sap.get_node()`. `select_node("texto", target_index=0)` localiza um nó por texto; em seguida use `click_selected_node()` ou `expand_selected_node()`. `get_content()` lista os nomes dos nós. Passe `skip_error=True` em `select_node` somente se a ausência do nó for esperada.

## Label

Obtenha com `sap.get_label()`. `get_content()` extrai cabeçalho e conteúdo de telas baseadas em labels; `get_all_screen_labels()` retorna as linhas encontradas.

> [!note]
> Os métodos de coleta retornam, em geral, um dicionário com `header` e `content`. Quando apenas uma coluna é solicitada, `get_columns(...)` pode retornar uma lista.

Para navegação direta por controles SAP GUI, veja [[Referência/Teclas virtuais|sendVKey e teclas virtuais]].
