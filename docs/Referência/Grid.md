---
tags: [referência, grid]
---

# Grid

Crie com `grid = sap.get_grid(target_index=0)`. `target_index` escolhe qual grid usar quando a tela contém mais de um.

## Leitura

- `count_rows()` retorna a quantidade total de linhas.
- `get_column_id(column_name)` transforma o título visível no ID técnico da coluna.
- `get_cell_value(index, column_id)` retorna uma célula; `index` começa em `0`.
- `get_row(row)` retorna os valores de uma linha.
- `get_columns(*column_id)` retorna uma lista para uma coluna ou `{"header", "content"}` para várias.
- `get_content()` retorna todas as colunas e linhas em `{"header", "content"}`.

## Interação

- `select_all_content()` seleciona todos os itens.
- `select_column(column_id)` seleciona uma coluna.
- `click_cell(index, column_id)` seleciona e dá duplo clique.
- `open_cell_modal(index, column_id)` abre a ação/modal da célula.
- `press_button(field_name, skip_error=False)` pressiona botão pelo tooltip/texto.
- `press_nested_button(*nested_fields, skip_error=False)` abre um botão/menu e escolhe o item seguinte; normalmente use dois textos, por exemplo `("Exportar", "Planilha")`.
- `select_layout(layout)` aplica um layout quando o pop-up de seleção de layout já está aberto.

O grid percorre páginas ao coletar dados; em telas muito grandes, considere solicitar `get_columns` em vez de `get_content`.
