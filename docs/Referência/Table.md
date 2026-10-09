---
tags: [referência, table]
---

# Table

Crie com `table = sap.get_table(target_index=0)`. A maioria das operações por linha usa a linha visível, começando em `0`.

- `count_visible_rows(skip_error=False)` retorna as linhas atualmente visíveis.
- `get_cell_value(row, column, skip_error=False)` lê uma célula.
- `write_cell_value(row, column, desired_text, skip_error=False)` altera uma célula editável.
- `flag_cell(row, column, desired_operator, skip_error=False)` marca/desmarca uma célula checkbox.
- `click_cell(row, column, skip_error=False)` foca e abre a célula com duplo clique.
- `select_entire_row(absolute_row, skip_error=False)` e `unselect_entire_row(absolute_row, skip_error=False)` trabalham com a linha absoluta, não a visível.
- `get_content(skip_error=False)` extrai todas as linhas em `{"header", "content"}`.
- `get_columns(*column_text, skip_error=False)` extrai as colunas pelos títulos; uma coluna retorna lista, várias retornam cabeçalho e conteúdo.
