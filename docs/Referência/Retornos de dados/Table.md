---
tags: [referência, retornos, table, ia]
---

# Retornos: Table

`table.get_content()` usa o mesmo contrato tabular do Grid:

```python
{
    "header": ["Documento", "Status"],
    "content": [["900001", "Aberto"], ["900002", "Concluído"]]
}
```

`table.get_columns("Status")` retorna `list`, por exemplo `["Aberto", "Concluído"]`. Ao solicitar várias colunas, retorna `{"header": [...], "content": [[...], ...]}`.

`get_cell_value(row, column)` retorna uma célula individual e `count_visible_rows()` retorna `int`. Os métodos de alteração e seleção (`write_cell_value`, `flag_cell`, `click_cell`, `select_entire_row` e `unselect_entire_row`) retornam `None` quando são bem-sucedidos.
