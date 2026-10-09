---
tags: [referência, retornos, grid, ia]
---

# Retornos: Grid

## `grid.get_content()`

Retorna todas as colunas e linhas como uma tabela serializável:

```python
{
    "header": ["Material", "Centro", "Quantidade"],
    "content": [
        ["MAT-001", "1000", "12"],
        ["MAT-002", "2000", None]
    ]
}
```

`content[i][j]` corresponde à coluna `header[j]`. O retorno não é uma lista de dicionários.

## `grid.get_columns("Quantidade")`

Ao solicitar exatamente uma coluna, retorna uma lista simples:

```python
["12", "8", None]
```

Com duas ou mais colunas, volta ao formato tabular:

```python
{
    "header": ["Material", "Quantidade"],
    "content": [["MAT-001", "12"], ["MAT-002", "8"]]
}
```

## Outras leituras

`get_cell_value(0, "MENGE")` retorna o valor individual, normalmente `str`. `get_row(0)` retorna os valores da linha como `list`. `count_rows()` retorna `int`; `get_column_id("Quantidade")` retorna o identificador técnico `str` da coluna.
