---
tags: [referência, retornos, tree, ia]
---

# Retornos: Tree

`tree.get_content()` retorna uma tabela de strings:

```python
{
    "header": ["Centro", "Descrição"],
    "content": [["1000", "Matriz"], ["2000", "Filial"]]
}
```

`tree.get_columns("Descrição")` retorna uma lista para uma coluna ou uma tabela para várias colunas.

`tree.get_hierarchy_content()` acrescenta os identificadores de estrutura antes das colunas originais:

```python
{
    "header": ["Row ID", "Parent ID", "Centro", "Descrição"],
    "content": [["1", "", "1000", "Matriz"], ["2", "1", "2000", "Filial"]]
}
```

Use `Parent ID` para reconstruir a árvore: ele aponta para o `Row ID` do pai.
