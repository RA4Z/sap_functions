---
tags:
  - guia
  - extração
---

# Ler resultados

Após executar a transação, escolha o objeto que corresponde ao controle exibido.

## Grid ALV

```python
grid = sap.get_grid()
dados = grid.get_content()
quantidades = grid.get_columns("Quantidade")
valor = grid.get_cell_value(0, "MENGE")
```

`get_content()` devolve um dicionário com `header` (lista de títulos) e `content` (lista de linhas). Para descobrir o ID técnico de uma coluna a partir de seu título, use `grid.get_column_id("Quantidade")`.

## Tabela clássica

```python
tabela = sap.get_table()
dados = tabela.get_content()
tabela.write_cell_value(0, 1, "Novo valor")
tabela.select_entire_row(5)
```

Os métodos que recebem linha e coluna da tabela trabalham com índices numéricos. A linha de `select_entire_row` é absoluta; as outras operações usam a linha visível na tela.

## Árvore e nós

```python
tree = sap.get_tree()
dados = tree.get_content()
hierarquia = tree.get_hierarchy_content()

node = sap.get_node()
node.select_node("Planejamento")
node.expand_selected_node()
```

## Relatório por labels

```python
labels = sap.get_label()
dados = labels.get_content()
```

Veja a lista detalhada de ações em [[Referência/Componentes SAP|Componentes SAP]].
