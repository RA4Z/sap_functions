---
tags: [referência, tree, node, label]
---

# Tree, Node e Label

## Tree

Obtenha com `tree = sap.get_tree()`.

- `get_content(skip_error=False)` retorna todas as colunas e linhas em `{"header", "content"}`.
- `get_columns(*column_text)` retorna apenas as colunas cujos títulos foram informados; uma coluna retorna lista, várias retornam dicionário.
- `get_hierarchy_content(skip_error=False)` retorna também `Row ID` e `Parent ID`, permitindo reconstruir a relação hierárquica fora do SAP.

## Node

Obtenha com `node = sap.get_node(target_index=0)`.

- `get_content()` lista todos os textos de nó.
- `select_node(node_text, target_index=0, skip_error=False)` seleciona o primeiro nó cujo texto contém `node_text`; escolha outra ocorrência por índice.
- `expand_selected_node()` expande o nó atualmente selecionado.
- `click_selected_node()` dá duplo clique no nó atualmente selecionado.

Selecione o nó antes de expandir ou clicar.

## Label

Obtenha com `label = sap.get_label()`.

- `get_content()` percorre a tela de labels e retorna `{"header", "content"}`.
- `get_all_screen_labels()` retorna as linhas encontradas, sem separar cabeçalho.

Esses métodos são indicados para relatórios SAP clássicos renderizados como labels, não para Grid ALV ou TableControl.
