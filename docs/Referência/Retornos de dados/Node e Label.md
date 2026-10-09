---
tags: [referência, retornos, node, label, ia]
---

# Retornos: Node e Label

`node.get_content()` retorna somente os textos dos nós, em ordem:

```python
["Planejamento", "Produção", "Relatórios"]
```

`label.get_content()` retorna tabela com cabeçalho e linhas:

```python
{
    "header": ["Material", "Estoque"],
    "content": [["MAT-001", "12"], ["MAT-002", "8"]]
}
```

`label.get_all_screen_labels()` retorna apenas as linhas, sem a chave `header`:

```python
[["MAT-001", "12"], ["MAT-002", "8"]]
```
