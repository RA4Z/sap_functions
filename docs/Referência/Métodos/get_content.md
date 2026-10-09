---
tags: [método, retornos, extração]
---

# `get_content`

`Grid`, `Table`, `Tree` e `Label` implementam `get_content()`, mas o objeto deve ser obtido antes: `sap.get_grid()`, `sap.get_table()`, `sap.get_tree()` ou `sap.get_label()`.

O contrato de retorno é `{"header": list, "content": list[list]}`. Nunca presuma que o retorno é uma lista de dicionários; converta-o se esse formato for necessário.

```python
resultado = sap.get_grid().get_content()
linhas = [dict(zip(resultado["header"], linha)) for linha in resultado["content"]]
```

Veja [[Referência/Retornos de dados/Visão geral|Retornos de dados]] para os formatos por componente.
