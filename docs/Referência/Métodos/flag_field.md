---
tags: [método, sap, checkbox]
---

# `flag_field`

Assinatura: `sap.flag_field(field_name, desired_operator, target_index=0, selected_tab=0, skip_error=False)`.

Marca o checkbox associado ao rótulo quando `desired_operator=True` e desmarca quando `False`. A versão `flag_field_at_side` adiciona `side_index` para escolher um checkbox deslocado em relação ao rótulo.

```python
sap.flag_field("Incluir concluídas", True)
sap.flag_field("Somente pendentes", False)
```
