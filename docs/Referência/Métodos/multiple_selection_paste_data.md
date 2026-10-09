---
tags: [método, sap, seleção-múltipla]
---

# `multiple_selection_paste_data`

Assinatura: `sap.multiple_selection_paste_data(data, delete_values=False, skip_error=False)`.

Importa `data`, uma `list[str]`, na janela de seleção múltipla que já deve estar aberta. Use primeiro `multiple_selection_field`. Quando `delete_values=True`, os dados são inseridos na aba de exclusão. Retorna `None` e usa um arquivo temporário em `C:/Temp`.

```python
sap.multiple_selection_field("Material")
sap.multiple_selection_paste_data(["MAT-001", "MAT-002"])
```
