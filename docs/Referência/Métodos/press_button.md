---
tags: [método, sap, botão]
---

# `press_button`

Assinatura: `sap.press_button(field_name, target_index=0, selected_tab=0, skip_error=False)`.

Pressiona o botão cujo texto ou tooltip contém `field_name`. Em telas com mais de um botão correspondente, `target_index=0` escolhe o primeiro. Retorna `None`; a falha lança `Press button failed.` por padrão.

```python
sap.press_button("Executar")
```
