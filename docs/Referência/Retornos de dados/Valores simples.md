---
tags: [referência, retornos, ia]
---

# Retornos: valores simples

```python
sap.find_text_field("Material")  # bool: True ou False
sap.get_text_at_side("Centro", 1)  # str: "1000"
sap.get_footer_message()  # str: "Nenhuma seleção encontrada"

grid.count_rows()  # int: 42
table.count_visible_rows()  # int: 16
grid.get_column_id("Quantidade")  # str: "MENGE"
```

Métodos que interagem com a tela — como `write_text_field`, `press_button`, `sendVKey`, `select_node` e `save_file` — retornam `None` se concluírem. Não use o retorno desses métodos como confirmação de dados; valide a tela, leia o rodapé ou extraia o resultado depois da ação.
