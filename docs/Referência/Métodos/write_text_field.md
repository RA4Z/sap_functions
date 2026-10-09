---
tags: [método, sap, formulário]
---

# `write_text_field`

Assinatura: `sap.write_text_field(field_name, desired_text, target_index=0, selected_tab=0, skip_error=False)`.

Localiza um campo de texto pelo rótulo `field_name` e substitui seu conteúdo por `desired_text`. `target_index` seleciona a ocorrência do rótulo (começa em zero). `selected_tab` aceita índice ou nome de aba. Não retorna valor; lança `Write text field failed.` se não conseguir concluir, salvo com `skip_error=True`.

```python
sap.write_text_field("Material", "MAT-001", selected_tab="Seleção")
```
