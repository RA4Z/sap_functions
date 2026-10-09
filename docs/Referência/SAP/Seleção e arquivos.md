---
tags: [referência, sap, seleção, arquivos]
---

# SAP: seleção múltipla e arquivos

Nota individual: [[Referência/Métodos/multiple_selection_paste_data]].

## `multiple_selection_field(field_name, target_index=0, selected_tab=0, skip_error=False)`

Abre a janela de seleção múltipla ligada ao campo. Ela não insere valores por si só; chame `multiple_selection_paste_data` em seguida.

## `multiple_selection_paste_data(data, delete_values=False, skip_error=False)`

Importa uma lista de strings para a janela de seleção múltipla aberta. A lista é gravada temporariamente em `C:/Temp`, importada pelo SAP e removida ao final. `delete_values=True` muda para a aba de exclusões antes da importação.

```python
sap.multiple_selection_field("Centro")
sap.multiple_selection_paste_data(["1000", "2000"])
```

## `open_focused_field_modal()`

Envia F4 para a janela atual, abrindo a ajuda/consulta do campo que está com foco. Execute `set_focus` antes quando necessário.

## `save_file(file_name, path, option=0, type_of_file="txt", skip_error=False)`

Exporta a lista atual para `path`, adicionando a extensão `type_of_file` ao nome. Para `txt`, `option=0` é não convertido, `1` é texto com tabulações e `2` é RTF. Se `type_of_file` contém `xls`, segue o fluxo padrão de exportação Excel do SAP.

```python
sap.save_file("estoque", r"C:\Relatorios", type_of_file="xls")
```

Esse método pressupõe a estrutura padrão de menu de lista SAP; trate a exceção em transações customizadas.

## `get_footer_message()`

Retorna o texto da barra de status (`wnd[0]/sbar`). Use-o para enriquecer logs e mensagens de erro.
