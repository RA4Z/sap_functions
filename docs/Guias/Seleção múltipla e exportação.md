---
tags:
  - guia
  - seleção-múltipla
  - exportação
---

# Seleção múltipla e exportação

## Inserir muitos valores

Abra a seleção múltipla do campo e envie uma lista de textos:

```python
materiais = ["MAT-001", "MAT-002", "MAT-003"]

sap.multiple_selection_field("Material")
sap.multiple_selection_paste_data(materiais)
```

Para registrar valores que devem ser excluídos do resultado, use a área de exclusão:

```python
sap.multiple_selection_field("Material")
sap.multiple_selection_paste_data(["MAT-OBSOLETO"], delete_values=True)
```

Durante a operação, a biblioteca cria temporariamente um arquivo em `C:/Temp` e o remove ao final. Portanto, o usuário que executa o script precisa ter permissão de escrita nessa pasta.

## Abrir ajuda de valor

Quando o foco já estiver no campo, abra a ajuda de pesquisa do SAP:

```python
sap.set_focus("Centro")
sap.open_focused_field_modal()
```

## Salvar uma lista em arquivo

Com um resultado aberto, use:

```python
sap.save_file("relatorio", r"C:\Users\usuario\Downloads", type_of_file="xls")
```

Para exportar texto, informe `type_of_file="txt"`. O parâmetro `option` controla o formato de texto: `0` para não convertido, `1` para texto com tabulações e `2` para RTF.

> [!warning]
> `save_file` segue a estrutura padrão de menu do SAP para listas. Ela pode não funcionar em transações com um menu de exportação diferente.

Para percorrer outra ação do menu manualmente, use `navigate_into_menu_header("Lista", "Exportar", "Arquivo local")`.
