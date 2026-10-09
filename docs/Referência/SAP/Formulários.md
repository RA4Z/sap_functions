---
tags: [referência, sap, formulários]
---

# SAP: formulários

Os métodos abaixo localizam controles pelo texto do rótulo. `target_index=0` usa a primeira ocorrência repetida; `selected_tab=0` usa a primeira aba (ou aceite o nome da aba); `skip_error=False` lança uma exceção quando o controle não é encontrado ou a ação falha.

## `write_text_field(field_name, desired_text, target_index=0, selected_tab=0, skip_error=False)`

Preenche o campo de texto associado a `field_name`, substituindo seu valor atual.

```python
sap.write_text_field("Material", "MAT-001")
```

## `write_text_field_until(field_name, desired_text, ...)`

Preenche o campo “até” associado a um rótulo de intervalo. Use para o fim de uma faixa de datas, materiais e afins.

```python
sap.write_text_field_until("Data", "31.12.2026")
```

## `choose_text_combo(field_name, desired_text, ...)`

Seleciona `desired_text` em um combo box ao lado do rótulo. Falha se a opção não estiver disponível.

## `flag_field(field_name, desired_operator, ...)`

Marca (`True`) ou desmarca (`False`) um checkbox associado ao rótulo.

## `option_field(field_name, ...)`

Seleciona o radio button associado ao rótulo. Não recebe um booleano: a operação é sempre selecionar.

## Métodos `*_at_side`

`flag_field_at_side(field_name, desired_operator, side_index=0, ...)` e `option_field_at_side(field_name, side_index=0, ...)` usam o controle em uma posição relativa ao rótulo. `side_index` positivo busca à direita e negativo à esquerda.

## `press_button(field_name, ...)`

Pressiona um botão cujo texto ou tooltip contenha `field_name`.

## `clean_all_fields(selected_tab=0, skip_error=False)`

Limpa os campos de texto encontrados na área da tela/aba. Use com cuidado: ele remove o conteúdo de todos os campos editáveis detectados.

## `find_text_field(field_name, selected_tab=0)`

Retorna `True` se um texto/rótulo for encontrado; retorna `False` caso contrário. É útil para decidir o fluxo sem capturar exceções.

## `get_text_at_side(field_name, side_index, target_index=0, selected_tab=0)`

Retorna o texto do componente relativo ao rótulo. `side_index` é obrigatório porque define qual vizinho ler.

## `set_focus(field_name, side_index=0, target_index=0, selected_tab=0)`

Move o foco para o controle relativo ao rótulo. Combine com `open_focused_field_modal()` para abrir a ajuda de valor.
