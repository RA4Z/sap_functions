---
tags:
  - referência
  - api
---

# Classe SAP

`SAP` é a porta de entrada da biblioteca. Ela localiza controles pela interface SAP GUI e executa ações comuns por texto, evitando a necessidade de lidar com IDs na maior parte dos casos.

Para detalhes completos e exemplos, escolha a área: [[Referência/SAP/Navegação|Navegação]], [[Referência/SAP/Formulários|Formulários]] ou [[Referência/SAP/Seleção e arquivos|Seleção múltipla e arquivos]].

## Navegação

| Método | Uso |
| --- | --- |
| `select_transaction(transaction)` | Abre uma transação SAP. |
| `select_main_screen()` | Volta à tela principal. |
| `run_actual_transaction()` | Tenta `Enter` e, se a tela não mudar, `F8`. |
| `change_active_tab(selected_tab)` | Seleciona uma aba por índice ou nome. |
| `navigate_into_menu_header(*nested_path)` | Percorre itens do menu superior. |
| `insert_variant(variant_name)` | Abre e aplica uma variante; não é compatível com todas as transações. |

## Campos e botões

| Método | Uso |
| --- | --- |
| `write_text_field(field_name, desired_text, ...)` | Preenche o campo ao lado de um rótulo. |
| `write_text_field_until(field_name, desired_text, ...)` | Preenche o campo "até" correspondente. |
| `choose_text_combo(field_name, desired_text, ...)` | Escolhe texto em um combo box. |
| `flag_field(field_name, desired_operator, ...)` | Marca ou desmarca checkbox. |
| `flag_field_at_side(field_name, desired_operator, side_index, ...)` | Marca/desmarca checkbox deslocado em relação ao rótulo. |
| `option_field(field_name, ...)` | Seleciona um radio button. |
| `option_field_at_side(field_name, side_index, ...)` | Seleciona radio button deslocado em relação ao rótulo. |
| `press_button(field_name, ...)` | Pressiona botão pelo texto ou tooltip. |
| `set_focus(field_name, side_index, ...)` | Coloca o foco em um controle. |
| `open_focused_field_modal()` | Abre a ajuda de valor (F4) do controle em foco. |
| `set_focus(field_name, ...)` | Move o foco para um campo. |
| `clean_all_fields(...)` | Limpa campos de texto da tela/aba. |

Os métodos acima normalmente aceitam `target_index=0`, `selected_tab=0` e `skip_error=False`. Consulte [[Guias/Primeira automação|Primeira automação]] para exemplos de abas e rótulos duplicados.

## Seleções múltiplas e resultado

| Método | Uso |
| --- | --- |
| `multiple_selection_field(field_name, ...)` | Abre a seleção múltipla de um campo. |
| `multiple_selection_paste_data(data, ...)` | Cola uma lista na janela de seleção múltipla. |
| `get_footer_message()` | Lê a mensagem exibida no rodapé SAP. |
| `get_text_at_side(field_name, side_index, ...)` | Lê o texto relativo a um rótulo. |
| `find_text_field(field_name, ...)` | Verifica se o texto existe na tela. |
| `save_file(file_name, path, option, type_of_file, ...)` | Exporta a lista atual pelo menu padrão do SAP. |

## Assinaturas e parâmetros comuns

`field_name` é o texto visível do rótulo ou do botão. `desired_text` é o conteúdo desejado. `target_index=0` escolhe a primeira ocorrência do mesmo rótulo; use `1`, `2` etc. para as próximas. `selected_tab=0` usa a primeira aba e também pode receber o nome da aba. `skip_error=False` mantém a falha explícita; só use `True` para ações opcionais.

Os métodos `*_at_side` recebem `side_index`: `0` é o controle associado diretamente, valores positivos avançam à direita e negativos à esquerda. Exemplos completos estão em [[Guias/Preencher telas|Preencher telas]].

## Controles especializados

| Método | Retorno |
| --- | --- |
| `get_grid()` | [[Referência/Componentes SAP#grid|Grid]] |
| `get_table()` | [[Referência/Componentes SAP#table|Table]] |
| `get_tree()` | [[Referência/Componentes SAP#tree|Tree]] |
| `get_node()` | [[Referência/Componentes SAP#node|Node]] |
| `get_label()` | [[Referência/Componentes SAP#label|Label]] |

As páginas de uso detalhado são [[Referência/Grid|Grid]], [[Referência/Table|Table]] e [[Referência/Tree, Node e Label|Tree, Node e Label]].
