---
tags:
  - guia
  - formulários
---

# Preencher telas

Os métodos da classe `SAP` encontram o controle pelo texto visível na tela. Isso torna os scripts menos dependentes de IDs técnicos do SAP GUI.

## Escrever em campos

```python
sap.write_text_field("Material", "MAT-001")
sap.write_text_field_until("Data de lançamento", "31.12.2026")
```

`write_text_field` escreve no primeiro campo relacionado ao rótulo. `write_text_field_until` escreve no segundo campo, normalmente o limite final de um intervalo.

## Combos, checkboxes e opções

```python
sap.choose_text_combo("Tipo de ordem", "Ordem de produção")
sap.flag_field("Incluir concluídas", True)
sap.option_field("Exibir lista")
```

Para desmarcar um checkbox, passe `False` a `flag_field`.

## Campo ao lado do rótulo

Em telas cuja disposição não é a padrão, use o deslocamento relativo `side_index`:

```python
sap.flag_field_at_side("Status", True, side_index=1)
sap.option_field_at_side("Processamento", side_index=-1)
texto = sap.get_text_at_side("Documento", side_index=1)
sap.set_focus("Material", side_index=1)
```

Um valor positivo procura controles à direita; um negativo, à esquerda. `target_index` decide qual ocorrência do rótulo usar quando há repetição e começa em `0`.

## Botões e validação

```python
if sap.find_text_field("Centro"):
    sap.press_button("Executar")

sap.clean_all_fields()
```

`press_button` procura o texto ou tooltip do botão. `clean_all_fields` apaga os campos de texto da tela ou da aba selecionada.

Todos os métodos de formulário aceitam `selected_tab`: passe o índice da aba ou o nome dela. Veja também [[Guias/Primeira automação|Primeira automação]].
