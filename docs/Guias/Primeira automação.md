---
tags:
  - guia
  - exemplo
---

# Primeira automação

O fluxo mais comum é: conectar, abrir a transação, preencher campos e executar.

```python
from sap_functions import SAP

sap = SAP()
sap.select_transaction("COOIS")
sap.write_text_field("Material", "MAT-001")
sap.run_actual_transaction()
```

## Encontrando campos repetidos

Quando a tela tiver mais de um rótulo igual, use `target_index`. O índice começa em `0`.

```python
sap.write_text_field("Centro", "1000", target_index=1)
```

## Abas

As funções que recebem `selected_tab` aceitam o índice da aba ou seu nome.

```python
sap.write_text_field("Data", "01.01.2026", selected_tab="Seleção")
```

## Operações frequentes

```python
sap.press_button("Executar")
sap.flag_field("Somente itens em aberto", True)
sap.choose_text_combo("Status", "Liberado")
sap.multiple_selection_field("Material")
```

Use `skip_error=True` somente para uma ação opcional. Para etapas essenciais, deixe a exceção aparecer: ela ajuda a identificar alteração na tela SAP ou um rótulo não encontrado.

Veja também [[Referência/Classe SAP|a referência completa da classe SAP]].
