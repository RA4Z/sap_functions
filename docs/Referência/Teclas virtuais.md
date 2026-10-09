---
tags:
  - referência
  - vkey
---

# sendVKey e teclas virtuais

Controles retornados por `sap.session.findById(...)` aceitam uma tecla pelo nome. O editor sugere os valores válidos ao escrever a string.

```python
window = sap.session.findById("wnd[0]")
window.sendVKey("Enter")
window.sendVKey("F8")
window.sendVKey("Ctrl+S")
```

Os códigos numéricos do SAP continuam aceitos, por compatibilidade:

```python
window.sendVKey(8)  # F8
```

## Teclas disponíveis

| Grupo | Valores |
| --- | --- |
| Básicas | `Enter`, `F1` a `F10`, `F12`, `Ctrl+S` |
| Shift | `Shift+F1` a `Shift+F9`, `Shift+F11`, `Shift+F12`, `Shift+Ctrl+0` |
| Ctrl | `Ctrl+F1` a `Ctrl+F12` |
| Ctrl + Shift | `Ctrl+Shift+F1` a `Ctrl+Shift+F12` |

Também é possível usar `sap.session.activeWindow.sendVKey("F8")`.

> [!warning]
> Prefira os nomes às teclas numéricas: eles tornam o script legível e evitam depender de decorar o código SAP de cada atalho.
