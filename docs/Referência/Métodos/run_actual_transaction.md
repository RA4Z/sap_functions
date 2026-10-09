---
tags: [método, sap, navegação]
---

# `run_actual_transaction`

Assinatura: `sap.run_actual_transaction(skip_error=False)`.

Executa a tela aberta: envia Enter e, se o título não mudar, envia F8. Retorna `None`. É apropriado após preencher a tela de seleção de uma transação.

```python
sap.run_actual_transaction()
```
