---
tags: [método, sap, navegação]
---

# `select_transaction`

Assinatura: `sap.select_transaction(transaction)`.

Abre a transação indicada. A entrada não diferencia maiúsculas/minúsculas; a biblioteca converte o código para maiúsculas e verifica a transação ativa. Retorna `None`. Se a SAP não abrir a transação, lança exceção que inclui a mensagem de rodapé.

```python
sap.select_transaction("coois")
```
