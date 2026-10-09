---
tags:
  - guia
  - diagnóstico
---

# Diagnóstico de erros

## Leia a mensagem do rodapé

Quando uma ação do SAP não funcionar, o rodapé normalmente explica o motivo:

```python
try:
    sap.run_actual_transaction()
except Exception:
    print(sap.get_footer_message())
    raise
```

## Verifique o estado antes de agir

```python
if sap.find_text_field("Material"):
    sap.write_text_field("Material", "MAT-001")
else:
    raise RuntimeError("O campo Material não está disponível nesta tela.")
```

## `skip_error`

Os métodos que possuem `skip_error=False` lançam exceção por padrão. Passe `True` apenas se a ação puder legitimamente não existir, como fechar uma janela opcional.

```python
sap.press_button("Mensagem opcional", skip_error=True)
```

Não use `skip_error=True` em todas as chamadas: isso pode ocultar uma mudança de tela, uma falha de autorização ou um rótulo digitado incorretamente.

## Problemas na conexão

Antes de criar `SAP()`, confirme que o SAP GUI está aberto, que há uma sessão autenticada e que SAP GUI Scripting está habilitado no cliente e no servidor SAP. Veja [[Primeiros passos/Instalação e conexão|Instalação e conexão]].
