---
tags:
  - referência
  - avançado
  - sap-gui-scripting
---

# Acesso direto ao SAP GUI

Quando uma tela ainda não é coberta por um método de alto nível, use a sessão SAP GUI exposta em `sap.session`. Isso é uma API avançada: os IDs podem variar entre transações, versões do SAP GUI e personalizações.

## Encontrar um controle por ID

```python
campo = sap.session.findById("wnd[0]/usr/ctxtS_MATNR-LOW")
campo.Text = "MAT-001"

sap.session.findById("wnd[0]/tbar[1]/btn[8]").press()
```

O objeto retornado encaminha atributos e métodos ao controle SAP GUI original. Assim, atributos como `Text`, `text`, `selected` e métodos como `press()`, `Select()`, `setFocus()` e `change()` continuam disponíveis conforme o tipo do controle.

## Janela ativa e sessões

```python
window = sap.session.activeWindow
print(window.text)
window.sendVKey("F8")

sap.session.startTransaction("COOIS")
```

Para abrir outra sessão SAP GUI, use `sap.session.CreateSession()`. `sap.session.info` expõe informações da sessão, como `transaction`, `user`, `systemName`, `client` e `language`.

## Teclas virtuais

Use nomes legíveis com `sendVKey` tanto na janela ativa quanto em um controle encontrado por ID:

```python
sap.session.findById("wnd[0]").sendVKey("Enter")
sap.session.activeWindow.sendVKey("Ctrl+S")
```

Consulte [[Referência/Teclas virtuais|sendVKey e teclas virtuais]] para todos os nomes aceitos.

> [!warning]
> Grave scripts pelo SAP GUI Scripting ou inspecione a tela para obter IDs confiáveis. Antes de adicionar um atalho direto ao projeto, prefira verificar se [[Referência/Classe SAP|a classe SAP]] já resolve a ação pelo texto visível.
