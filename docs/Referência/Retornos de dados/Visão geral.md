---
tags: [referência, retornos, ia]
---

# Retornos de dados: visão geral

Esta nota define os formatos retornados pela biblioteca. Use-a como contexto para IAs: os nomes das chaves e o comportamento abaixo são parte da interface de uso.

| Origem | Método | Formato |
| --- | --- | --- |
| Grid | `get_content()` | `{"header": list[str], "content": list[list[Any]]}` |
| Grid | `get_columns(...)` | `list[Any]` (uma coluna) ou dicionário de tabela (várias) |
| Table | `get_content()` | `{"header": list[str], "content": list[list[Any]]}` |
| Table | `get_columns(...)` | `list[Any]` (uma coluna) ou dicionário de tabela (várias) |
| Tree | `get_content()` | `{"header": list[str], "content": list[list[str]]}` |
| Tree | `get_hierarchy_content()` | tabela com `Row ID` e `Parent ID` no início |
| Node | `get_content()` | `list[str]` |
| Label | `get_content()` | `{"header": list[str], "content": list[list[str]]}` |
| SAP | `find_text_field()` | `bool` |
| SAP | `get_text_at_side()`, `get_footer_message()` | `str` |

As listas de `content` mantêm a mesma ordem das colunas de `header`. Valores ausentes podem aparecer como `None` ao ler Grid/Table. As operações que alteram a tela retornam `None`; em caso de falha, normalmente lançam `Exception`.

Veja exemplos completos: [[Referência/Retornos de dados/Grid]], [[Referência/Retornos de dados/Table]], [[Referência/Retornos de dados/Tree]], [[Referência/Retornos de dados/Node e Label]] e [[Referência/Retornos de dados/Valores simples]].
