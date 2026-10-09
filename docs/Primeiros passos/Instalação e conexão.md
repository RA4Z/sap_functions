---
tags:
  - guia
  - instalação
---

# Instalação e conexão

## Pré-requisitos

- Windows com SAP GUI instalado e com **SAP GUI Scripting** habilitado.
- Python 3.9 ou superior.
- Uma sessão do SAP aberta e autenticada antes de executar o script.

## Instalação

Instale a versão publicada da biblioteca:

```powershell
pip install sap_functions
```

Para contribuir ou testar a cópia local do repositório:

```powershell
pip install -e .
```

## Criando uma conexão

```python
from sap_functions import SAP

sap = SAP()
```

`SAP()` se conecta à primeira sessão SAP disponível. Para usar outra janela/sessão, informe seu índice:

```python
sap = SAP(window=1)
```

Se não houver SAP GUI aberto, Scripting estiver desabilitado ou o usuário não estiver autenticado, a criação do objeto falhará. Trate a exceção no programa que chama sua automação quando for necessário informar o usuário de forma amigável.

Próximo passo: [[Guias/Primeira automação|crie sua primeira automação]].
