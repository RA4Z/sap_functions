---
tags:
  - manutenção
  - contribuição
---

# Como documentar mudanças

Cada recurso novo deve atualizar a nota de referência mais próxima e, quando fizer sentido, incluir um exemplo de uso em um guia.

## Convenções deste vault

- Use títulos curtos em português e links internos no formato `[[Nome da nota]]`.
- Mantenha uma nota por assunto; prefira links a copiar a mesma explicação em vários arquivos.
- Adicione `tags` ao frontmatter para facilitar busca no Obsidian.
- Inclua exemplos executáveis e descreva pré-requisitos ou limitações conhecidas.
- Ao remover ou renomear uma API, marque a alternativa e a versão em que a mudança ocorreu.

## Checklist de pull request

- [ ] O novo método aparece na nota de referência adequada.
- [ ] Um exemplo foi adicionado ou atualizado, quando a API é pública.
- [ ] Links internos foram conferidos no Obsidian.
- [ ] O comportamento de compatibilidade/obsolescência está documentado.

O arquivo `doc-generator.py` ainda pode gerar uma listagem técnica no `README.md`; este vault é a documentação orientada a quem usa a biblioteca.
