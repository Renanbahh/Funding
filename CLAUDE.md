# CLAUDE.md — instruções para o Claude Code neste repositório

## O que é
Página estática (GitHub Pages) que compara duas versões do arquivo de funding PLA-FOR-010 da Aligned LatAm: o arquivo anterior (*Cons*, cor rosa) e o atual (*PROD*, cor azul). Tudo roda no navegador. Não há servidor.

## Estrutura
- `index.html`: a página publicada, em um único arquivo (CSS, HTML e JS). É gerada a partir de `ferramentas/capex-summary-compare/assets/dashboard.html` trocando `__DATA__` por `null`.
- `ferramentas/capex-summary-compare/SKILL.md`: regras de negócio completas (exclusões, fontes de Budget, EFC, Contingency e Saving, regras de cada aba). Leia antes de mudar qualquer cálculo.
- `ferramentas/capex-summary-compare/scripts/build_dashboard.py`: as mesmas regras em Python, para gerar um dashboard já preenchido.

## Regras
1. **Edite sempre o template** (`assets/dashboard.html`) e depois regenere o `index.html`:
   `python -c "s=open('ferramentas/capex-summary-compare/assets/dashboard.html',encoding='utf-8').read(); open('index.html','w',encoding='utf-8').write(s.replace('__DATA__','null'))"`
2. **Nunca coloque dados da empresa no repositório.** Isso inclui `.xlsx` e dashboards gerados ou baixados com dados embutidos. O GitHub Pages é público.
3. **Regras de leitura existem em dois lugares:** `parseBook` (JS, no template) e `read()` (Python). Altere os dois juntos.
4. **Valide antes do commit:** a sintaxe do JS (`node -e "new Function(...)"` no último `<script>`) e se `index.html` contém `>null</script>` no bloco `dash-data`.
5. **Toda mudança vai ao ar.** O push em `main` publica no GitHub Pages. Para mudanças grandes, use uma branch e abra um pull request.
6. **Textos da interface em português.** Use as cores da marca: rosa `#F07AC0` para o anterior, azul `#6FB0FF` para o atual e primary blue `#07497E`.
