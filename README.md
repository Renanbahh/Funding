# Funding — Capex Comparison (PLA-FOR-010)

Página web para comparar duas versões do arquivo de funding **PLA-FOR-010** (Aligned LatAm): o arquivo **anterior** (*Cons*, em rosa) e o **atual** (*PROD*, em azul).

## Como usar
1. Abra a página publicada (GitHub Pages) no **Chrome ou Edge**.
2. Na aba **Início**, selecione o arquivo anterior e o atual (os `.xlsx` da pasta sincronizada do SharePoint/OneDrive). Opcionalmente, cole os links do SharePoint para referência.
3. Clique em **Realizar comparação**. Em poucos segundos as abas são liberadas:
   Odata CapEx Control · Capex Comparison · Change Details · Comparison w/o FX · Building Budget.
4. Para guardar o resultado, use **Baixar este dashboard (.html)** e salve no SharePoint da equipe.

**Privacidade:** os arquivos são lidos **somente no navegador** de quem usa a página. Nada é enviado para a internet nem para o GitHub. A página publicada não contém nenhum dado da empresa.

## Publicar no GitHub Pages
Settings → Pages → *Build and deployment* → Source: **Deploy from a branch** → Branch: `main` / `(root)` → Save.
O endereço fica `https://renanbahh.github.io/Funding/`.

> Atenção: páginas do GitHub Pages são públicas na internet (mesmo com repositório privado, salvo no GitHub Enterprise com páginas privadas). Por isso **não** coloque no repositório dashboards baixados com dados (`.gitignore` já bloqueia os nomes padrão).

## Estrutura
| Caminho | O que é |
|---|---|
| `index.html` | A página (tudo em um arquivo, sem dados, sem dependências externas). |
| `ferramentas/capex-summary-compare/` | Skill do Claude e scripts Python opcionais (gerar o dashboard já preenchido). |
| `ferramentas/capex-summary-compare/assets/dashboard.html` | Fonte da página. Ao alterar, copie para `index.html` trocando `__DATA__` por `null`. |
