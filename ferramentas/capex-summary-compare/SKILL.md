---
name: capex-summary-compare
description: Compara a aba "CapEx Summary" de duas versões do arquivo de funding PLA-FOR-010 (Aligned LatAm), a consolidada ("Cons", base antiga) e a "PROD" (atual), e gera um dashboard HTML com projetos novos, removidos, possíveis renomeações, valores alterados (Total Budget, EFC, Spend to Date, Remaining Capex, % Procured), totais por país e por bloco, e mudanças na curva mensal, inclusive em meses já realizados. Use sempre que o usuário pedir para comparar o Cons com o PROD, comparar o CapEx Summary, ver "o que mudou" entre versões do funding, listar projetos novos ou valores atualizados na pasta "01. Funding", ou gerar o "comparativo"/"dash de comparação", mesmo que não cite a aba pelo nome. Não use para gerar o CAPEX Control Dashboard nem para as conferências da aba CapEx Detail USD (isso é do skill capex-control-dashboard).
---

# Comparativo CapEx Summary (Cons × PROD)

Gera um dashboard HTML autocontido que mostra o que mudou na aba **CapEx Summary** entre a versão consolidada do PLA-FOR-010 (base) e a versão PROD (atual). O trabalho pesado fica num script; seu papel é localizar os arquivos, rodar, validar e resumir o resultado.

## Arquivos do skill
- `scripts/compare_capex_summary.py`: lê as duas versões, compara e grava o HTML e um JSON com os mesmos dados.
- `assets/template.html`: layout do dashboard (cores da marca Aligned). Os dados entram no lugar de `__DATA__`.

## Fluxo

1. **Localizar os arquivos.**
   - A pasta padrão é `C:\Users\renan.baptista\Aligned Energy Holdings LP\CAPEX Control - Documentos\11. Arquivos Claude - Renan\01. Funding` (OneDrive sincronizado do SharePoint). Se não estiver conectada, peça acesso com `request_cowork_directory` passando esse caminho.
   - Na pasta sempre há dois arquivos: um com **"Cons"** no nome (base, mais antiga) e um com **"PROD"** (atual). Identifique-os pelo nome, não pela data de modificação. O OneDrive costuma mostrar o Cons como modificado depois do PROD, e a data inverteria a comparação. Se houver mais de um de cada, o script usa a data do nome (dd-mm-aa).
   - Se o usuário mandar só um link do SharePoint, peça que sincronize a pasta. O conector entrega texto, e o script precisa do `.xlsx`.
   - Não altere os originais.

2. **Rodar o script em cópias** (os arquivos têm cerca de 40 MB; a leitura leva uns 20 segundos):
   ```bash
   mkdir -p /tmp/cs && cp "<pasta>/"*.xlsx /tmp/cs/
   python3 -I ${CLAUDE_SKILL_DIR}/scripts/compare_capex_summary.py --folder /tmp/cs \
       --template ${CLAUDE_SKILL_DIR}/assets/template.html \
       --out <pasta_saida>/Comparativo_CapEx_Summary_<dd-mm-aa>.html
   ```
   - Use a data do nome do arquivo PROD em `<dd-mm-aa>`.
   - Por padrão, salve em `outputs` e apresente com `present_files`. Grave na pasta do SharePoint só se o usuário pedir.
   - Para escolher arquivos específicos, use `--old <cons.xlsx> --new <prod.xlsx>` no lugar de `--folder`.
   - `--tol` define a menor diferença, em US$, que conta como alteração (padrão 1).
   - Os valores são lidos do cache do Excel. Se o arquivo não foi salvo depois de recalcular, os números podem estar defasados.

3. **Validar.**
   - O script imprime um resumo em JSON (novos, removidos, alterados, curva). Confira se os números fazem sentido.
   - Teste a sintaxe do JS do HTML gerado: extraia o `<script>` e rode `node -e "new Function(...)"`.
   - Faça uma captura de tela do HTML com `screenshot_file_preview` para conferir o layout.
   - Se o script falhar por estrutura (cabeçalho não encontrado, coluna ausente), a planilha mudou. Abra a aba, entenda a mudança e ajuste o script. Não contorne o erro preenchendo dados à mão.

4. **Entregar** o HTML com um resumo curto no chat (até 10 linhas):
   - os arquivos comparados e a data de corte de cada um;
   - a variação total de EFC, Total Budget e Spend to Date, e os países que mais mudaram;
   - projetos novos, removidos e possíveis renomeações, pelo nome;
   - os 3 a 5 projetos com maior variação de EFC;
   - os alertas: mudanças em meses já realizados e erros de fórmula (#REF!) na versão atual.
   - Descreva o que mudou, não por que mudou. A causa só o dono da planilha confirma.

## Regras fixas definidas pelo usuário
- **Filtros em cascata, em todas as abas:** cada caixa de filtro mostra só as opções compatíveis com o que foi escolhido nas caixas acima dela, de cima para baixo. Por exemplo, escolher Brazil em Country deixa Other, Type, Category e Project/Site só com o que existe no Brasil.
- **Projetos de ajuste ficam sempre fora, nas duas versões:** todo Phase/Project que começa com "Ajuste" (hoje: BG01 Ajuste Colombia, QR02 Ajuste Mexico, SP02 Ajuste Brasil e ST02 Ajuste Chile). Os dois scripts aplicam isso com `EXCLUDE_PROJECT`.

## Ajuste à tela (todas as páginas)
O dashboard é desenhado numa tela fixa de 1920 px de largura (`#app`) e escalado pela largura do monitor (`fit()`), de modo que mantém as mesmas proporções no notebook ou no monitor grande. A altura usa todo o espaço disponível pela variável `--appH`. Por isso, use `var(--appH)` no lugar de `vh`, e deixe o tooltip (`#tip`) fora do `#app`. Não reutilize nomes de classe que já são usados nas tabelas (ex.: `.pick` está em `tr.pick`).

## Versão publicada (página web, ex.: GitHub Pages)
O projeto roda como página web, sem ninguém executar nada localmente.
- A página publicada é o `index.html` na raiz do repositório https://github.com/Renanbahh/Funding. Ela é uma cópia de `assets/dashboard.html` com `__DATA__` trocado por `null`, ou seja, sem dados.
- Sem dados, só a aba Início funciona. As outras abas ficam desabilitadas até a primeira comparação.
- Os .xlsx escolhidos são lidos só no navegador de quem usa a página; nada vai para a internet.
- Ao alterar o template, atualize também o `index.html` (mesma troca) e faça o commit.
- O GitHub Pages é público. Nunca publique um dashboard com dados embutidos (gerado pelo `build_dashboard.py` ou baixado pela página).

## Página 0 — "Início" (abre primeiro)

É a página inicial, com cara de programa:
- **Campo rosa:** arquivo anterior (Cons). **Campo azul:** arquivo atual (PROD). Cada campo tem uma caixa para colar o link do SharePoint e um botão "Selecionar arquivo do computador…".
- **Botão "Realizar comparação":**
  - Com os dois arquivos selecionados, o próprio navegador lê os .xlsx, aplica as mesmas regras do `build_dashboard.py` (função `parseBook`, que é a mesma lógica em JavaScript) e abre o dashboard novo numa página nova. A leitura usa um leitor de .xlsx embutido (`readBook`: descompacta com o `DecompressionStream` do navegador e lê o XML das abas), sem bibliotecas externas e sem internet. Leva cerca de 2 segundos por arquivo de 40 MB e precisa de Chrome ou Edge atualizados.
  - Só com links, a página explica que o navegador não consegue baixar do SharePoint (por causa do login) e monta um pedido pronto para colar no Claude.
- **Confirmação:** "Comparação realizada entre os arquivos", com o nome, a data de corte, o último salvamento, o número de projetos e o link de cada arquivo, nas cores rosa e azul. Tem atalhos para as abas e um botão "Baixar este dashboard (.html)".
- **Quando o Claude gera o dashboard:** passe os links com `--prev-url` e `--cur-url` para que eles apareçam na confirmação:
  ```bash
  python3 -I ${CLAUDE_SKILL_DIR}/scripts/build_dashboard.py --folder /tmp/cs --template ${CLAUDE_SKILL_DIR}/assets/dashboard.html \
      --out <saida>/Odata_CapEx_Control.html --prev-url "<link anterior>" --cur-url "<link atual>"
  ```
- **Onde ficam os dados:** num `<script id="dash-data" type="application/json">`. O script lê esse bloco na carga e guarda uma cópia da página em `SRC`, para remontar o dashboard com outros arquivos.
- **Ao mudar regras de leitura:** mude no Python e no JavaScript (`parseBook`). O leitor embutido e o `parseBook` foram conferidos com os dois arquivos: mesmos projetos e valores do Python (zero diferenças), mesma data de corte, mesmo último salvamento e mesmas renomeações.

## Página 1 — "Odata CapEx Control" (só arquivo PROD)

Gerada por `scripts/build_dashboard.py` com `assets/dashboard.html`:
```bash
python3 -I ${CLAUDE_SKILL_DIR}/scripts/build_dashboard.py --folder /tmp/cs \
    --template ${CLAUDE_SKILL_DIR}/assets/dashboard.html --out <pasta_saida>/Odata_CapEx_Control.html
```
Regras definidas com o usuário (os números batem com o Power BI de referência):
- **Cabeçalho:** à esquerda, sempre o nome do arquivo PROD lido, sem a extensão (vem do arquivo, nunca digitado). À direita, "Last Update" com a data e a hora do último salvamento gravadas pelo Excel (`docProps/core.xml`), no horário de Brasília, no formato dd/mm/aa hh:mm. Não use a data do sistema de arquivos, que muda ao copiar ou sincronizar.
- **Exclusões:** categorias "Fitout Revenue" e "Utility Revenue", Type "Reimbursement" e os projetos "Ajuste…".
- **Type exibido:** categoria OFE vira "Spec - OFE"; categoria Fitout vira "Fitout"; nos demais, o Type da planilha. "Non-Spec" e "Non-spec" ficam separados, como no Power BI.
- **Spent e Remaining:** Spent são os meses até a data de corte; Remaining são os meses seguintes. O rótulo do corte (ex.: "Set-26") aparece na faixa "Total in the period".
- **Filtros:** Country, Other, Type, Category e Project (o filtro Commit? foi removido a pedido). Type e Category mostram todos os itens, sem barra de rolagem. A lista de projetos se ajusta aos outros filtros.
- **Período:** datas com dia, mês e ano (padrão: 01/01 a 31/12 do ano de corte). Cada mês conta pela data do último dia dele.
- **Filtro cruzado por clique, como no Power BI:** o gráfico mensal, a tabela e as barras por Type e a tabela de projetos são clicáveis. Cada clique filtra os outros visuais; o visual clicado continua mostrando tudo, com o item escolhido destacado e os demais apagados.
  - Um mês clicado filtra todos os visuais, menos o "Annual total".
  - Um Type ou um projeto clicado filtra todos os visuais, inclusive o "Annual total".
  - Ctrl+clique seleciona vários itens. Clicar de novo no mesmo item, o botão azul ao lado das datas ou Esc limpam a seleção.
- **Annual total:** de "Until 2022" até o ano de corte + 4, com Spent e Remaining separados pela data de corte.

## Página 2 — "Capex Comparison" (Cons = anterior × PROD = atual)

Fica no mesmo HTML, numa aba na barra inferior. O script lê os dois arquivos (`--folder`, ou `--prod` e `--cons`).
- **Cabeçalho:** os dois nomes de arquivo, o anterior em rosa e o atual em azul. O título é "Capex Comparison".
- **Filtros:** cada arquivo tem os seus, independentes: Country, Type, Capex categorization e Project. Os do anterior (rosa) ficam à esquerda e os do atual (azul) à direita.
- **Type desta página:** é o Type da planilha, mas a categoria Fitout vira "Fitout" e "Non-Spec"/"Non-spec" são unificados como "Non-spec". Aqui OFE fica dentro de Spec, ao contrário da página 1.
- **Gráficos:** um por arquivo, com o mês de corte do arquivo, o total do período e o número de projetos. O cartão "Incurred" soma os meses até o corte; o "Forecast" soma os meses seguintes.
- **Period Comparison:** quatro tabelas (Project, Type, Category e Country) com Capex Previous, Capex Current, uma seta e Capex Variation (valor absoluto). A seta é ▲ verde quando o valor sobe, ▼ vermelha quando cai e "-" quando não muda. As tabelas começam ordenadas pela maior variação e podem ser reordenadas clicando no cabeçalho.
- **Filtro cruzado:** clicar num mês (em qualquer dos dois gráficos) ou numa linha de tabela filtra os outros visuais dos dois arquivos. Ctrl+clique seleciona vários itens; o botão ao lado das datas ou Esc limpam a seleção.

## Página 3 — "Change Details" (mesmo visual escuro)

**Foco, a pedido do usuário:** Budget, EFC (sem FX), Contingency e Saving, além dos projetos novos, removidos e possivelmente renomeados. Esta aba não compara Spend to Date, Remaining Capex nem a curva mensal.
- **Dados:** as mesmas fontes da página 4, com as mesmas exclusões (Revenue, Reimbursement e Ajuste), lidas de `ALL.prev` e `ALL.cur` por projeto.
  - Budget = coluna "Total Budget".
  - EFC = coluna "C".
  - Contingency = CapEx Detail USD.
  - Saving = Budget − EFC − Contingency.
  - Os totais são a soma dos projetos filtrados.
- **Possíveis renomeações:** vêm de `cmp.renames`, calculado em `compare()`: um projeto removido e um novo no mesmo site e bloco.
- **Filtros à esquerda, iguais aos da página 1:** Country, Other, Type, Category e Project, em cascata e com zoom ao passar o mouse.
- **Cartões:** Budget, EFC, Contingency e Saving (valor atual e Δ). Também mostra as contagens de New projects, Removed projects, Possible renames e Relevant changes (|Δ| ≥ US$ 0,5 M ou ≥ 5% do Budget).
- **Seções:**
  - By Country: anterior, atual e Δ das 4 métricas.
  - Top 10 variações, com seletor de métrica.
  - Novos, removidos e renomeados.
  - Sites com mudanças.
  - Projetos com valores alterados: Δ das 4 métricas e mudanças de Country, Other, Type ou Category. Uma faixa vermelha marca as mudanças relevantes; clique numa linha para ver anterior × atual.

## Página 4 — "Comparison w/o FX" (Budget, EFC, Contingency, Saving)

- **Cabeçalho:** os dois nomes de arquivo. O título é "Capex Comparison w/o FX Variation", com o mês de corte do anterior (rosa) e do atual (azul).
- **Fontes por projeto (Site + Phase/Project), nos dois arquivos:**
  - Budget = coluna "Total Budget" da CapEx Summary.
  - EFC = coluna "C" da CapEx Summary, que é o EFC sem FX Variation. Ela bateu com o Power BI. Não use a coluna "EFC", que inclui FX.
  - Contingency = soma do "Remaining" das linhas de categoria "Contingency/Escalation" da aba **CapEx Detail USD**.
  - Saving = Budget − EFC − Contingency.
- **Exclusões:** as mesmas das páginas 1 e 2 (Revenue, Reimbursement, Ajuste). Além disso, o Type "Expansion" começa desmarcado, como no Power BI. Com isso, os totais de Budget e EFC batem com o print.
- **Layout:** três tabelas (Budget, EFC e Contingency + Saving) com Previous, Current, seta e Variation, ordenadas pela maior variação. Abaixo ficam os filtros (Country, Other, Type, Capex categorization), que valem para os dois arquivos e crescem de baixo para cima com o mouse por cima. A tabela-resumo do canto inferior direito foi retirada a pedido do usuário.
- **Interatividade:** clicar num projeto em qualquer tabela filtra as outras.
- **Ponto em aberto:** em alguns projetos a Contingency do Power BI não tem linha "Contingency/Escalation" correspondente no arquivo. Confirme a fonte com o usuário antes de mudar a regra.

## Página 5 — "Building Budget" (projetos de categoria Building)

Usa as mesmas fontes da página 4 (Budget, EFC "C", Contingency e Saving).
- **Projetos:** só os de categoria "Building" com algum valor. Um botão alterna entre o arquivo atual (padrão) e o anterior.
- **Duas barras por projeto:**
  - "Approved" mostra o Budget (azul-ciano #00A9CE).
  - "Current" mostra EFC (cinza) + Contingency (verde) + Saving (amarelo), e somadas fecham no Budget.
  - Quando o Saving é negativo, a parte que passa do Budget aparece listrada em vermelho como "Estouro", com uma marca branca na posição do Budget, sem borda vermelha (a borda indica mudança entre arquivos).
- **Cabeçalho de cada projeto:** mostra o que mudou em relação ao outro arquivo (Δ Budget, Δ EFC, Δ Cont., Δ Saving), "novo" ou "sem alteração". Não mostra mais "% usado" nem Saving/Estouro.
- **Contorno vermelho** = mudança relevante em relação ao outro arquivo: |Δ| ≥ US$ 500 mil ou ≥ 5% do Budget em qualquer métrica, ou projeto novo. O limite fica em `REL5` no template.
- **Cartões:** Projetos, Budget, EFC (% do Budget), Contingency (% do Budget), Saving, Mudanças relevantes e Projetos com estouro.
- **Ordem e layout:** por padrão, os projetos aparecem em ordem de nome (BG01, BG01B, … SP02 Building 1, Building 2…). A lista enche a coluna da esquerda até a altura da tela e continua na da direita; se não couber nas duas, divide ao meio e rola.
- **Controles:**
  - Escala por projeto (padrão, cada barra ocupa a largura toda) ou escala comum (todas as barras na mesma escala).
  - Busca por projeto.
  - Ordenação por maior Budget, menor Saving, maior Contingency, maior % do Budget consumido ou nome.
- **Filtros:** Country, Other, Type e Site, com o efeito de crescer com o mouse por cima.

## Como a comparação funciona (para explicar resultados)

- **Estrutura da aba.** O cabeçalho é a linha com `Site | Phase/Project` (hoje, a linha 6). As colunas são achadas pelo nome, então inserir colunas não quebra o script. A área de projetos termina na primeira linha que tem só o texto da coluna A, sem projeto (hoje, "Expansions", perto da linha 479). Tudo abaixo disso são tabelas de apoio e fica fora da comparação.
- **Chave do projeto:** `Site | Phase/Project`. Nomes repetidos no mesmo site recebem `#2`, `#3` (ex.: ST02 "Energy investment"). Cada projeto pertence a um **bloco**, que é o nome da linha "Total" logo abaixo dele (ex.: "SP01 CapEx", "SP01 Fitout").
- **Possível renomeação:** um projeto removido e um novo no mesmo site e bloco. É uma sugestão a confirmar, não uma conclusão.
- **Totais por país e por bloco** vêm das linhas "Total" da própria planilha, não da soma das linhas de projeto. As fórmulas de Total excluem algumas linhas (por exemplo, a "Expansion" de vários sites), e somar tudo inflaria os valores em bilhões. Por isso uma alteração num projeto pode não aparecer no total do bloco.
- **Métricas comparadas:** Total Budget, EFC, EFC Non-Spec, EFC Spec, Spend to Date, Remaining Capex e % Procured (em pontos percentuais).
- **Atributos comparados:** Capex categorization, Type, Fitout, Support Debt Covenants, Other (status), Commit?, BBNB Capex e Country.
- **Curva mensal:** é a primeira sequência contínua de colunas com data no cabeçalho. Um mês "realizado" é aquele menor ou igual ao mês de corte (a data acima de "Remaining Capex"). Mudança em mês realizado merece atenção, porque o realizado normalmente não deveria mudar entre versões. Se as datas de corte das duas versões forem diferentes, o dashboard avisa que parte da variação de Spend to Date pode ser só o avanço do mês.

## O que o dashboard mostra
- **Resumo:** cartões de EFC, Total Budget, Spend to Date e Remaining Capex (atual e Δ), contagens, tabela por país, blocos com total alterado e gráfico das 10 maiores variações de EFC.
- **Novos e removidos:** possíveis renomeações, projetos novos e projetos removidos, com os valores.
- **Valores alterados:** tabela com busca, filtro por país e ordenação por métrica. Clique numa linha para ver base × atual de cada campo.
- **Curva mensal:** gráfico do total mês a mês (base × atual), meses com diferença, mudanças em meses realizados e projetos com curva remanejada.
