# Dashboard de Tarifas Aéreas — Porto Alegre

Dashboard estático preparado para **GitHub Pages** e alimentado diretamente por uma planilha Excel.

## Atualização mais simples
1. Atualize a planilha mantendo os nomes das abas e colunas do modelo.
2. Renomeie o arquivo para `dados.xlsx`.
3. No repositório do GitHub, substitua o arquivo `data/dados.xlsx`.
4. Faça o commit. O GitHub Pages passa a exibir os novos dados sem alteração no código.

## Testar antes de publicar
Abra o dashboard e use **Carregar planilha**. A leitura é feita no navegador e serve para conferir os dados antes de substituir o arquivo do GitHub.

## Abas lidas pelo dashboard
- `BASE_TARIFAS_ANAC`
- `BASE_GOOGLE_FLIGHTS`
- `INCENTIVOS_MEDIDAS`
- `GLOSSARIO_METODOLOGIA`

## Filtros
- Estado de origem
- Aeroporto de origem
- Companhia aérea
- Ano

## Período
O modelo já está preparado para jan/2025 a jun/2026 e aceita meses posteriores automaticamente.

## Publicar no GitHub Pages
1. Crie um repositório no GitHub.
2. Envie todo o conteúdo desta pasta para a raiz do repositório.
3. Vá em **Settings > Pages**.
4. Em *Build and deployment*, escolha **Deploy from a branch**.
5. Selecione a branch `main` e a pasta `/ (root)`.
6. Salve.

> Importante: ao abrir `index.html` diretamente pelo Explorador do Windows, o navegador pode bloquear a leitura automática do Excel. No GitHub Pages isso funciona normalmente. Para teste local, você também pode usar o botão “Carregar planilha”.


## Mercados emissores prioritários
SP • RJ • MG • SC • PR

## Atualização v3
- Filtro de aeroporto removido.
- Ano padrão: 2026.
- Novo gráfico mensal comparando 2025 × 2026, filtrável por origem e companhia.
- Base preenchida com os dados transcritos dos prints de RJ, SP, MG, SC e PR.

## Atualização v4
- Comparativo 2025 × 2026 limitado aos meses disponíveis nos dois anos.
- Enquanto 2026 tiver dados até junho, 2025 também aparece apenas de janeiro a junho.
- Variação percentual mensal exibida diretamente no gráfico comparativo.
- A direção da variação acompanha o ano selecionado: 2026 vs 2025 ou 2025 vs 2026.
- KPI "Variação anual" compara a média dos mesmos meses dos dois anos e inverte a comparação ao trocar o filtro de ano.

## Atualização v5
- Tooltips explicativos adicionados aos KPIs.
- Ao passar o mouse sobre Tarifa média, Menor tarifa, Maior tarifa e Variação anual, o painel explica o indicador e sua lógica de cálculo.
