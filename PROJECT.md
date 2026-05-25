# One Page Financeiro · 49 Educação

Sistema de reporte mensal consolidado: planilha xlsx preenchida pelos líderes de área + dashboard HTML que lê a planilha e gera um one page financeiro do mês.

## Contexto do negócio

A 49 Educação tem **3 áreas de receita**, cada uma com um líder responsável:

| Área | Líder | O que faz |
|------|-------|-----------|
| 🛠 Produto | Tiago Toigo | Entrega aulas/hot seats pra base contratada, cuida de tech, vende 49 Code |
| 🤝 B2B | Marcelo Câmara | Capta novos clientes (Sebrae, SATC, PUC etc), marketing e funil B2B |
| 📋 Licitações | Rafael Peck | Receita via editais públicos + prestação de contas de editais em execução |

**Financeiro** é gerido pela Luana (Coordenadora Financeira), que consolida o resultado consolidado mensal e entrega o one page pro CEO Leandro Piazza / Conselho.

## Princípio fundamental: Vendido vs Recebido

A planilha separa explicitamente:
- **Vendido** (preenchido pelos líderes): contratos fechados / receita gerada no mês
- **Recebido** (preenchido pela Luana): entrada efetiva em caixa no mês

Isso porque vender ≠ receber — alguns contratos só entram em caixa meses depois. O one page mostra o gap entre os dois.

## Arquivos do projeto

- `Planilha_Resultados_2026.xlsx` — planilha que os líderes preenchem
- `One_Page_Financeiro.html` — dashboard standalone (abre no navegador, lê a planilha via upload)

## Estrutura da planilha

6 abas, nessa ordem:

### 📋 Instruções
Manual de uso, responsáveis, prazos (dia 15 do mês seguinte), legenda de cores.

### 🛠 Produto (Tiago Toigo)
- **Entrega**: Aulas/Hot Seats, Marcos entregues, SLA de entrega (%)
- **Saúde da base**: Contratos ativos, NPS
- **Receita (vendida)**: Vendas 49 Code (qtd), Receita 49 Code, Receita Total da Área
- **Tech**: Uptime, Bugs reportados/resolvidos, Features lançadas

### 🤝 B2B (Marcelo Câmara)
- **Marketing/Funil**: Investimento ADS, Leads, MQLs, SQLs, CAC (auto)
- **Pipeline & Vendas**: Pipeline ativo, Novos leads, Propostas, Contratos fechados, Receita B2B vendida, Ticket médio (auto), Taxa de fechamento (auto), Ciclo médio

### 📋 Licitações (Rafael Peck)
- **Pipeline**: Editais monitorados, submetidos, ganhos, Taxa de êxito (auto)
- **Receita (vendida)**: Receita editais ganhos, Pipeline editais
- **Prestação de contas**: Contratos vigentes, Prestações entregues/pendentes/em atraso
- **Operacional**: TRs em aprovação, Follow-ups

### 💰 Financeiro (Luana)
- **Caixa & Recebimentos**: Saldo inicial, Recebimentos por área (B2B/Produto/Editais/Outras), Total (auto)
- **Vendido no Mês (auto)**: Puxa Vendido das 3 áreas via fórmula cross-sheet
- **Despesas**: Folha & RH, Marketing/ADS, Infra & Tech, Terceiros, Outras, Total (auto)
- **Resultado & Indicadores (auto)**: Resultado, Burn Rate, Runway, Margem
- **Contas a Receber + Inadimplência** (input)
- **Vendido vs Recebido (auto)**: Gap e % conversão venda→caixa

### 📤 Export
Aba consolidada com 61 indicadores em formato tabular. **É essa aba que o HTML lê.**

Formato: `chave | area | label | Jan ... Dez | Total`

Cada linha referencia via fórmula a célula de origem nas outras abas. Quando preenchem as abas, o Export atualiza sozinho.

## Convenções da planilha

**Cores de células:**
- Texto azul (#0000FF) + fundo creme (#FFFAF2) = campo de input
- Texto laranja itálico (#E87A3E) + fundo laranja claro (#FCE9DC) = calculado automaticamente
- Cabeçalhos verde escuro (#1F4D3F) com texto branco

**Formatos numéricos:**
- Currency: `R$ #,##0.00;(R$ #,##0.00);"-"`
- Integer: `#,##0;(#,##0);"-"`
- Percent: `0.0%;(0.0%);"-"`
- Decimal: `0.0;(0.0);"-"`

**Fonte:** Arial (toda a planilha)

**Freeze panes:**
- Abas operacionais: C6 (congela header e coluna do indicador)
- Export: D2 (congela header e 3 primeiras colunas)

**Fórmulas auto importantes:**
- B2B CAC = Investimento ADS ÷ Contratos Fechados
- B2B Ticket Médio = Receita Vendida ÷ Contratos Fechados
- B2B Taxa Fechamento = Contratos Fechados ÷ SQLs
- Licitações Taxa de Êxito = Editais Ganhos ÷ Submetidos
- Financeiro Recebimentos Totais = soma dos 4 recebimentos
- Financeiro Vendido (3 áreas) = referência cross-sheet: `='🤝 B2B'!C18`, `='🛠 Produto'!C18`, `='📋 Licitações'!C13`
- Financeiro Resultado = Recebimentos Totais − Despesas Totais
- Financeiro Burn = IF(Resultado<0, -Resultado, 0)
- Financeiro Runway = Saldo Inicial ÷ Burn
- Financeiro Margem = Resultado ÷ Recebimentos
- Gap Vendido−Recebido = Total Vendido − Total Recebido
- % Conversão Venda→Caixa = Recebido ÷ Vendido

**Importante:** Nomes de aba com emoji precisam de aspas simples nas fórmulas: `='🤝 B2B'!C18`

## Estrutura do HTML

`One_Page_Financeiro.html` — arquivo único standalone (HTML + CSS + JS inline). Dependências externas:
- Google Fonts: Inter + Fraunces
- SheetJS (xlsx.full.min.js) via CDN cdnjs

**Fluxo:**
1. Usuário arrasta `.xlsx` na zona de upload
2. SheetJS lê a aba `📤 Export`
3. Parse: cada linha vira `dataStore[chave] = { area, label, valores: [12 meses], total }`
4. Helper `get(chave, month)` retorna o valor
5. Render usa o mês selecionado no dropdown
6. Auto-seleciona o último mês com dados ao carregar

**Seções do dashboard:**
1. KPIs principais (4 cards): Recebimentos, Vendido, Resultado, Runway — com deltas vs mês anterior
2. 3 cards de área (Produto/B2B/Licitações) com receita vendida em destaque + 7-8 métricas operacionais
3. Tabela Recebimentos por categoria + tabela/bars de Despesas
4. Tabela Vendido vs Recebido (gap por área)
5. Strip YTD verde (acumulado até o mês selecionado)
6. Timeline de recebimentos (12 meses, clicável)

**Sistema visual (CSS vars):**
```css
--cream: #FFFAF2     /* fundo principal */
--cream-deep: #F5EFE2
--green: #1F4D3F     /* primária — cabeçalhos, texto destaque */
--green-soft: #2E6B58
--green-bg: #E8F0EC
--orange: #E87A3E    /* secundária — receitas, destaque numérico */
--orange-soft: #F4A574
--orange-bg: #FCE9DC
--black: #1A1A1A
--gray: #6B6B6B
--gray-light: #C8C2B2
--red: #C8493E       /* burn, gap negativo, runway baixo */
--red-bg: #FBE8E5
--border: #D4CFC1
```

**Tipografia:**
- Fraunces (serif): títulos, valores numéricos grandes (KPIs, YTD)
- Inter (sans): texto geral, labels, tabelas

**Estados condicionais:**
- Card de Resultado fica vermelho se < 0
- Card de Runway: vermelho se < 6 meses, laranja se < 12, verde se ≥ 12
- Delta verde (↑) se positivo vs mês anterior, vermelho (↓) se negativo

## Como rodar / iterar

**Adicionar um novo indicador:**
1. Adicionar linha na aba da área correspondente (Produto/B2B/Licitações/Financeiro)
2. Adicionar entrada em `EXPORT_ROWS` na aba 📤 Export
3. Adicionar `get('chave_nova', m)` no JS e renderizar onde quiser

**Mudar formato de número:**
- Na planilha: `cell.number_format = 'formato'`
- No HTML: usar `fmtCurrency`, `fmtCurrencyShort`, `fmtPct`, `fmtInt`, `fmtDecimal`

**Recalcular fórmulas após editar via openpyxl:**
```bash
python3 /mnt/skills/public/xlsx/scripts/recalc.py Planilha_Resultados_2026.xlsx
```
LibreOffice é necessário (já instalado no ambiente).

**Testar o HTML sem servidor:**
Abre direto no navegador (file://). Tudo é client-side, não precisa de backend.

## Decisões de design importantes

1. **Por que aba Export?** Centralizar a leitura em uma única aba evita o HTML ter que parsear 5 abas diferentes com formatos distintos. O Export é gerado automaticamente via fórmulas — quem preenche não toca nele.

2. **Por que receita é vendida nos líderes e recebida no Financeiro?** Como mencionado, vendido ≠ recebido. Os líderes têm visibilidade do que fecharam; só a Luana sabe o que entrou no banco. O one page reconcilia os dois.

3. **Por que dropdown de mês + timeline clicável?** Dois caminhos de navegação: dropdown formal pro CEO/conselho que quer "ver março", timeline visual pra navegação rápida exploratória.

4. **Por que YTD acumula até o mês selecionado e não YTD do ano todo?** Quando se está olhando o mês de março, o YTD relevante é jan+fev+mar, não o ano fechado.

5. **Prazo dia 15 do mês seguinte:** Janela razoável pros líderes consolidarem dados do mês anterior, antes do fechamento contábil mensal.

## Stack técnico

- **Planilha**: Python + openpyxl pra geração, LibreOffice pra recalcular fórmulas
- **HTML**: vanilla HTML/CSS/JS + SheetJS (xlsx parsing) + Google Fonts
- **Sem build, sem framework, sem backend** — tudo standalone

## Próximos passos possíveis

- Versão B2B mais detalhada com motivo de perda (já existia na planilha antiga)
- Adicionar gráfico de evolução mensal (Chart.js) no HTML
- Exportar one page como PDF via print CSS
- Versão web hosted (GitHub Pages) com persistência via window.storage
- Notificações automáticas (dia 14) lembrando os líderes via Telegram bot (integração com o bot que já existe)
