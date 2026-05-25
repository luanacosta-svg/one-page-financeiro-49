# One Page Financeiro · 49 Educação

Dashboard standalone que lê a Planilha de Resultados (Google Sheets) e gera o one page financeiro mensal pra apresentar pro CEO/Conselho.

## Arquivos

- **`One_Page_Financeiro.html`** — Dashboard standalone (abre direto no navegador)
- **`Planilha_Resultados_2026.xlsx`** — Planilha que os líderes preenchem (já no Drive)
- **`build_planilha.py`** — Script que gera o xlsx aplicando o design system
- **`PROJECT.md`** — Contexto completo do projeto

## Como usar

1. Abre `One_Page_Financeiro.html` no navegador
2. Ele conecta automaticamente na [Planilha de Resultados](https://docs.google.com/spreadsheets/d/1MCHVSFd_q9Q5SkNu94aS2lgBzDcsWns4GPlyV3ce6Lk)
3. Mostra o one page atualizado

Botão **Atualizar** força nova leitura. Botão **Imprimir** gera PDF.

## Stack

- HTML/CSS/JS vanilla + Tailwind CDN + PapaParse (CSV) + Inter via system font
- Sem build, sem backend — single file

## Áreas

| Área | Líder | Aba na planilha |
|---|---|---|
| 🛠 Produto | Tiago Toigo | `🛠 Produto` |
| 🤝 B2B | Marcelo Câmara | `🤝 B2B` |
| 📋 Licitações | Rafael Peck | `📋 Licitações` |
| 💰 Financeiro | Luana Costa | `💰 Financeiro` |
| 📊 Consolidação | Automático | `📊 Dashboard (Não preencher)` |

A aba `📊 Dashboard (Não preencher)` é lida pelo HTML — não editar manualmente.
