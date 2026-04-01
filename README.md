# Manipulador de Hierarquia

Ferramenta web **100% no navegador** para importar hierarquias (estilo planilha), validar relacionamentos pai–filho, navegar em árvore por níveis e visualizar um **relatório tabular** com colunas de ano/mês — com exportação para **Excel** e **PDF**, modo **página inteira** para foco no relatório, **tema claro/escuro** e interface em **português** ou **inglês**.

> **English (short):** Single-page app to paste hierarchies from Excel, expand/collapse tree branches, validate parents, export hierarchy to `.xlsx` and a visual snapshot of the report to `.pdf`, with light/dark theme and PT/EN UI.

![Manipulador de Hierarquia — importação, ferramentas e relatório (tema escuro, hierarquia de exemplo)](docs/screenshots/tela-principal.png)

---

## Sumário

- [Funcionalidades](#funcionalidades)
- [Como executar](#como-executar)
- [Fluxo de uso rápido](#fluxo-de-uso-rápido)
- [Formato da grade (ID, DESCRICAO, PAI)](#formato-da-grade-id-descrição-pai)
- [Relatório na tabela](#relatório-na-tabela)
- [Ferramentas da barra lateral](#ferramentas-da-barra-lateral)
- [Barra acima do relatório](#barra-acima-do-relatório)
- [Exportação Excel](#exportação-excel)
- [Exportação PDF](#exportação-pdf)
- [Atalhos de teclado](#atalhos-de-teclado)
- [Persistência (localStorage)](#persistência-localstorage)
- [Estrutura do repositório](#estrutura-do-repositório)
- [Tecnologias e dependências](#tecnologias-e-dependências)
- [Limitações conhecidas](#limitações-conhecidas)
- [Licença](#licença)

---

## Funcionalidades

### Importação e edição

- **Grade estilo Excel** com colunas **ID**, **DESCRICAO** e **PAI**.
- **Colar direto do Excel** (Ctrl+V) na grade.
- Botões **Aplicar hierarquia** (monta a árvore e o relatório) e **Limpar**.
- **Preencher exemplo** com uma hierarquia mínima para testes.
- **Desfazer** alterações na grade com **Ctrl+Z** (Windows/Linux) ou **Cmd+Z** (macOS), mesmo sem foco na célula.
- **Seleção retangular** na grade (arrastar), **Delete/Backspace** para limpar bloco selecionado (quando o foco não está digitando numa célula).
- **Ctrl+C / Ctrl+X** na grade com cópia/recorte para a área de transferência (com suporte a bloco selecionado).

### Validação e busca

- **Validar hierarquia**: abre modal com resumo (total, válidas, com erro) e lista de problemas (ex.: pai inexistente).
- **Buscar ID** e **Buscar descrição**: filtram linhas do relatório mantendo contexto dos níveis superiores visíveis.
- **Localizar e substituir** em massa em IDs e/ou descrições.
- Indicador de **linhas válidas | com erro** na grade.

### Visualização do relatório

- Tabela com **duas linhas de cabeçalho** (mês por extenso e ano/mês) e primeira coluna com **árvore expansível** (setas para abrir/fechar ramos).
- **Expandir tudo** / **Recolher tudo** na barra imediatamente acima da tabela.
- **Relatório em página inteira**: ocupa a área da aba (não é tela cheia F11 do navegador); esconde cabeçalho, importação e toolbar superior; **Esc** volta ao modo normal (respeitando fechamento de modais primeiro).
- Contador de **itens** na árvore.

### Aparência e idioma

- **Tema claro / escuro** (ícone lua/sol), preferência salva no navegador.
- **Português / inglês** (ícone de idioma), inclusive cabeçalhos do relatório e textos de botões.

### Exportação

- **Exportar Excel** (`.xlsx`) da hierarquia atual (grade ou dados aplicados).
- **Baixar PDF do relatório**: captura visual da área da tabela (html2canvas + jsPDF), **A4 paisagem**, várias páginas se necessário; **spinner animado** ao lado do botão enquanto gera; nomes de arquivo alinhados ao Excel (veja abaixo).

---

## Como executar

1. Clone ou baixe este repositório.
2. Abra o arquivo **`Manipulador Hierarquia/index.html`** no navegador.

**Recomendado:** servir a pasta por um servidor HTTP local (evita restrições de alguns navegadores com `file://` e garante carregamento estável de scripts):

```bash
# Na pasta do repositório (exemplo com Python 3)
cd "Manipulador Hierarquia"
python -m http.server 8080
```

Acesse `http://localhost:8080` no navegador.

> A biblioteca **SheetJS (xlsx)** ainda é carregada por **CDN**; **html2canvas** e **jsPDF** estão em **`vendor/`** para funcionar offline após o primeiro uso local. **Font Awesome** vem do CDN (ícones).

---

## Fluxo de uso rápido

1. Cole na grade os dados (ID, DESCRICAO, PAI) ou clique em **Preencher exemplo**.
2. Clique em **Aplicar hierarquia**.
3. Use as setas na primeira coluna do relatório para expandir/recolher níveis.
4. (Opcional) **Validar hierarquia**, **Exportar Excel** ou **Baixar PDF do relatório**.
5. (Opcional) **Relatório em página inteira** para apresentar só a tabela; **Esc** para sair.

---

## Formato da grade (ID, descrição, pai)

| Coluna     | Significado |
|-----------|-------------|
| **ID**    | Identificador único do nó. |
| **DESCRICAO** | Texto exibido na árvore. |
| **PAI**   | ID do pai; vazio ou valor sentinela de raiz conforme seus dados (o exemplo usa hierarquia com raiz explícita). |

Erros comuns detectados na validação: **PAI** referenciando um ID que não existe na lista.

---

## Relatório na tabela

- Colunas de valores são geradas para os **12 meses do ano corrente** (formato ano+mês no cabeçalho).
- A coluna indicadores mostra a **hierarquia indentada** com controles de expandir/recolher.
- Filtros de busca reduzem as linhas visíveis sem perder a cadeia de ancestrais necessária para contexto.

---

## Ferramentas da barra lateral

- **Preencher exemplo** · **Exportar Excel** · **Validar hierarquia** · **Localizar e substituir**
- Campos **Buscar ID** e **Buscar descrição**
- Chip de saúde da grade · **Idioma** · **Tema** · dica de uso das setas · contagem de itens

---

## Barra acima do relatório

- **Expandir tudo** / **Recolher tudo**
- **Baixar PDF do relatório** (com indicador de carregamento animado durante a geração)
- **Relatório em página inteira** / **Voltar à visualização normal** (mesmo botão alterna o modo)

---

## Exportação Excel

- Gera planilha com colunas ID, DESCRICAO, PAI.
- Se a grade tiver linhas preenchidas, usa a grade; senão usa os itens da hierarquia aplicada.
- Nome do arquivo: **`<DescriçãoPrimeiraLinhaSanitizada>_Hierarquia.xlsx`**, com a descrição da **primeira linha de dados** normalizada (sem acentos, caracteres inválidos para arquivo, espaços viram `_`). Se vazio, usa `hierarquia`.

---

## Exportação PDF

- Captura a **área rolável do relatório** (tabela visível na região do `.table-area`), não a grade de importação.
- PDF em **A4 paisagem**, imagem dividida em **várias páginas** se a altura for grande.
- **Nome do arquivo** segue a **mesma base** do Excel: **`<mesma_base>_Hierarquia.pdf`**.
- Requer linhas no corpo do relatório; bibliotecas em `vendor/` (**html2canvas**, **jspdf**).
- Em hierarquias muito grandes a geração pode levar vários segundos; o **spinner** permanece visível até concluir.

---

## Atalhos de teclado

| Atalho | Ação |
|--------|------|
| **Ctrl+Z** / **Cmd+Z** | Desfazer na grade |
| **Ctrl+C** (grade) | Copiar célula ou seleção |
| **Ctrl+X** (grade) | Recortar célula ou seleção |
| **Delete** / **Backspace** | Limpar células selecionadas (com condições; ver código) |
| **Esc** | Fecha modais abertos; em **modo página inteira** do relatório, volta ao layout normal (se o foco não estiver em campo de edição) |

**Colar na grade:** use **Ctrl+V** com foco na área da grade (comportamento de planilha).

---

## Persistência (localStorage)

Chaves aproximadas (nomes podem ser conferidos em `script.js`):

- Tema claro/escuro
- Idioma (pt/en)

---

## Estrutura do repositório

```
Modelador de Hierarquia/
├── README.md                      ← este arquivo
├── docs/
│   └── screenshots/               ← coloque aqui tela-principal.png e outras imagens
├── Manipulador Hierarquia/
│   ├── index.html                 ← página principal
│   ├── script.js                  ← lógica da aplicação
│   ├── styles.css                 ← estilos e tema escuro
│   └── vendor/
│       ├── html2canvas.min.js     ← captura para PDF (local)
│       └── jspdf.umd.min.js       ← geração de PDF (local)
```

---

## Tecnologias e dependências

| Uso | Origem |
|-----|--------|
| Lógica e DOM | JavaScript **vanilla** (sem framework) |
| Planilhas `.xlsx` | [SheetJS](https://sheetjs.com/) (CDN) |
| Captura HTML → imagem | [html2canvas](https://html2canvas.hertzen.com/) (arquivo em `vendor/`) |
| Montagem do PDF | [jsPDF](https://github.com/parallax/jsPDF) (arquivo em `vendor/`) |
| Ícones | [Font Awesome 6](https://fontawesome.com/) (CDN) |

---

## Limitações conhecidas

- O **PDF** é uma **imagem** do relatório (não é texto selecionável).
- **xlsx** depende de rede na primeira carga se usar apenas o CDN.
- Hierarquias **muito grandes** aumentam tempo de renderização e de **geração do PDF**.
- **Font Awesome** via CDN: sem internet, ícones podem não aparecer (a aplicação continua utilizável).

---

## Licença

Defina aqui a licença do projeto (ex.: MIT). Até lá, todos os direitos reservados pelo autor, salvo se o repositório já tiver um arquivo `LICENSE`.

---

## Créditos

Projeto **Manipulador de Hierarquia** — modelagem e visualização de hierarquias para fluxos de trabalho e aprovação com gestores (Excel + PDF).

Se este README te ajudou, uma estrela no repositório ajuda a divulgar.
