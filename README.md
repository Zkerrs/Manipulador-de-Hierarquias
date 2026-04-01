<div align="center">

# Manipulador de Hierarquia

**Ferramenta web para hierarquias no navegador** — planilha, validação, árvore, relatório, Excel e PDF.

![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=flat-square&logo=html5&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=flat-square&logo=javascript&logoColor=111)
![Licença](https://img.shields.io/badge/licença-MIT-22c55e?style=flat-square)

</div>

Ferramenta web **100% no navegador** para importar hierarquias (estilo planilha), validar relacionamentos pai–filho, navegar em árvore por níveis e visualizar um **relatório tabular** com colunas de ano/mês — com exportação para **Excel** e **PDF**, modo **página inteira** para foco no relatório, **tema claro/escuro** e interface em **português** ou **inglês**.

> **English (short):** Single-page app to paste hierarchies from Excel, expand/collapse tree branches, validate parents, export hierarchy to `.xlsx` and a visual snapshot of the report to `.pdf`, with light/dark theme and PT/EN UI.

<br>

<p align="center">
  <img src="tela-principal.png" alt="Manipulador de Hierarquia — importação, ferramentas e relatório (tema escuro, hierarquia de exemplo)" width="92%" />
</p>

<p align="center"><sub>Captura: tema escuro com hierarquia de exemplo aplicada.</sub></p>

<br>

---

## Sumário

| | |
|:---|:---|
| [Funcionalidades](#funcionalidades) | Importação, validação, relatório, tema, exportação |
| [Fluxo de uso rápido](#fluxo-de-uso-rápido) | Passos essenciais |
| [Formato da grade](#formato-da-grade-id-descrição-pai) | Colunas ID, DESCRICAO, PAI |
| [Relatório na tabela](#relatório-na-tabela) | Meses e árvore |
| [Licença](#licença) | MIT e autoria |
| [Créditos](#créditos) | Sobre o projeto |

---

## Funcionalidades

### Importação e edição

- **Grade estilo Excel** com colunas **ID**, **DESCRICAO** e **PAI**.
- **Colar direto do Excel** (Ctrl+V) na grade.
- Botões **Aplicar hierarquia** (monta a árvore e o relatório) e **Limpar**.
- **Preencher exemplo** com uma hierarquia mínima para testes.
- **Desfazer** alterações na grade com **Ctrl+Z** (Windows/Linux) ou **Cmd+Z** (macOS).
- **Ctrl+C / Ctrl+X** na grade com cópia/recorte para a área de transferência.

### Validação e busca

- **Validar hierarquia**: abre modal com resumo (total, válidas, com erro) e lista de problemas (ex.: pai inexistente).
- **Buscar ID** e **Buscar descrição**: filtram linhas do relatório mantendo contexto dos níveis superiores visíveis.
- **Localizar e substituir** em IDs e/ou descrições.
- Indicador de **linhas válidas | com erro** na grade.

### Visualização do relatório

- Tabela com **duas linhas de cabeçalho** (mês por extenso e ano/mês) e primeira coluna com **árvore expansível** (setas para abrir/fechar ramos).
- **Expandir tudo** / **Recolher tudo** na barra imediatamente acima da tabela.
- **Relatório em página inteira**: ocupa a área da aba; esconde cabeçalho, importação e toolbar superior; **Esc** volta ao modo normal.
- Contador de **itens** na árvore.

### Aparência e idioma

- **Tema claro / escuro** (ícone lua/sol), preferência salva no navegador.
- **Português / inglês** (ícone de idioma).

### Exportação

- **Exportar Excel** (`.xlsx`) da hierarquia atual (grade ou dados aplicados).
- **Baixar PDF do relatório**: captura visual da área da tabela, **A4 paisagem**, várias páginas se necessário.

---

## Fluxo de uso rápido

1. Cole na grade os dados (ID, DESCRICAO, PAI) ou clique em **Preencher exemplo**.
2. Clique em **Aplicar hierarquia**.
3. Use as setas na primeira coluna do relatório para expandir/recolher níveis.
4. (Opcional) **Validar hierarquia**, **Exportar Excel** ou **Baixar PDF do relatório**.
5. (Opcional) **Relatório em página inteira** para apresentar só a tabela; **Esc** para sair.

---

## Formato da grade (ID, descrição, pai)

| Coluna | Significado |
|--------|-------------|
| **ID** | Identificador único do nó. |
| **DESCRICAO** | Texto exibido na árvore. |
| **PAI** | ID do pai; vazio ou valor sentinela de raiz conforme seus dados (o exemplo usa hierarquia com raiz explícita). |

Erros comuns detectados na validação: **PAI** referenciando um ID que não existe na lista.

---

## Relatório na tabela

- Colunas de valores são geradas para os **12 meses do ano corrente** (formato ano+mês no cabeçalho).
- A coluna indicadores mostra a **hierarquia indentada** com controles de expandir/recolher.
- Filtros de busca reduzem as linhas visíveis sem perder a cadeia de ancestrais necessária para contexto.

---

## Licença

Este repositório está sob a licença **MIT** — uso livre com poucas condições, inclusive em contexto corporativo, desde que se mantenha o aviso de copyright.

```
MIT License

Copyright (c) 2026 Zkerrs

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

> **Autoria:** código e documentação deste projeto — **Zkerrs**, salvo trechos de bibliotecas de terceiros (SheetJS, html2canvas, jsPDF, Font Awesome), cada qual sob a respectiva licença.

*Se quiser o texto jurídico também em arquivo separado, copie o bloco acima para um arquivo `LICENSE` na raiz do repositório.*

---

## Créditos

<div align="center">

**Manipulador de Hierarquia** · modelagem e visualização de hierarquias para fluxos de trabalho e aprovação com gestores (Excel + PDF)

Se este README ajudou, considere dar uma estrela no repositório.

</div>
