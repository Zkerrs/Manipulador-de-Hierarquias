const hierarchyItems = [];

function buildHierarchy(items) {
  const map = new Map();
  const roots = [];

  for (const item of items) {
    map.set(item.id, { id: item.id, label: item.label, children: [] });
  }

  for (const item of items) {
    const current = map.get(item.id);
    if (!item.parentId) {
      roots.push(current);
      continue;
    }
    const parent = map.get(item.parentId);
    if (parent) {
      parent.children.push(current);
    } else {
      roots.push(current);
    }
  }

  return roots;
}

function coletarErrosDePai(itens) {
  const ids = new Set(itens.map((item) => item.id));
  const erros = [];
  for (let i = 0; i < itens.length; i++) {
    const item = itens[i];
    if (!item.parentId) continue;
    if (!ids.has(item.parentId)) {
      erros.push(`Linha ${i + 1}: pai inexistente '${item.parentId}' para ID '${item.id}'.`);
    }
  }
  return erros;
}

let itensDaHierarquia = Array.isArray(window.HIERARQUIA_ITENS) && window.HIERARQUIA_ITENS.length
  ? window.HIERARQUIA_ITENS
  : hierarchyItems;
let hierarchy = buildHierarchy(itensDaHierarquia);
const colunasAnoMes = gerarColunasAnoMes();
const columns = colunasAnoMes.length;
const openNodes = new Set(["FLUXO_CAIXA_0000"]);
const historicoGrade = [];
const refazerGrade = [];
let celulaAtiva = null;
const chaveTema = "manipulador-hierarquia-tema";
const chaveIdioma = "manipulador-hierarquia-idioma";
let idiomaAtual = localStorage.getItem(chaveIdioma) === "en" ? "en" : "pt";
let selecionando = false;
let ancoraSelecao = null;
const celulasSelecionadas = new Set();
let filtroId = "";
let filtroDesc = "";

const I18N = {
  pt: {
    title: "Manipulador de Hierarquia",
    subtitle: "Visualizacao e manipulacao de hierarquias com navegacao em cascata por niveis",
    importTitle: "Importar hierarquia por texto",
    importHint: "Grade estilo Excel: cole direto do Excel (Ctrl+V). Colunas: <strong>ID</strong>, <strong>DESCRICAO</strong>, <strong>PAI</strong>.",
    toolbarHint: "Use as setas para abrir e fechar cada ramo da arvore.",
    btnExpand: "Expandir tudo",
    btnCollapse: "Recolher tudo",
    btnExample: "Preencher exemplo",
    btnExport: "Exportar Excel",
    btnValidate: "Validar hierarquia",
    btnReplaceOpen: "Localizar e substituir",
    btnApply: "Aplicar hierarquia",
    btnClear: "Limpar",
    searchIdPh: "Buscar ID",
    searchDescPh: "Buscar Descrição",
    healthDefault: "0 linhas válidas | 0 com erro",
    validationTitle: "Relatório de validação",
    replaceTitle: "Localizar e substituir",
    replaceHint: "Substitui em massa os indicadores carregados, sem editar item por item.",
    replaceBtn: "Substituir tudo",
    findPh: "Localizar (ex: PAGTO)",
    replPh: "Substituir por (ex: PAGAMENTO)",
    labelIds: "Aplicar nos codigos (ID)",
    labelLabels: "Aplicar nas descricoes",
    month: "Mes",
    indicator: "Ano/Mês",
    items: "itens",
    months: ["Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"],
    btnReportFillPage: "Relatório em página inteira",
    btnReportExitFillPage: "Voltar à visualização normal",
    ariaReportFillPage: "Usar a página inteira só para o relatório",
    ariaReportExitFillPage: "Sair do modo página inteira e mostrar importação e ferramentas",
    btnExportPdf: "Baixar PDF do relatório",
    btnExportPdfBusy: "Gerando PDF…",
    ariaExportPdf: "Baixar uma imagem do relatório em arquivo PDF",
    pdfErrEmpty: "Não há linhas no relatório para exportar.",
    pdfErrLibs: "Bibliotecas de PDF não carregaram. Verifique a conexão.",
    pdfErrGeneric: "Não foi possível gerar o PDF.",
    pdfOk: "PDF do relatório baixado.",
    pdfSpinnerStatus: "Gerando PDF, aguarde"
  },
  en: {
    title: "Hierarchy Manager",
    subtitle: "View and edit hierarchies with cascade navigation by levels",
    importTitle: "Import hierarchy from text",
    importHint: "Excel-like grid: paste directly from Excel (Ctrl+V). Columns: <strong>ID</strong>, <strong>DESCRIPTION</strong>, <strong>PARENT</strong>.",
    toolbarHint: "Use arrows to open and close each branch.",
    btnExpand: "Expand all",
    btnCollapse: "Collapse all",
    btnExample: "Fill sample",
    btnExport: "Export Excel",
    btnValidate: "Validate hierarchy",
    btnReplaceOpen: "Find and replace",
    btnApply: "Apply hierarchy",
    btnClear: "Clear",
    searchIdPh: "Search ID",
    searchDescPh: "Search Description",
    healthDefault: "0 valid rows | 0 with error",
    validationTitle: "Validation report",
    replaceTitle: "Find and replace",
    replaceHint: "Bulk replace indicator values without editing one by one.",
    replaceBtn: "Replace all",
    findPh: "Find (e.g. PAY)",
    replPh: "Replace with (e.g. PAYMENT)",
    labelIds: "Apply to codes (ID)",
    labelLabels: "Apply to descriptions",
    month: "Month",
    indicator: "Year/Month",
    items: "items",
    months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
    btnReportFillPage: "Report fills the page",
    btnReportExitFillPage: "Back to normal view",
    ariaReportFillPage: "Use the full page for the report only",
    ariaReportExitFillPage: "Leave full-page report mode and show import and tools",
    btnExportPdf: "Download report PDF",
    btnExportPdfBusy: "Generating PDF…",
    ariaExportPdf: "Download a snapshot of the report as a PDF file",
    pdfErrEmpty: "There are no rows in the report to export.",
    pdfErrLibs: "PDF libraries failed to load. Check your connection.",
    pdfErrGeneric: "Could not generate the PDF.",
    pdfOk: "Report PDF downloaded.",
    pdfSpinnerStatus: "Generating PDF, please wait"
  }
};

function t(chave) {
  return I18N[idiomaAtual][chave] ?? chave;
}

function relatorioPaginaInteiraAtivo() {
  return document.body.classList.contains("report-fill-page");
}

function atualizarBarraModoRelatorioPagina() {
  const btn = document.getElementById("toggleReportFillPage");
  const icon = document.getElementById("toggleReportFillPageIcon");
  const label = document.getElementById("toggleReportFillPageLabel");
  const ativo = relatorioPaginaInteiraAtivo();
  if (icon) {
    icon.classList.remove("fa-maximize", "fa-minimize");
    icon.classList.add(ativo ? "fa-minimize" : "fa-maximize");
  }
  if (label) label.textContent = ativo ? t("btnReportExitFillPage") : t("btnReportFillPage");
  if (btn) {
    btn.setAttribute("aria-label", ativo ? t("ariaReportExitFillPage") : t("ariaReportFillPage"));
    btn.classList.toggle("primary", !ativo);
  }
}

function definirRelatorioPaginaInteira(ativo) {
  document.body.classList.toggle("report-fill-page", ativo);
  atualizarBarraModoRelatorioPagina();
}

function aplicarIdioma(idioma) {
  idiomaAtual = idioma === "en" ? "en" : "pt";
  localStorage.setItem(chaveIdioma, idiomaAtual);

  const byIdText = (id, valor) => {
    const el = document.getElementById(id);
    if (el) el.textContent = valor;
  };

  byIdText("titleText", t("title"));
  byIdText("subtitleText", t("subtitle"));
  byIdText("importTitleText", t("importTitle"));
  const importHint = document.getElementById("importHintText");
  if (importHint) importHint.innerHTML = t("importHint");
  byIdText("toolbarHintText", t("toolbarHint"));
  byIdText("expandAll", t("btnExpand"));
  byIdText("collapseAll", t("btnCollapse"));
  byIdText("fillExample", t("btnExample"));
  byIdText("exportExcel", t("btnExport"));
  byIdText("validateHierarchy", t("btnValidate"));
  byIdText("openReplaceModal", t("btnReplaceOpen"));
  byIdText("applyHierarchy", t("btnApply"));
  byIdText("clearInput", t("btnClear"));
  byIdText("replaceModalTitle", t("replaceTitle"));
  byIdText("replaceModalHint", t("replaceHint"));
  byIdText("replaceAll", t("replaceBtn"));

  const find = document.getElementById("findText");
  const repl = document.getElementById("replaceText");
  if (find) find.placeholder = t("findPh");
  if (repl) repl.placeholder = t("replPh");
  const searchIdInput = document.getElementById("searchId");
  const searchDescInput = document.getElementById("searchDesc");
  if (searchIdInput) searchIdInput.placeholder = t("searchIdPh");
  if (searchDescInput) searchDescInput.placeholder = t("searchDescPh");
  const validationTitle = document.querySelector("#validationModal .modal-title");
  if (validationTitle) validationTitle.textContent = t("validationTitle");

  const pdfBtn = document.getElementById("exportReportPdf");
  const pdfLabel = document.getElementById("exportReportPdfLabel");
  if (pdfBtn && pdfLabel && !pdfBtn.classList.contains("report-pdf-busy")) {
    pdfLabel.textContent = t("btnExportPdf");
  }
  if (pdfBtn) pdfBtn.setAttribute("aria-label", t("ariaExportPdf"));
  const pdfSpinner = document.getElementById("exportReportPdfSpinner");
  if (pdfSpinner) pdfSpinner.setAttribute("aria-label", t("pdfSpinnerStatus"));

  atualizarBarraModoRelatorioPagina();

  const labels = document.querySelectorAll(".replace-options label");
  if (labels[0]) labels[0].childNodes[1].nodeValue = ` ${t("labelIds")}`;
  if (labels[1]) labels[1].childNodes[1].nodeValue = ` ${t("labelLabels")}`;

  renderHeader();
  render();
}

function gerarColunasAnoMes() {
  const anoAtual = new Date().getFullYear();
  const colunas = [];
  for (let mes = 1; mes <= 12; mes++) {
    colunas.push(`${anoAtual}${String(mes).padStart(2, "0")}`);
  }
  return colunas;
}

function renderHeader() {
  const monthRow = document.getElementById("monthRow");
  const headerRow = document.getElementById("headerRow");
  monthRow.innerHTML = "";
  headerRow.innerHTML = "";

  const thMes = document.createElement("th");
  thMes.textContent = t("month");
  monthRow.appendChild(thMes);

  const thIndicador = document.createElement("th");
  thIndicador.textContent = t("indicator");
  headerRow.appendChild(thIndicador);

  const nomesMeses = t("months");

  for (let i = 0; i < colunasAnoMes.length; i++) {
    const thMesNome = document.createElement("th");
    thMesNome.textContent = nomesMeses[i];
    monthRow.appendChild(thMesNome);
  }

  for (const coluna of colunasAnoMes) {
    const th = document.createElement("th");
    th.textContent = coluna;
    headerRow.appendChild(th);
  }
}

function parseHierarchyRows(linhas) {
  const itens = [];
  let encontrouRoot = false;

  for (let i = 0; i < linhas.length; i++) {
    const partes = linhas[i];

    if (partes.length !== 3) {
      throw new Error(`Linha ${i + 1} invalida. Use 3 colunas: ID, DESCRICAO, PAI.`);
    }

    const [id, label, parentRaw] = partes;
    if (!id || !label) {
      throw new Error(`Linha ${i + 1} invalida: ID e DESCRICAO obrigatorios.`);
    }

    if (parentRaw === "<root>") {
      encontrouRoot = true;
    }

    const parentId = parentRaw === "<root>" ? null : parentRaw;
    itens.push({ id, label, parentId });
  }

  if (!encontrouRoot) {
    throw new Error("Obrigatorio informar pelo menos uma linha com <root> na coluna PAI.");
  }

  return itens;
}

function criarGradeHierarquia(totalLinhas = 35) {
  const tbody = document.getElementById("hierarchyGridBody");
  tbody.innerHTML = "";
  adicionarLinhasGrade(totalLinhas);
  limparSelecaoCelulas();
}

function adicionarLinhasGrade(quantidade) {
  const tbody = document.getElementById("hierarchyGridBody");
  const inicio = tbody.children.length;
  for (let row = inicio; row < inicio + quantidade; row++) {
    const tr = document.createElement("tr");
    for (let col = 0; col < 3; col++) {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.type = "text";
      input.dataset.row = String(row);
      input.dataset.col = String(col);
      td.appendChild(input);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
}

function garantirQuantidadeLinhas(totalNecessario) {
  const tbody = document.getElementById("hierarchyGridBody");
  const atual = tbody.children.length;
  if (totalNecessario > atual) {
    adicionarLinhasGrade(totalNecessario - atual);
  }
}

function obterMatrizDaGrade() {
  const linhas = [];
  const inputs = document.querySelectorAll("#hierarchyGridBody input");
  const mapa = new Map();

  for (const input of inputs) {
    const r = Number(input.dataset.row);
    const c = Number(input.dataset.col);
    const valor = input.value.trim();
    if (!mapa.has(r)) mapa.set(r, ["", "", ""]);
    mapa.get(r)[c] = valor;
  }

  for (let i = 0; i < mapa.size; i++) {
    const linha = mapa.get(i) ?? ["", "", ""];
    if (linha.some((v) => v !== "")) {
      linhas.push(linha);
    }
  }
  return linhas;
}

function obterLinhasParaExportacaoHierarquia() {
  const linhasGrade = obterMatrizDaGrade();
  return linhasGrade.length
    ? linhasGrade
    : itensDaHierarquia.map((item) => [item.id, item.label, item.parentId ?? "<root>"]);
}

function obterNomeBaseArquivoHierarquia(linhasJaCarregadas) {
  const linhas = linhasJaCarregadas ?? obterLinhasParaExportacaoHierarquia();
  const descricaoPrimeiro = (linhas[0]?.[1] || "hierarquia")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|]/g, "")
    .trim()
    .replace(/\s+/g, "_");
  return descricaoPrimeiro || "hierarquia";
}

function normalizarMatriz(matriz) {
  return matriz.map((linha) => [linha[0] ?? "", linha[1] ?? "", linha[2] ?? ""]);
}

function serializarMatriz(matriz) {
  return JSON.stringify(normalizarMatriz(matriz));
}

function salvarEstadoNoHistorico() {
  const atual = obterMatrizDaGrade();
  const assinaturaAtual = serializarMatriz(atual);
  const ultimo = historicoGrade[historicoGrade.length - 1];
  if (!ultimo || serializarMatriz(ultimo) !== assinaturaAtual) {
    historicoGrade.push(normalizarMatriz(atual));
    if (historicoGrade.length > 200) historicoGrade.shift();
    refazerGrade.length = 0;
  }
}

function aplicarMatrizNaGrade(matriz) {
  const itens = matriz.map((linha) => ({
    id: linha[0] ?? "",
    label: linha[1] ?? "",
    parentId: (linha[2] ?? "") === "<root>" || (linha[2] ?? "") === "" ? null : linha[2]
  }));
  preencherGradeComItens(itens.filter((item) => item.id || item.label || item.parentId));
}

function desfazerGrade() {
  if (historicoGrade.length <= 1) return;
  const estadoAtual = historicoGrade.pop();
  refazerGrade.push(estadoAtual);
  const estadoAnterior = historicoGrade[historicoGrade.length - 1];
  aplicarMatrizNaGrade(estadoAnterior);
}

function ajustarLarguraColunasGrade() {
  const limites = [
    { min: 6, max: 40 },
    { min: 8, max: 100 },
    { min: 6, max: 40 }
  ];
  const maxChars = [2, 9, 3];
  const inputs = document.querySelectorAll("#hierarchyGridBody input");

  for (const input of inputs) {
    const col = Number(input.dataset.col);
    const tamanho = input.value.trim().length;
    if (!Number.isNaN(col) && tamanho > maxChars[col]) {
      maxChars[col] = tamanho;
    }
  }

  const colIds = ["colId", "colDescricao", "colPai"];
  for (let i = 0; i < colIds.length; i++) {
    const larguraCh = Math.min(limites[i].max, Math.max(limites[i].min, maxChars[i] + 0.6));
    const col = document.getElementById(colIds[i]);
    if (col) col.style.width = `${larguraCh}ch`;
  }
}

function keyCelula(row, col) {
  return `${row}:${col}`;
}

function parseKeyCelula(chave) {
  const [r, c] = chave.split(":").map(Number);
  return { row: r, col: c };
}

function getInputPorRowCol(row, col) {
  return document.querySelector(`#hierarchyGridBody input[data-row="${row}"][data-col="${col}"]`);
}

function limparSelecaoCelulas() {
  for (const chave of celulasSelecionadas) {
    const { row, col } = parseKeyCelula(chave);
    const input = getInputPorRowCol(row, col);
    if (input?.parentElement) input.parentElement.classList.remove("selected-cell");
  }
  celulasSelecionadas.clear();
}

function aplicarSelecaoRetangular(inicio, fim) {
  if (!inicio || !fim) return;
  limparSelecaoCelulas();
  const minRow = Math.min(inicio.row, fim.row);
  const maxRow = Math.max(inicio.row, fim.row);
  const minCol = Math.min(inicio.col, fim.col);
  const maxCol = Math.max(inicio.col, fim.col);
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      const input = getInputPorRowCol(r, c);
      if (!input) continue;
      const chave = keyCelula(r, c);
      celulasSelecionadas.add(chave);
      if (input.parentElement) input.parentElement.classList.add("selected-cell");
    }
  }
}

function getSelecaoBounds() {
  if (!celulasSelecionadas.size) return null;
  let minRow = Infinity;
  let maxRow = -Infinity;
  let minCol = Infinity;
  let maxCol = -Infinity;
  for (const chave of celulasSelecionadas) {
    const { row, col } = parseKeyCelula(chave);
    minRow = Math.min(minRow, row);
    maxRow = Math.max(maxRow, row);
    minCol = Math.min(minCol, col);
    maxCol = Math.max(maxCol, col);
  }
  return { minRow, maxRow, minCol, maxCol };
}

function matrizDaSelecao() {
  const b = getSelecaoBounds();
  if (!b) return [];
  const linhas = [];
  for (let r = b.minRow; r <= b.maxRow; r++) {
    const linha = [];
    for (let c = b.minCol; c <= b.maxCol; c++) {
      const input = getInputPorRowCol(r, c);
      linha.push(input ? input.value : "");
    }
    linhas.push(linha);
  }
  return linhas;
}

function textoDaSelecao() {
  return matrizDaSelecao().map((l) => l.join("\t")).join("\n");
}

function aplicarMatrizEm(rowInicio, colInicio, matriz) {
  garantirQuantidadeLinhas(rowInicio + matriz.length + 20);
  for (let i = 0; i < matriz.length; i++) {
    for (let j = 0; j < matriz[i].length; j++) {
      const c = colInicio + j;
      if (c > 2) continue;
      const input = getInputPorRowCol(rowInicio + i, c);
      if (input) input.value = matriz[i][j] ?? "";
    }
  }
}

function preencherGradeComItens(itens) {
  const totalNecessario = Math.max(35, itens.length + 8);
  criarGradeHierarquia(totalNecessario);
  itens.forEach((item, idx) => {
    const linha = [item.id, item.label, item.parentId ?? "<root>"];
    linha.forEach((valor, col) => {
      const input = document.querySelector(`#hierarchyGridBody input[data-row="${idx}"][data-col="${col}"]`);
      if (input) input.value = valor;
    });
  });
  ajustarLarguraColunasGrade();
}

function aplicarPasteExcel(event) {
  const alvo = event.target;
  const alvoNaGrade = alvo instanceof HTMLInputElement && alvo.closest("#hierarchyGridBody");
  if (!alvoNaGrade && celulasSelecionadas.size === 0) return;

  const texto = event.clipboardData?.getData("text");
  if (!texto) return;

  salvarEstadoNoHistorico();
  event.preventDefault();
  const linhas = texto.replace(/\r/g, "").split("\n").filter((l) => l.length > 0).map((l) => l.split("\t").map((v) => v.trim()));
  const bounds = getSelecaoBounds();
  const inicioRow = bounds ? bounds.minRow : Number(alvo.dataset.row);
  const inicioCol = bounds ? bounds.minCol : Number(alvo.dataset.col);

  if (bounds) {
    const linhasSelecao = bounds.maxRow - bounds.minRow + 1;
    const colunasSelecao = bounds.maxCol - bounds.minCol + 1;
    const linhasPaste = linhas.length;
    const colunasPaste = Math.max(...linhas.map((l) => l.length));

    // Comportamento estilo Excel: 1 celula copiada + bloco selecionado => replica em todo bloco.
    if (linhasPaste === 1 && colunasPaste === 1) {
      const valor = linhas[0][0] ?? "";
      for (const chave of celulasSelecionadas) {
        const { row, col } = parseKeyCelula(chave);
        const input = getInputPorRowCol(row, col);
        if (input) input.value = valor;
      }
      ajustarLarguraColunasGrade();
      salvarEstadoNoHistorico();
      return;
    }

    // Se tamanho do bloco copiado bater com o bloco selecionado, cola respeitando a area.
    if (linhasPaste === linhasSelecao && colunasPaste === colunasSelecao) {
      for (let r = 0; r < linhasSelecao; r++) {
        for (let c = 0; c < colunasSelecao; c++) {
          const input = getInputPorRowCol(bounds.minRow + r, bounds.minCol + c);
          if (input) input.value = linhas[r]?.[c] ?? "";
        }
      }
      ajustarLarguraColunasGrade();
      salvarEstadoNoHistorico();
      return;
    }
  }

  aplicarMatrizEm(inicioRow, inicioCol, linhas);
  ajustarLarguraColunasGrade();
  salvarEstadoNoHistorico();
}

function flatten(nodes, level = 0, parentVisible = true) {
  const list = [];
  for (const node of nodes) {
    const hasChildren = Array.isArray(node.children) && node.children.length > 0;
    const visible = parentVisible;
    list.push({ ...node, level, hasChildren, visible });
    const isOpen = openNodes.has(node.id);
    if (hasChildren) {
      list.push(...flatten(node.children, level + 1, visible && isOpen));
    }
  }
  return list;
}

function buildEmptyValueCell() {
  const td = document.createElement("td");
  td.textContent = "—";
  return td;
}

function render() {
  const tbody = document.getElementById("rows");
  tbody.innerHTML = "";
  const rows = flatten(hierarchy);
  const rowsFiltradas = aplicarFiltroRows(rows);
  document.getElementById("nodeCount").textContent = `${rowsFiltradas.length} ${t("items")}`;

  for (const row of rowsFiltradas) {

    const tr = document.createElement("tr");
    tr.className = `level-${row.level}`;

    const nameTd = document.createElement("td");
    const nameWrap = document.createElement("div");
    nameWrap.className = "name-cell";
    nameWrap.style.paddingLeft = `${row.level * 18}px`;

    if (row.hasChildren) {
      const btn = document.createElement("button");
      btn.className = "toggle";
      btn.type = "button";
      btn.textContent = openNodes.has(row.id) ? "▼" : "▶";
      btn.setAttribute("aria-label", "Expandir ou recolher");
      btn.onclick = () => {
        if (openNodes.has(row.id)) {
          openNodes.delete(row.id);
        } else {
          openNodes.add(row.id);
        }
        render();
      };
      nameWrap.appendChild(btn);
    } else {
      const spacer = document.createElement("span");
      spacer.className = "spacer";
      nameWrap.appendChild(spacer);
    }

    const label = document.createElement("span");
    label.className = "node-label";
    label.textContent = row.label;
    nameWrap.appendChild(label);

    nameTd.appendChild(nameWrap);
    tr.appendChild(nameTd);

    for (let i = 0; i < columns; i++) {
      tr.appendChild(buildEmptyValueCell());
    }

    tbody.appendChild(tr);
  }
}

function aplicarFiltroRows(rows) {
  const qId = filtroId.trim().toLowerCase();
  const qDesc = filtroDesc.trim().toLowerCase();
  if (!qId && !qDesc) return rows.filter((r) => r.visible);

  const mapPorId = new Map(rows.map((r) => [r.id, r]));
  const incluir = new Set();

  for (const row of rows) {
    if (!row.visible) continue;
    const bateId = !qId || row.id.toLowerCase().includes(qId);
    const bateDesc = !qDesc || row.label.toLowerCase().includes(qDesc);
    if (!(bateId && bateDesc)) continue;

    incluir.add(row.id);
    let nivelPai = row.level - 1;
    for (let i = rows.indexOf(row) - 1; i >= 0 && nivelPai >= 0; i--) {
      if (rows[i].level === nivelPai) {
        incluir.add(rows[i].id);
        nivelPai--;
      }
    }
  }

  return rows.filter((r) => r.visible && incluir.has(r.id) && mapPorId.has(r.id));
}

function collectParentNodes(nodes, set) {
  for (const n of nodes) {
    if (n.children && n.children.length) {
      set.add(n.id);
      collectParentNodes(n.children, set);
    }
  }
}

function expandirTudoHierarquia() {
  openNodes.clear();
  collectParentNodes(hierarchy, openNodes);
  render();
}

function recolherTudoHierarquia() {
  openNodes.clear();
  render();
}

function obterConstrutorJsPDF() {
  return window.jspdf?.jsPDF || window.jsPDF;
}

function montarNomeArquivoPdfRelatorio() {
  return `${obterNomeBaseArquivoHierarquia()}_Hierarquia.pdf`;
}

function adicionarCanvasAoPdfPaisagem(canvas, pdf) {
  const imgData = canvas.toDataURL("image/png", 1.0);
  const margin = 10;
  const pdfW = pdf.internal.pageSize.getWidth();
  const pdfH = pdf.internal.pageSize.getHeight();
  const pageInnerH = pdfH - 2 * margin;
  const imgWidthMm = pdfW - 2 * margin;
  const imgHeightMm = (canvas.height / canvas.width) * imgWidthMm;

  let heightLeft = imgHeightMm;
  let y = margin;
  pdf.addImage(imgData, "PNG", margin, y, imgWidthMm, imgHeightMm);
  heightLeft -= pageInnerH;

  while (heightLeft > 0) {
    y = margin - (imgHeightMm - heightLeft);
    pdf.addPage();
    pdf.addImage(imgData, "PNG", margin, y, imgWidthMm, imgHeightMm);
    heightLeft -= pageInnerH;
  }
}

async function exportarRelatorioComoPdf() {
  const area = document.querySelector(".table-area");
  const tbody = document.getElementById("rows");
  const btn = document.getElementById("exportReportPdf");
  const labelSpan = document.getElementById("exportReportPdfLabel");
  const spinner = document.getElementById("exportReportPdfSpinner");
  const popupErro = document.getElementById("popupErro");
  const popupSucesso = document.getElementById("popupSucesso");

  const toastErro = (msg) => {
    if (!popupErro) return;
    popupSucesso?.classList.remove("show");
    popupErro.textContent = msg;
    popupErro.classList.add("show");
    clearTimeout(toastErro._t);
    toastErro._t = setTimeout(() => popupErro.classList.remove("show"), 4200);
  };
  const toastOk = (msg) => {
    if (!popupSucesso) return;
    popupErro?.classList.remove("show");
    popupSucesso.textContent = msg;
    popupSucesso.classList.add("show");
    clearTimeout(toastOk._t);
    toastOk._t = setTimeout(() => popupSucesso.classList.remove("show"), 3200);
  };

  if (!area || !tbody || !btn) return;
  if (!tbody.children.length) {
    toastErro(t("pdfErrEmpty"));
    return;
  }
  const JsPDF = obterConstrutorJsPDF();
  if (typeof html2canvas === "undefined" || typeof JsPDF !== "function") {
    toastErro(t("pdfErrLibs"));
    return;
  }

  const textoNormal = t("btnExportPdf");
  const textoBusy = t("btnExportPdfBusy");
  btn.classList.add("report-pdf-busy");
  btn.setAttribute("aria-busy", "true");
  if (labelSpan) labelSpan.textContent = textoBusy;
  if (spinner) spinner.hidden = false;

  const escuro = document.body.classList.contains("dark");
  const bg = escuro ? "#22272f" : "#ffffff";

  try {
    const w = area.scrollWidth;
    const h = area.scrollHeight;
    const canvas = await html2canvas(area, {
      scale: 2,
      useCORS: true,
      logging: false,
      width: w,
      height: h,
      windowWidth: w,
      windowHeight: h,
      scrollX: 0,
      scrollY: 0,
      backgroundColor: bg,
      onclone(clonedDoc) {
        const cloneArea = clonedDoc.querySelector(".table-area");
        if (!cloneArea) return;
        cloneArea.style.overflow = "visible";
        cloneArea.style.height = "auto";
        cloneArea.style.maxHeight = "none";
        cloneArea.querySelectorAll(".report-table th, .report-table td").forEach((cell) => {
          cell.style.position = "static";
          cell.style.left = "auto";
          cell.style.top = "auto";
          cell.style.boxShadow = "none";
        });
      }
    });

    const pdf = new JsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    adicionarCanvasAoPdfPaisagem(canvas, pdf);
    pdf.save(montarNomeArquivoPdfRelatorio());
    toastOk(t("pdfOk"));
  } catch {
    toastErro(t("pdfErrGeneric"));
  } finally {
    btn.classList.remove("report-pdf-busy");
    btn.removeAttribute("aria-busy");
    if (labelSpan) labelSpan.textContent = textoNormal;
    if (spinner) spinner.hidden = true;
  }
}

document.getElementById("expandAll").addEventListener("click", expandirTudoHierarquia);
document.getElementById("collapseAll").addEventListener("click", recolherTudoHierarquia);

document.getElementById("applyHierarchy").addEventListener("click", () => {
  const status = document.getElementById("importStatus");
  const excelWrapper = document.getElementById("excelWrapper");
  const popupErro = document.getElementById("popupErro");
  const popupSucesso = document.getElementById("popupSucesso");

  function mostrarPopupErro(mensagem) {
    popupSucesso.classList.remove("show");
    popupErro.textContent = mensagem;
    popupErro.classList.add("show");
    clearTimeout(mostrarPopupErro.timerId);
    mostrarPopupErro.timerId = setTimeout(() => {
      popupErro.classList.remove("show");
    }, 3800);
  }

  function mostrarPopupSucesso(mensagem) {
    popupErro.classList.remove("show");
    popupSucesso.textContent = mensagem;
    popupSucesso.classList.add("show");
    clearTimeout(mostrarPopupSucesso.timerId);
    mostrarPopupSucesso.timerId = setTimeout(() => {
      popupSucesso.classList.remove("show");
    }, 2800);
  }

  try {
    const linhasGrade = obterMatrizDaGrade();
    const novosItens = parseHierarchyRows(linhasGrade);
    const errosDePai = coletarErrosDePai(novosItens);
    itensDaHierarquia = novosItens;
    hierarchy = buildHierarchy(itensDaHierarquia);
    openNodes.clear();
    for (const item of itensDaHierarquia) {
      if (!item.parentId) openNodes.add(item.id);
    }
    render();
    ajustarLarguraColunasGrade();
    excelWrapper.classList.remove("input-error");
    if (errosDePai.length) {
      status.textContent = `Hierarquia aplicada com aviso: ${errosDePai.length} erro(s) de pai.`;
      mostrarPopupErro(`Hierarquia aplicada com erro: ${errosDePai.length} pai(s) inexistente(s).`);
    } else {
      status.textContent = `Hierarquia aplicada com sucesso (${novosItens.length} linhas).`;
      mostrarPopupSucesso("Hierarquia valida e aplicada com sucesso.");
    }
  } catch (error) {
    excelWrapper.classList.add("input-error");
    status.textContent = `Erro: ${error.message}`;
    mostrarPopupErro(`Erro na importacao: ${error.message}`);
  }
});

document.getElementById("clearInput").addEventListener("click", () => {
  salvarEstadoNoHistorico();
  preencherGradeComItens([]);
  ajustarLarguraColunasGrade();
  document.getElementById("importStatus").textContent = "Texto limpo.";
  salvarEstadoNoHistorico();
});

document.getElementById("fillExample").addEventListener("click", () => {
  salvarEstadoNoHistorico();
  const exemplo = [
    { id: "HIER_0000", label: "HIERARQUIA PRINCIPAL", parentId: null },
    { id: "HIER_0100", label: "FINANCEIRO", parentId: "HIER_0000" },
    { id: "HIER_0200", label: "RECEITAS", parentId: "HIER_0100" },
    { id: "HIER_0300", label: "DESPESAS", parentId: "HIER_0100" }
  ];
  preencherGradeComItens(exemplo);
  document.getElementById("importStatus").textContent = "Exemplo preenchido. Clique em Aplicar hierarquia.";
  salvarEstadoNoHistorico();
});

const replaceModal = document.getElementById("replaceModal");
const openReplaceModalBtn = document.getElementById("openReplaceModal");
const closeReplaceModalBtn = document.getElementById("closeReplaceModal");

function abrirModalSubstituicao() {
  replaceModal.classList.add("show");
  replaceModal.setAttribute("aria-hidden", "false");
  document.getElementById("findText").focus();
}

function fecharModalSubstituicao() {
  replaceModal.classList.remove("show");
  replaceModal.setAttribute("aria-hidden", "true");
}

openReplaceModalBtn.addEventListener("click", abrirModalSubstituicao);
closeReplaceModalBtn.addEventListener("click", fecharModalSubstituicao);

replaceModal.addEventListener("click", (event) => {
  if (event.target === replaceModal) {
    fecharModalSubstituicao();
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key !== "Escape") return;
  if (replaceModal.classList.contains("show")) {
    fecharModalSubstituicao();
    return;
  }
  const modalValidacaoEsc = document.getElementById("validationModal");
  if (modalValidacaoEsc && modalValidacaoEsc.classList.contains("show")) {
    modalValidacaoEsc.classList.remove("show");
    modalValidacaoEsc.setAttribute("aria-hidden", "true");
    return;
  }
  if (!relatorioPaginaInteiraAtivo()) return;
  const alvo = event.target;
  const tag = alvo && alvo.tagName;
  if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT" || alvo?.isContentEditable) return;
  event.preventDefault();
  definirRelatorioPaginaInteira(false);
});

document.getElementById("replaceAll").addEventListener("click", () => {
  salvarEstadoNoHistorico();
  const status = document.getElementById("importStatus");
  const findText = document.getElementById("findText").value;
  const replaceText = document.getElementById("replaceText").value;
  const replaceInIds = document.getElementById("replaceInIds").checked;
  const replaceInLabels = document.getElementById("replaceInLabels").checked;

  if (!itensDaHierarquia.length) {
    status.textContent = "Erro: carregue uma hierarquia antes de substituir.";
    return;
  }
  if (!findText) {
    status.textContent = "Erro: informe o texto para localizar.";
    return;
  }
  if (!replaceInIds && !replaceInLabels) {
    status.textContent = "Erro: selecione ao menos um alvo (ID ou descricao).";
    return;
  }

  let alteracoes = 0;
  itensDaHierarquia = itensDaHierarquia.map((item) => {
    let novoId = item.id;
    let novoLabel = item.label;
    let novoParent = item.parentId;

    if (replaceInIds) {
      const idAtualizado = novoId.split(findText).join(replaceText);
      if (idAtualizado !== novoId) alteracoes++;
      novoId = idAtualizado;

      if (novoParent) {
        novoParent = novoParent.split(findText).join(replaceText);
      }
    }

    if (replaceInLabels) {
      const labelAtualizada = novoLabel.split(findText).join(replaceText);
      if (labelAtualizada !== novoLabel) alteracoes++;
      novoLabel = labelAtualizada;
    }

    return { id: novoId, label: novoLabel, parentId: novoParent };
  });

  hierarchy = buildHierarchy(itensDaHierarquia);
  openNodes.clear();
  for (const item of itensDaHierarquia) {
    if (!item.parentId) openNodes.add(item.id);
  }
  preencherGradeComItens(itensDaHierarquia);
  render();
  status.textContent = `Substituicao concluida. Itens alterados: ${alteracoes}.`;
  fecharModalSubstituicao();
  salvarEstadoNoHistorico();
});

function validarHierarquiaCompleta() {
  const linhas = obterMatrizDaGrade();
  const erros = [];
  const ids = new Map();
  let qtdRoot = 0;

  linhas.forEach((linha, idx) => {
    const [id, descricao, pai] = linha.map((v) => (v || "").trim());
    if (!id || !descricao) {
      erros.push(`Linha ${idx + 1}: ID e DESCRIÇÃO são obrigatórios.`);
    }
    if (pai === "<root>") qtdRoot++;
    if (id) {
      if (!ids.has(id)) ids.set(id, []);
      ids.get(id).push(idx + 1);
    }
  });

  for (const [id, ocorrencias] of ids.entries()) {
    if (ocorrencias.length > 1) {
      erros.push(`ID duplicado '${id}' nas linhas: ${ocorrencias.join(", ")}.`);
    }
  }

  linhas.forEach((linha, idx) => {
    const [id, _descricao, pai] = linha.map((v) => (v || "").trim());
    if (!id || !pai || pai === "<root>") return;
    if (!ids.has(pai)) erros.push(`Linha ${idx + 1}: pai inexistente '${pai}'.`);
  });

  if (qtdRoot === 0) erros.push("Nenhuma linha com <root> foi encontrada.");

  const paiDe = new Map();
  linhas.forEach((linha) => {
    const [id, _descricao, pai] = linha.map((v) => (v || "").trim());
    if (id && pai && pai !== "<root>") paiDe.set(id, pai);
  });
  const visitando = new Set();
  const visitado = new Set();
  function detectarCiclo(id) {
    if (visitando.has(id)) return true;
    if (visitado.has(id)) return false;
    visitando.add(id);
    const pai = paiDe.get(id);
    if (pai && paiDe.has(pai) && detectarCiclo(pai)) return true;
    visitando.delete(id);
    visitado.add(id);
    return false;
  }
  for (const id of paiDe.keys()) {
    if (detectarCiclo(id)) {
      erros.push("Ciclo detectado na relação de pai/filho.");
      break;
    }
  }

  const linhasComErro = new Set();
  for (const erro of erros) {
    const m = erro.match(/Linha (\d+)/i);
    if (m) linhasComErro.add(Number(m[1]));
  }
  const validas = Math.max(0, linhas.length - linhasComErro.size);
  return { erros, total: linhas.length, validas, comErro: linhasComErro.size };
}

function atualizarIndicadorSaude(resultado) {
  const health = document.getElementById("healthIndicator");
  if (!health) return;
  health.textContent = `${resultado.validas} linhas válidas | ${resultado.comErro} com erro`;
}

function mostrarRelatorioValidacao(resultado) {
  const modal = document.getElementById("validationModal");
  const resumo = document.getElementById("validationSummary");
  const lista = document.getElementById("validationList");
  if (!modal || !resumo || !lista) return;

  atualizarIndicadorSaude(resultado);
  resumo.textContent = `Total: ${resultado.total} | Válidas: ${resultado.validas} | Com erro: ${resultado.comErro}`;
  lista.innerHTML = "";
  if (!resultado.erros.length) {
    const li = document.createElement("li");
    li.textContent = "Nenhum erro encontrado. Hierarquia válida.";
    lista.appendChild(li);
  } else {
    for (const erro of resultado.erros) {
      const li = document.createElement("li");
      li.textContent = erro;
      lista.appendChild(li);
    }
  }
  modal.classList.add("show");
  modal.setAttribute("aria-hidden", "false");
}

const botaoValidar = document.getElementById("validateHierarchy");
if (botaoValidar) {
  botaoValidar.addEventListener("click", () => {
    const resultado = validarHierarquiaCompleta();
    mostrarRelatorioValidacao(resultado);
  });
}

const botaoFecharValidacao = document.getElementById("closeValidationModal");
if (botaoFecharValidacao) {
  botaoFecharValidacao.addEventListener("click", () => {
    const modal = document.getElementById("validationModal");
    if (!modal) return;
    modal.classList.remove("show");
    modal.setAttribute("aria-hidden", "true");
  });
}

const modalValidacao = document.getElementById("validationModal");
if (modalValidacao) {
  modalValidacao.addEventListener("click", (event) => {
    if (event.target === modalValidacao) {
      modalValidacao.classList.remove("show");
      modalValidacao.setAttribute("aria-hidden", "true");
    }
  });
}

document.getElementById("exportExcel").addEventListener("click", () => {
  const status = document.getElementById("importStatus");
  const popupErro = document.getElementById("popupErro");
  const popupSucesso = document.getElementById("popupSucesso");
  const linhasParaExportar = obterLinhasParaExportacaoHierarquia();

  if (!linhasParaExportar.length) {
    status.textContent = "Erro: nao ha dados para exportar.";
    popupSucesso.classList.remove("show");
    popupErro.textContent = "Erro: nao ha hierarquia para exportar.";
    popupErro.classList.add("show");
    clearTimeout(popupErro._timer);
    popupErro._timer = setTimeout(() => popupErro.classList.remove("show"), 3800);
    return;
  }

  if (typeof XLSX === "undefined") {
    status.textContent = "Erro: biblioteca de Excel nao carregada.";
    popupSucesso.classList.remove("show");
    popupErro.textContent = "Erro: biblioteca de Excel nao carregada.";
    popupErro.classList.add("show");
    clearTimeout(popupErro._timer);
    popupErro._timer = setTimeout(() => popupErro.classList.remove("show"), 3800);
    return;
  }

  const aoa = [["ID", "DESCRICAO", "PAI"], ...linhasParaExportar];
  const worksheet = XLSX.utils.aoa_to_sheet(aoa);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Hierarquia");

  const nomeArquivo = `${obterNomeBaseArquivoHierarquia(linhasParaExportar)}_Hierarquia.xlsx`;
  XLSX.writeFile(workbook, nomeArquivo);
  status.textContent = `Excel exportado com sucesso: ${nomeArquivo}`;
});

function atualizarIconeTema() {
  const icone = document.getElementById("themeIcon");
  if (!icone) return;
  const escuro = document.body.classList.contains("dark");
  icone.classList.remove("fa-moon", "fa-sun");
  icone.classList.add(escuro ? "fa-sun" : "fa-moon");
}

function aplicarTema(tema) {
  if (tema === "dark") {
    document.body.classList.add("dark");
  } else {
    document.body.classList.remove("dark");
  }
  localStorage.setItem(chaveTema, tema);
  atualizarIconeTema();
}

const botaoTema = document.getElementById("toggleTheme");
if (botaoTema) {
  botaoTema.addEventListener("click", () => {
    const escuro = document.body.classList.contains("dark");
    aplicarTema(escuro ? "light" : "dark");
  });
}

const botaoIdioma = document.getElementById("toggleLanguage");
if (botaoIdioma) {
  botaoIdioma.addEventListener("click", () => {
    aplicarIdioma(idiomaAtual === "pt" ? "en" : "pt");
  });
}

const toggleReportFill = document.getElementById("toggleReportFillPage");
if (toggleReportFill) {
  toggleReportFill.addEventListener("click", () => {
    definirRelatorioPaginaInteira(!relatorioPaginaInteiraAtivo());
  });
}

const exportReportPdfBtn = document.getElementById("exportReportPdf");
if (exportReportPdfBtn) {
  exportReportPdfBtn.addEventListener("click", () => {
    exportarRelatorioComoPdf();
  });
}

const searchIdInput = document.getElementById("searchId");
const searchDescInput = document.getElementById("searchDesc");
if (searchIdInput) {
  searchIdInput.addEventListener("input", () => {
    filtroId = searchIdInput.value;
    render();
  });
}
if (searchDescInput) {
  searchDescInput.addEventListener("input", () => {
    filtroDesc = searchDescInput.value;
    render();
  });
}

criarGradeHierarquia();
ajustarLarguraColunasGrade();
document.addEventListener("paste", aplicarPasteExcel);
document.addEventListener("input", (event) => {
  if (event.target instanceof HTMLInputElement && event.target.closest("#hierarchyGridBody")) {
    const row = Number(event.target.dataset.row);
    const tbody = document.getElementById("hierarchyGridBody");
    if (!Number.isNaN(row) && row >= tbody.children.length - 1) {
      adicionarLinhasGrade(20);
    }
    ajustarLarguraColunasGrade();
    salvarEstadoNoHistorico();
  }
});
document.addEventListener("focusin", (event) => {
  if (event.target instanceof HTMLInputElement && event.target.closest("#hierarchyGridBody")) {
    celulaAtiva = event.target;
  }
});

document.addEventListener("mousedown", (event) => {
  const alvo = event.target;
  if (!(alvo instanceof HTMLInputElement) || !alvo.closest("#hierarchyGridBody")) return;
  selecionando = true;
  const row = Number(alvo.dataset.row);
  const col = Number(alvo.dataset.col);
  ancoraSelecao = { row, col };
  aplicarSelecaoRetangular(ancoraSelecao, ancoraSelecao);
});

document.addEventListener("mouseover", (event) => {
  if (!selecionando || !ancoraSelecao) return;
  const alvo = event.target;
  if (!(alvo instanceof HTMLInputElement) || !alvo.closest("#hierarchyGridBody")) return;
  const row = Number(alvo.dataset.row);
  const col = Number(alvo.dataset.col);
  aplicarSelecaoRetangular(ancoraSelecao, { row, col });
});

document.addEventListener("mouseup", () => {
  selecionando = false;
});
document.addEventListener("keydown", (event) => {
  const tecla = event.key.toLowerCase();
  const alvoNaGrade = event.target instanceof HTMLInputElement && event.target.closest("#hierarchyGridBody");
  const ehCtrl = event.ctrlKey || event.metaKey;

  // Só limpa o bloco selecionado quando o foco NÃO está digitando numa célula.
  if ((tecla === "delete" || tecla === "backspace") && celulasSelecionadas.size && !alvoNaGrade) {
    event.preventDefault();
    salvarEstadoNoHistorico();
    for (const chave of celulasSelecionadas) {
      const { row, col } = parseKeyCelula(chave);
      const input = getInputPorRowCol(row, col);
      if (input) input.value = "";
    }
    ajustarLarguraColunasGrade();
    salvarEstadoNoHistorico();
    return;
  }

  if (!ehCtrl) return;

  if (tecla === "z") {
    // Desfaz globalmente, sem exigir foco na celula.
    event.preventDefault();
    desfazerGrade();
    ajustarLarguraColunasGrade();
    return;
  }

  if (tecla === "c" && alvoNaGrade) {
    let valor = "";
    if (celulasSelecionadas.size) {
      valor = textoDaSelecao();
    } else if (celulaAtiva) {
      const textoSelecionado = celulaAtiva.value.substring(celulaAtiva.selectionStart ?? 0, celulaAtiva.selectionEnd ?? 0);
      valor = textoSelecionado || celulaAtiva.value;
    }
    if (!valor) return;
    event.preventDefault();
    navigator.clipboard.writeText(valor).catch(() => {});
  }

  if (tecla === "x" && alvoNaGrade) {
    let valor = "";
    if (celulasSelecionadas.size) {
      valor = textoDaSelecao();
      salvarEstadoNoHistorico();
      for (const chave of celulasSelecionadas) {
        const { row, col } = parseKeyCelula(chave);
        const input = getInputPorRowCol(row, col);
        if (input) input.value = "";
      }
      ajustarLarguraColunasGrade();
      salvarEstadoNoHistorico();
    } else if (celulaAtiva) {
      valor = celulaAtiva.value;
      salvarEstadoNoHistorico();
      celulaAtiva.value = "";
      salvarEstadoNoHistorico();
    }
    if (!valor) return;
    event.preventDefault();
    navigator.clipboard.writeText(valor).catch(() => {});
  }
});

salvarEstadoNoHistorico();
aplicarTema(localStorage.getItem(chaveTema) === "dark" ? "dark" : "light");
document.body.classList.remove("compact");
localStorage.removeItem("manipulador-hierarquia-compacto");
aplicarIdioma(idiomaAtual);
renderHeader();
render();
