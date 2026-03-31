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
  thMes.textContent = "Mes";
  monthRow.appendChild(thMes);

  const thIndicador = document.createElement("th");
  thIndicador.textContent = "Indicador";
  headerRow.appendChild(thIndicador);

  const nomesMeses = [
    "Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
  ];

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
  if (!(alvo instanceof HTMLInputElement) || !alvo.closest("#hierarchyGridBody")) return;

  const texto = event.clipboardData?.getData("text");
  if (!texto) return;

  salvarEstadoNoHistorico();
  event.preventDefault();
  const inicioRow = Number(alvo.dataset.row);
  const inicioCol = Number(alvo.dataset.col);
  const linhas = texto.replace(/\r/g, "").split("\n").filter((l) => l.length > 0);
  garantirQuantidadeLinhas(inicioRow + linhas.length + 20);

  for (let i = 0; i < linhas.length; i++) {
    const colunas = linhas[i].split("\t");
    for (let j = 0; j < colunas.length; j++) {
      const r = inicioRow + i;
      const c = inicioCol + j;
      if (c > 2) continue;
      const input = document.querySelector(`#hierarchyGridBody input[data-row="${r}"][data-col="${c}"]`);
      if (input) input.value = colunas[j].trim();
    }
  }
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
  document.getElementById("nodeCount").textContent = `${rows.length} itens`;

  for (const row of rows) {
    if (!row.visible) continue;

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

function collectParentNodes(nodes, set) {
  for (const n of nodes) {
    if (n.children && n.children.length) {
      set.add(n.id);
      collectParentNodes(n.children, set);
    }
  }
}

document.getElementById("expandAll").addEventListener("click", () => {
  openNodes.clear();
  collectParentNodes(hierarchy, openNodes);
  render();
});

document.getElementById("collapseAll").addEventListener("click", () => {
  openNodes.clear();
  render();
});

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
    itensDaHierarquia = novosItens;
    hierarchy = buildHierarchy(itensDaHierarquia);
    openNodes.clear();
    for (const item of itensDaHierarquia) {
      if (!item.parentId) openNodes.add(item.id);
    }
    render();
    ajustarLarguraColunasGrade();
    excelWrapper.classList.remove("input-error");
    status.textContent = `Hierarquia aplicada com sucesso (${novosItens.length} linhas).`;
    mostrarPopupSucesso("Hierarquia valida e aplicada com sucesso.");
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
  document.getElementById("importPanel").open = true;
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
  if (event.key === "Escape" && replaceModal.classList.contains("show")) {
    fecharModalSubstituicao();
  }
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
document.addEventListener("keydown", (event) => {
  const ehCtrl = event.ctrlKey || event.metaKey;
  if (!ehCtrl) return;

  const tecla = event.key.toLowerCase();
  const alvoNaGrade = event.target instanceof HTMLInputElement && event.target.closest("#hierarchyGridBody");

  if (tecla === "z") {
    // Desfaz globalmente, sem exigir foco na celula.
    event.preventDefault();
    desfazerGrade();
    ajustarLarguraColunasGrade();
    return;
  }

  if (tecla === "c" && alvoNaGrade) {
    if (!celulaAtiva) return;
    const textoSelecionado = celulaAtiva.value.substring(celulaAtiva.selectionStart ?? 0, celulaAtiva.selectionEnd ?? 0);
    const valor = textoSelecionado || celulaAtiva.value;
    event.preventDefault();
    navigator.clipboard.writeText(valor).catch(() => {});
  }
});

salvarEstadoNoHistorico();
aplicarTema(localStorage.getItem(chaveTema) === "dark" ? "dark" : "light");
renderHeader();
render();
