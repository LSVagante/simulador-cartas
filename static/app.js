const selectedCodes = new Set();
const selectedProposalIds = new Set();
const MAX_RENDER_ROWS = 5000;

let currentRows = [];
let currentSelectionMode = "cards";
let baseRows = [];
let baseMode = "cards";

const fields = [
  "q",
  "credito_min",
  "credito_max",
  "entrada_min",
  "entrada_max",
  "parcela_min",
  "parcela_max",
  "custo_min",
  "custo_max",
  "disponivel",
];

const calcFields = ["calc_credito", "categoria", "calc_entrada", "calc_parcela", "calc_saldo"];

const el = (id) => document.getElementById(id);

async function fetchJson(url, options = {}) {
  const response = await fetch(url, options);
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || "Erro na requisicao");
  }
  return data;
}

function safe(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;",
  })[char]);
}

function formatMoneyValue(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return "";
  return number.toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function formatPercentValue(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return "";
  return `${(number * 100).toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}%`;
}

function moneyText(formatted, raw) {
  return formatted ?? formatMoneyValue(raw);
}

function percentText(formatted, raw) {
  return formatted ?? formatPercentValue(raw);
}

function selectedAdministradoras() {
  return Array.from(document.querySelectorAll("#adminMenu input[type='checkbox']:checked"))
    .map((checkbox) => checkbox.value)
    .filter(Boolean);
}

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function clearSelection() {
  selectedCodes.clear();
  selectedProposalIds.clear();
}

function paramsFromFilters() {
  const params = new URLSearchParams();
  for (const field of fields) {
    const input = el(field);
    if (!input) continue;
    if (input.type === "checkbox") {
      params.set(field, input.checked ? "1" : "0");
      continue;
    }
    const value = input.value.trim();
    if (value) params.set(field, value);
  }
  params.set("limit", "800");
  return params;
}

async function loadStatus() {
  const status = await fetchJson("/api/status");
  const source = status.source || status.workbook || "base local";
  const updated = status.updatedAt ? `, atualizado em ${status.updatedAt}` : "";
  el("workbookInfo").textContent = `${status.totalCartas} cartas carregadas de ${source}${updated}`;
  fillSelect("categoria", status.categorias, "Categoria");
  fillAdminDropdown(status.administradoras);
}

function fillSelect(id, values, emptyLabel) {
  const select = el(id);
  const current = select.value;
  select.innerHTML = `<option value="">${safe(emptyLabel)}</option>`;
  for (const value of values) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  }
  select.value = current;
}

function fillAdminDropdown(values) {
  const selected = new Set(selectedAdministradoras());
  const menu = el("adminMenu");
  menu.innerHTML = "";

  for (const value of values) {
    const item = document.createElement("label");
    item.className = "admin-option";
    item.innerHTML = `
      <input type="checkbox" value="${safe(value)}" ${selected.has(value) ? "checked" : ""}>
      <span>${safe(value)}</span>
    `;
    menu.appendChild(item);
  }

  menu.querySelectorAll("input[type='checkbox']").forEach((checkbox) => {
    checkbox.addEventListener("change", updateAdminSummary);
  });
  updateAdminSummary();
}

function updateAdminSummary() {
  const selected = selectedAdministradoras();
  if (!selected.length) {
    el("adminSummary").textContent = "Todas";
  } else if (selected.length === 1) {
    el("adminSummary").textContent = selected[0];
  } else {
    el("adminSummary").textContent = `${selected.length} selecionadas`;
  }
}

async function loadCards() {
  clearSelection();
  const data = await fetchJson("/api/cartas?disponivel=0&limit=5000");
  baseRows = data.items;
  baseMode = "cards";
  el("calcInfo").textContent = "";
  applyVisualFilters();
}

async function calculateProposals() {
  clearSelection();
  el("calcInfo").textContent = "Calculando...";
  const payload = {
    creditoAlvo: el("calc_credito").value.trim(),
    categoria: el("categoria").value,
    entradaMax: el("calc_entrada").value.trim(),
    parcelaMax: el("calc_parcela").value.trim(),
    saldo: el("calc_saldo").value,
    administradoras: selectedAdministradoras(),
    limit: 20000,
  };

  try {
    const data = await fetchJson("/api/calcular", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    baseRows = data.items;
    baseMode = "proposals";
    applyVisualFilters();
    el("calcInfo").textContent = data.truncated
      ? `${data.totalPropostas} propostas geradas, mostrando ${data.items.length} de ${data.totalLinhas} linhas`
      : `${data.totalPropostas} propostas geradas, ${data.totalLinhas} linhas`;
  } catch (error) {
    el("calcInfo").textContent = error.message;
  }
}

function applyVisualFilters() {
  clearSelection();
  const rows = filterVisualRows(baseRows);
  const visibleRows = rows.slice(0, MAX_RENDER_ROWS);
  renderRows(visibleRows);
  updateCountInfo(rows, visibleRows);
  updateSelectedInfo();
}

function updateCountInfo(rows, visibleRows = rows) {
  const limited = visibleRows.length < rows.length;

  if (baseMode === "proposals") {
    const filteredProposals = uniqueProposalCount(rows);
    const totalProposals = uniqueProposalCount(baseRows);
    if (limited) {
      const visibleProposals = uniqueProposalCount(visibleRows);
      el("countInfo").textContent = `${visibleProposals} propostas exibidas de ${filteredProposals} filtradas`;
      return;
    }
    el("countInfo").textContent = `${filteredProposals} de ${totalProposals} propostas`;
    return;
  }

  if (limited) {
    el("countInfo").textContent = `${visibleRows.length} cartas exibidas de ${rows.length} filtradas`;
    return;
  }

  el("countInfo").textContent = `${rows.length} de ${baseRows.length} cartas`;
}

function uniqueProposalCount(rows) {
  return new Set(rows.map((item) => item.proposta).filter(Boolean)).size;
}

function filterVisualRows(rows) {
  if (baseMode === "proposals") {
    const groups = new Map();
    for (const row of rows) {
      if (!groups.has(row.proposta)) groups.set(row.proposta, []);
      groups.get(row.proposta).push(row);
    }

    return Array.from(groups.values())
      .filter((group) => groupMatchesVisualFilters(group))
      .flat();
  }

  return rows.filter((row) => groupMatchesVisualFilters([row]));
}

function groupMatchesVisualFilters(group) {
  const q = normalizeText(el("q").value);
  const selectedAdmins = new Set(selectedAdministradoras().map(normalizeText));
  const onlyAvailable = el("disponivel").checked;

  if (baseMode !== "proposals" && onlyAvailable && group.some((item) => normalizeText(item.status) !== "disponivel")) {
    return false;
  }

  if (selectedAdmins.size && !group.some((item) => selectedAdmins.has(normalizeText(item.administradora)))) {
    return false;
  }

  if (q && !group.some((item) => rowSearchText(item).includes(q))) {
    return false;
  }

  const reference = group[0];
  return (
    inRange(metricValue(reference, "credito"), readMoneyFilter("credito_min"), readMoneyFilter("credito_max")) &&
    inRange(metricValue(reference, "entrada"), readMoneyFilter("entrada_min"), readMoneyFilter("entrada_max")) &&
    inRange(metricValue(reference, "parcela"), readMoneyFilter("parcela_min"), readMoneyFilter("parcela_max")) &&
    (baseMode !== "proposals" || inRange(metricValue(reference, "custo"), readPercentFilter("custo_min"), readPercentFilter("custo_max")))
  );
}

function rowSearchText(item) {
  return normalizeText([
    item.proposta,
    item.codigo,
    item.categoria,
    item.administradora,
  ].join(" "));
}

function metricValue(item, metric) {
  if (metric === "credito") return numberOrNull(item.creditoTotal ?? item.credito);
  if (metric === "entrada") return numberOrNull(item.entradaTotal ?? item.entrada);
  if (metric === "parcela") return numberOrNull(item.parcelaTotal ?? item.valorParcela);
  if (metric === "custo") return numberOrNull(item.custoPercent);
  return null;
}

function numberOrNull(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function readMoneyFilter(id) {
  return parseLocaleNumber(el(id).value);
}

function readPercentFilter(id) {
  const value = parseLocaleNumber(el(id).value);
  if (value === null) return null;
  return value > 10 ? value / 100 : value;
}

function parseLocaleNumber(value) {
  const text = String(value ?? "")
    .replace("R$", "")
    .replace("%", "")
    .replace(/\s/g, "")
    .trim();

  if (!text) return null;

  const normalized = text.includes(",")
    ? text.replace(/\./g, "").replace(",", ".")
    : text.replace(/\./g, "");
  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function inRange(value, min, max) {
  if (value === null) return min === null && max === null;
  if (min !== null && value < min) return false;
  if (max !== null && value > max) return false;
  return true;
}

function renderRows(rows) {
  currentRows = rows;
  currentSelectionMode = rows.some((item) => item.proposta) ? "proposals" : "cards";

  const body = el("cardsBody");
  const seenProposals = new Set();

  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="15" class="empty">Nenhum resultado encontrado.</td></tr>';
    return;
  }

  const htmlRows = [];
  for (const item of rows) {
    const proposalId = item.proposta || "";
    const isProposalRow = Boolean(proposalId);
    const isFirstProposalRow = !isProposalRow || !seenProposals.has(proposalId);

    if (isProposalRow && isFirstProposalRow) {
      seenProposals.add(proposalId);
    }

    const checkboxCell = buildCheckboxCell(item, isProposalRow, isFirstProposalRow);
    const proposalCell = isProposalRow ? (isFirstProposalRow ? proposalId : "") : "";
    const totals = item.formattedTotals || {};
    const rowClass = isProposalRow && !isFirstProposalRow ? ' class="proposal-child-row"' : "";
    const custoText = isProposalRow && !isFirstProposalRow
      ? ""
      : percentText(totals.custoPercent || item.formatted?.custoPercent, item.custoPercent);

    htmlRows.push(`
      <tr${rowClass}>
      <td>${checkboxCell}</td>
      <td>${safe(proposalCell)}</td>
      <td>${safe(item.codigo)}</td>
      <td>${safe(item.categoria)}</td>
      <td>${safe(item.administradora)}</td>
      <td class="num">${safe(moneyText(item.formatted?.credito, item.credito))}</td>
      <td class="num">${safe(moneyText(item.formatted?.entrada, item.entrada))}</td>
      <td class="num">${safe(item.numParcelas)}</td>
      <td class="num">${safe(moneyText(item.formatted?.valorParcela, item.valorParcela))}</td>
      <td class="num">${safe(moneyText(item.formatted?.saldoDevedor, item.saldoDevedor))}</td>
      <td class="num">${safe(moneyText(totals.creditoTotal, item.creditoTotal))}</td>
      <td class="num">${safe(moneyText(totals.entradaTotal, item.entradaTotal))}</td>
      <td class="num">${safe(moneyText(totals.parcelaTotal, item.parcelaTotal))}</td>
      <td class="num">${safe(moneyText(totals.saldoTotal, item.saldoTotal))}</td>
      <td class="num">${safe(custoText)}</td>
      </tr>
    `);
  }

  body.innerHTML = htmlRows.join("");
}

function buildCheckboxCell(item, isProposalRow, isFirstProposalRow) {
  if (isProposalRow) {
    if (!isFirstProposalRow) return "";
    const checked = selectedProposalIds.has(item.proposta) ? "checked" : "";
    return `<input type="checkbox" data-proposal="${safe(item.proposta)}" ${checked}>`;
  }

  const checked = selectedCodes.has(item.codigo) ? "checked" : "";
  return `<input type="checkbox" data-code="${safe(item.codigo)}" ${checked}>`;
}

function toggleProposalSelection(proposalId, checked) {
  clearSelection();
  if (checked) {
    selectedProposalIds.add(proposalId);
    currentRows
      .filter((item) => item.proposta === proposalId)
      .forEach((item) => selectedCodes.add(item.codigo));
  }
  updateSelectedInfo();
  renderRows(currentRows);
}

function toggleCardSelection(code, checked) {
  if (checked) selectedCodes.add(code);
  else selectedCodes.delete(code);
  updateSelectedInfo();
}

function updateSelectedInfo() {
  if (currentSelectionMode === "proposals") {
    const total = selectedProposalIds.size;
    el("selectedInfo").textContent = `${total} proposta${total === 1 ? "" : "s"} selecionada${total === 1 ? "" : "s"}`;
    return;
  }

  const total = selectedCodes.size;
  el("selectedInfo").textContent = `${total} carta${total === 1 ? "" : "s"} selecionada${total === 1 ? "" : "s"}`;
}

async function generateProposal() {
  if (!selectedCodes.size) {
    el("proposalText").textContent = "Selecione uma proposta ou uma ou mais cartas antes de gerar.";
    return;
  }

  const proposal = await fetchJson("/api/proposta", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ codigos: Array.from(selectedCodes) }),
  });
  el("proposalText").textContent = proposal.texto;
}

function clearFilters() {
  clearSelection();
  for (const field of fields) {
    const input = el(field);
    if (!input) continue;
    if (input.type === "checkbox") {
      input.checked = true;
    } else {
      input.value = "";
      if (input.classList.contains("money-input")) input.dataset.rawDigits = "";
    }
  }
  document.querySelectorAll("#adminMenu input[type='checkbox']").forEach((checkbox) => {
    checkbox.checked = false;
  });
  updateAdminSummary();
  applyVisualFilters();
}

async function reloadWorkbook() {
  await fetchJson("/api/reload", { method: "POST" });
  await loadStatus();
  await loadCards();
}

async function updateDataFromPlay() {
  const senha = window.prompt("Senha para atualizar os dados:");
  if (senha === null) return;

  const button = el("updateDataBtn");
  const previousText = button.textContent;
  button.disabled = true;
  button.textContent = "Atualizando...";
  el("workbookInfo").textContent = "Baixando dados da Play...";

  try {
    const status = await fetchJson("/api/update-play", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ senha }),
    });
    await loadStatus();
    await loadCards();
    el("calcInfo").textContent = `${status.totalCartas} cartas atualizadas da Play`;
  } catch (error) {
    el("workbookInfo").textContent = error.message;
  } finally {
    button.disabled = false;
    button.textContent = previousText;
  }
}

function formatMoneyFromDigits(digits) {
  if (!digits) return "";
  const value = Number(digits);
  return value.toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function setMoneyDigits(input, digits) {
  const cleaned = digits.replace(/\D/g, "").replace(/^0+(?=\d)/, "");
  input.dataset.rawDigits = cleaned;
  input.value = formatMoneyFromDigits(cleaned);
}

function setupMoneyInputs() {
  document.querySelectorAll(".money-input").forEach((input) => {
    input.dataset.rawDigits = input.value.replace(/\D/g, "");

    input.addEventListener("keydown", (event) => {
      if (event.ctrlKey || event.metaKey || event.altKey) return;
      if (["Tab", "Enter", "ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return;

      if (/^\d$/.test(event.key)) {
        event.preventDefault();
        setMoneyDigits(input, `${input.dataset.rawDigits || ""}${event.key}`);
        return;
      }

      if (event.key === "Backspace") {
        event.preventDefault();
        setMoneyDigits(input, (input.dataset.rawDigits || "").slice(0, -1));
        return;
      }

      if (event.key === "Delete" || event.key === "Escape") {
        event.preventDefault();
        setMoneyDigits(input, "");
        return;
      }

      event.preventDefault();
    });

    input.addEventListener("paste", (event) => {
      event.preventDefault();
      const text = event.clipboardData?.getData("text") || "";
      setMoneyDigits(input, text.replace(/\D/g, ""));
    });
  });
}

function setupAdminDropdown() {
  const dropdown = el("adminDropdown");
  const toggle = el("adminToggle");
  const menu = el("adminMenu");

  toggle.addEventListener("click", () => {
    dropdown.classList.toggle("open");
  });

  menu.addEventListener("click", (event) => {
    event.stopPropagation();
  });

  document.addEventListener("click", (event) => {
    if (!dropdown.contains(event.target)) {
      dropdown.classList.remove("open");
    }
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      dropdown.classList.remove("open");
    }
  });
}

el("applyBtn").addEventListener("click", applyVisualFilters);
el("clearBtn").addEventListener("click", clearFilters);
el("proposalBtn").addEventListener("click", generateProposal);
el("updateDataBtn").addEventListener("click", updateDataFromPlay);
el("reloadBtn").addEventListener("click", reloadWorkbook);
el("calculateBtn").addEventListener("click", calculateProposals);
el("cardsBody").addEventListener("change", (event) => {
  const checkbox = event.target;
  if (!checkbox || checkbox.type !== "checkbox") return;

  if (checkbox.dataset.proposal) {
    toggleProposalSelection(checkbox.dataset.proposal, checkbox.checked);
  } else if (checkbox.dataset.code) {
    toggleCardSelection(checkbox.dataset.code, checkbox.checked);
  }
});
el("copyBtn").addEventListener("click", async () => {
  await navigator.clipboard.writeText(el("proposalText").textContent);
});

for (const field of [...fields, ...calcFields]) {
  const input = el(field);
  if (!input) continue;
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      if (calcFields.includes(field)) calculateProposals();
      else applyVisualFilters();
    }
  });
}

setupMoneyInputs();
setupAdminDropdown();

loadStatus().then(loadCards).catch((error) => {
  el("workbookInfo").textContent = error.message;
});
