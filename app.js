const DAILY_GOAL = 200;
const STORAGE_KEY = "cycleCountProduction.snapshots.v2";
const AISLES = "ABCDEFGHIJKL".split("");

const PEOPLE = [
  { name: "Carico", initials: "CH", aisles: ["A", "B"], production: true },
  { name: "Ernie", initials: "EH", aisles: ["C", "D"], production: true },
  { name: "Cherish", initials: "CC", aisles: ["E", "F"], production: true },
  { name: "Layne", initials: "LM", aisles: ["G", "H"], production: true },
  { name: "Madison", initials: "MJ", aisles: ["I", "J"], production: true },
  { name: "Antoine", initials: "AH", aisles: ["K", "L"], production: true },
  { name: "Greg", initials: "GR", aisles: [], production: false },
  { name: "Denise", initials: "DW", aisles: [], production: false },
];

const productionPeople = PEOPLE.filter((person) => person.production);
const initialsMap = Object.fromEntries(PEOPLE.map((person) => [person.initials, person]));

const state = {
  detailWorkbook: null,
  detailRows: [],
  createdByAisle: Object.fromEntries(AISLES.map((aisle) => [aisle, 0])),
  initialsByPerson: Object.fromEntries(PEOPLE.map((person) => [person.name, 0])),
  totalsByPerson: Object.fromEntries(PEOPLE.map((person) => [person.name, 0])),
  pastedRows: [],
  reviewRows: [],
};

const $ = (id) => document.getElementById(id);
const sourceFile = $("sourceFile");
const sourceSheet = $("sourceSheet");
const batchColumn = $("batchColumn");

function normalize(value) {
  return String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9%#]+/g, " ").trim();
}

function readWorkbook(file) {
  return file.arrayBuffer().then((buffer) => XLSX.read(buffer, { type: "array", cellDates: true, cellStyles: true }));
}

function matrixFor(workbook, sheetName) {
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", blankrows: false, raw: false });
}

function setOptions(select, options, selectedValue = "") {
  select.innerHTML = "";
  options.forEach(({ value, label }) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label;
    option.selected = String(value) === String(selectedValue);
    select.appendChild(option);
  });
}

function findHeaderRow(matrix, keywords, maxRows = 25) {
  let bestIndex = 0;
  let bestScore = -1;
  matrix.slice(0, maxRows).forEach((row, index) => {
    const cells = row.map(normalize);
    const score = keywords.reduce((sum, keyword) => sum + (cells.some((cell) => cell.includes(keyword)) ? 1 : 0), 0);
    if (score > bestScore) { bestIndex = index; bestScore = score; }
  });
  return bestIndex;
}

function detectColumn(headers, candidates) {
  const normalized = headers.map(normalize);
  const exact = normalized.findIndex((header) => candidates.includes(header));
  if (exact >= 0) return exact;
  return normalized.findIndex((header) => candidates.some((candidate) => header.includes(candidate)));
}

function parseCreatedBatch(value) {
  const text = String(value ?? "").trim().toUpperCase();
  const match = text.match(/^([A-L])-\d{2}-\d{2}$/);
  return match ? match[1] : null;
}

function loadDetailSheet() {
  const matrix = matrixFor(state.detailWorkbook, sourceSheet.value);
  const headerIndex = findHeaderRow(matrix, ["batch", "bin", "count date"]);
  const headers = matrix[headerIndex] || [];
  const options = headers.map((header, index) => ({ value: index, label: `${XLSX.utils.encode_col(index)} — ${header || "(blank header)"}` }));
  let detected = detectColumn(headers, ["batch"]);
  if (detected < 0) detected = detectColumn(headers, ["bin #", "bin"]);
  if (detected < 0) detected = 0;
  setOptions(batchColumn, options, detected);
  state.detailRows = matrix.slice(headerIndex + 1);
  calculateCreatedCounts();
}

function calculateCreatedCounts() {
  state.createdByAisle = Object.fromEntries(AISLES.map((aisle) => [aisle, 0]));
  if (state.detailRows.length && batchColumn.value !== "") {
    const column = Number(batchColumn.value);
    state.detailRows.forEach((row) => {
      const aisle = parseCreatedBatch(row[column]);
      if (aisle) state.createdByAisle[aisle] += 1;
    });
  }
  calculateTotals();
}

function parsePastedRows(text) {
  return text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean).map((line, index) => {
    const parts = line.includes("\t") ? line.split("\t") : line.split(/[,;]|\s{2,}/);
    const cleaned = parts.map((part) => String(part).trim()).filter(Boolean);
    return { row: index + 1, item: cleaned[0] || "", initials: (cleaned[1] || "").toUpperCase() };
  }).filter((row) => !normalize(row.item).includes("item number") && !normalize(row.initials).includes("initial"));
}

function readPastedList() {
  state.pastedRows = parsePastedRows($("alreadyPaste").value);
  state.initialsByPerson = Object.fromEntries(PEOPLE.map((person) => [person.name, 0]));
  state.reviewRows = [];
  state.pastedRows.forEach((row) => {
    if (!row.item && !row.initials) return;
    const person = initialsMap[row.initials];
    if (person) state.initialsByPerson[person.name] += 1;
    else state.reviewRows.push({ row: row.row, item: row.item || "(blank)", initials: row.initials || "Blank" });
  });
  $("alreadyStatus").textContent = `${state.pastedRows.length} pasted rows read`;
  $("alreadyStatus").classList.add("success");
  calculateTotals();
}

function calculateTotals() {
  const totals = Object.fromEntries(PEOPLE.map((person) => [person.name, state.initialsByPerson[person.name] || 0]));
  productionPeople.forEach((person) => {
    totals[person.name] += person.aisles.reduce((sum, aisle) => sum + state.createdByAisle[aisle], 0);
  });
  state.totalsByPerson = totals;
  render();
}

function render() {
  if (!state.detailWorkbook && !state.pastedRows.length) return;
  const createdTotal = Object.values(state.createdByAisle).reduce((sum, count) => sum + count, 0);
  const pastedTotal = Object.values(state.initialsByPerson).reduce((sum, count) => sum + count, 0);
  const teamTotal = productionPeople.reduce((sum, person) => sum + state.totalsByPerson[person.name], 0);
  const reviewTotal = state.reviewRows.length;

  $("resultsSection").classList.remove("hidden");
  $("recordSummary").textContent = `${teamTotal} team counts • ${reviewTotal} review`;
  $("kpiStrip").innerHTML = `<div class="kpi"><span>Created counts</span><strong>${createdTotal}</strong></div><div class="kpi"><span>Initials-list counts</span><strong>${pastedTotal}</strong></div><div class="kpi"><span>Production-team total</span><strong>${teamTotal}</strong></div><div class="kpi"><span>Needs review</span><strong>${reviewTotal}</strong></div>`;

  $("productionCards").innerHTML = productionPeople.map((person) => {
    const created = person.aisles.reduce((sum, aisle) => sum + state.createdByAisle[aisle], 0);
    const initials = state.initialsByPerson[person.name] || 0;
    const total = state.totalsByPerson[person.name] || 0;
    const percent = total / DAILY_GOAL * 100;
    return `<article class="summary-card"><div class="summary-card-top"><div><strong>${person.name}</strong><span>Created: ${created} • Initials ${person.initials}: ${initials}</span></div><b>${total}</b></div><div class="meter"><span style="width:${Math.min(percent,100)}%"></span></div><div class="percent-row"><span>${percent.toFixed(1)}%</span><small>${total - DAILY_GOAL >= 0 ? "+" : ""}${total - DAILY_GOAL} vs goal</small></div></article>`;
  }).join("");

  $("otherCards").innerHTML = PEOPLE.filter((person) => !person.production).map((person) => `<article class="summary-card"><div class="summary-card-top"><div><strong>${person.name}</strong><span>${person.initials} • tracked separately</span></div><b>${state.initialsByPerson[person.name] || 0}</b></div></article>`).join("");

  const details = $("reviewDetails");
  if (reviewTotal) {
    details.classList.remove("hidden");
    $("reviewCount").textContent = reviewTotal;
    $("reviewList").innerHTML = state.reviewRows.map((row) => `<div><span>Row ${row.row}</span><strong>${escapeHtml(row.item)}</strong><small>${escapeHtml(row.initials)}</small></div>`).join("");
  } else details.classList.add("hidden");
}

function escapeHtml(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#039;");
}

function loadHistory() {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]"); } catch { return []; }
}

function saveHistory(history) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(history));
}

function currentRecord() {
  return {
    id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now()),
    savedAt: new Date().toISOString(),
    createdByAisle: { ...state.createdByAisle },
    initialsByPerson: { ...state.initialsByPerson },
    totalsByPerson: { ...state.totalsByPerson },
    reviewCount: state.reviewRows.length,
  };
}

function saveSnapshot() {
  const history = loadHistory();
  const record = currentRecord();
  history.unshift(record);
  saveHistory(history.slice(0, 100));
  showSaveMessage(`Results saved at ${new Date(record.savedAt).toLocaleString()}.`);
  renderHistory();
}

function renderHistory() {
  const records = loadHistory();
  $("historyBody").innerHTML = records.length ? records.map((record) => {
    const team = productionPeople.reduce((sum, person) => sum + (record.totalsByPerson?.[person.name] || 0), 0);
    return `<tr><td>${new Date(record.savedAt).toLocaleString()}</td><td>${team}</td><td>${record.totalsByPerson?.Greg || 0}</td><td>${record.totalsByPerson?.Denise || 0}</td><td>${record.reviewCount || 0}</td><td><button data-delete-id="${record.id}" type="button">Delete</button></td></tr>`;
  }).join("") : `<tr><td colspan="6" class="empty-row">No saved results yet.</td></tr>`;
}

function showSaveMessage(text, isError = false) {
  const message = $("saveMessage");
  message.textContent = text;
  message.classList.remove("hidden", "error");
  if (isError) message.classList.add("error");
}

function downloadSummary() {
  const record = currentRecord();
  const employeeRows = PEOPLE.map((person) => ({ Employee: person.name, Initials: person.initials, "Created Counts": person.production ? person.aisles.reduce((sum, aisle) => sum + record.createdByAisle[aisle], 0) : 0, "Already Counted List": record.initialsByPerson[person.name] || 0, "Total Counts": record.totalsByPerson[person.name] || 0, "Production %": person.production ? (record.totalsByPerson[person.name] || 0) / DAILY_GOAL : "" }));
  const workbook = XLSX.utils.book_new();
  const summary = XLSX.utils.json_to_sheet(employeeRows);
  employeeRows.forEach((row, index) => { if (typeof row["Production %"] === "number") summary[`F${index + 2}`].z = "0.0%"; });
  XLSX.utils.book_append_sheet(workbook, summary, "Production Summary");
  if (state.reviewRows.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(state.reviewRows), "Needs Review");
  const stamp = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 16);
  XLSX.writeFile(workbook, `Cycle_Count_Production_${stamp}.xlsx`);
}

renderHistory();
sourceFile.addEventListener("change", async () => {
  const file = sourceFile.files[0];
  if (!file) return;
  try {
    state.detailWorkbook = await readWorkbook(file);
    setOptions(sourceSheet, state.detailWorkbook.SheetNames.map((name) => ({ value: name, label: name })));
    loadDetailSheet();
    $("sourceControls").classList.remove("hidden");
    $("sourceStatus").textContent = `${file.name} loaded`;
    $("sourceStatus").classList.add("success");
  } catch {
    $("sourceStatus").textContent = "Could not read file";
  }
});
sourceSheet.addEventListener("change", loadDetailSheet);
batchColumn.addEventListener("change", calculateCreatedCounts);
$("readPasteBtn").addEventListener("click", readPastedList);
$("clearPasteBtn").addEventListener("click", () => { $("alreadyPaste").value = ""; state.pastedRows = []; state.initialsByPerson = Object.fromEntries(PEOPLE.map((person) => [person.name, 0])); state.reviewRows = []; $("alreadyStatus").textContent = "Waiting for data"; calculateTotals(); });
$("saveSnapshotBtn").addEventListener("click", saveSnapshot);
$("downloadSummaryBtn").addEventListener("click", downloadSummary);
$("historyBody").addEventListener("click", (event) => { const button = event.target.closest("button[data-delete-id]"); if (!button) return; saveHistory(loadHistory().filter((record) => record.id !== button.dataset.deleteId)); renderHistory(); });