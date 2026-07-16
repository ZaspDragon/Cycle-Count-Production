const DAILY_GOAL = 200;
const STORAGE_KEY = "cycleCountProduction.snapshots.v3";
const AISLES = "ABCDEFGHIJKL".split("");

const ASSIGNMENTS = [
  { name: "Carico", aisles: ["A", "B"] },
  { name: "Ernie", aisles: ["C", "D"] },
  { name: "Cherish", aisles: ["E", "F"] },
  { name: "Layne", aisles: ["G", "H"] },
  { name: "Madison", aisles: ["I", "J"] },
  { name: "Antoine", aisles: ["K", "L"] },
];

const state = {
  workbook: null,
  rows: [],
  headerIndex: 0,
  aisleTotals: Object.fromEntries(AISLES.map((aisle) => [aisle, 0])),
  employeeTotals: Object.fromEntries(ASSIGNMENTS.map(({ name }) => [name, 0])),
  uncreditedRows: [],
};

const $ = (id) => document.getElementById(id);
const sourceFile = $("sourceFile");
const sourceSheet = $("sourceSheet");
const batchColumn = $("batchColumn");

function normalize(value) {
  return String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9#%]+/g, " ").trim();
}

function readWorkbook(file) {
  return file.arrayBuffer().then((buffer) => XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    cellStyles: true,
  }));
}

function matrixFor(workbook, sheetName) {
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
    blankrows: false,
    raw: false,
  });
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
    const score = keywords.reduce(
      (sum, keyword) => sum + (cells.some((cell) => cell.includes(keyword)) ? 1 : 0),
      0
    );
    if (score > bestScore) {
      bestIndex = index;
      bestScore = score;
    }
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

function loadSelectedSheet() {
  const matrix = matrixFor(state.workbook, sourceSheet.value);
  const headerIndex = findHeaderRow(matrix, ["batch", "bin", "count date"]);
  const headers = matrix[headerIndex] || [];
  const options = headers.map((header, index) => ({
    value: index,
    label: `${XLSX.utils.encode_col(index)} — ${header || "(blank header)"}`,
  }));

  let detected = detectColumn(headers, ["batch"]);
  if (detected < 0) detected = detectColumn(headers, ["bin #", "bin"]);
  if (detected < 0) detected = 0;

  setOptions(batchColumn, options, detected);
  state.headerIndex = headerIndex;
  state.rows = matrix.slice(headerIndex + 1);
  calculateProduction();
}

function calculateProduction() {
  state.aisleTotals = Object.fromEntries(AISLES.map((aisle) => [aisle, 0]));
  state.uncreditedRows = [];

  if (state.rows.length && batchColumn.value !== "") {
    const column = Number(batchColumn.value);
    state.rows.forEach((row, offset) => {
      const batch = String(row[column] ?? "").trim();
      if (!batch) return;
      const aisle = parseCreatedBatch(batch);
      if (aisle) {
        state.aisleTotals[aisle] += 1;
      } else {
        state.uncreditedRows.push({
          row: state.headerIndex + offset + 2,
          batch,
        });
      }
    });
  }

  state.employeeTotals = Object.fromEntries(
    ASSIGNMENTS.map(({ name, aisles }) => [
      name,
      aisles.reduce((sum, aisle) => sum + state.aisleTotals[aisle], 0),
    ])
  );

  renderResults();
}

function renderResults() {
  if (!state.workbook) return;

  const credited = Object.values(state.aisleTotals).reduce((sum, count) => sum + count, 0);
  const uncredited = state.uncreditedRows.length;

  $("resultsSection").classList.remove("hidden");
  $("recordSummary").textContent = `${credited} credited • ${uncredited} uncredited`;
  $("kpiStrip").innerHTML = `
    <div class="kpi"><span>Credited created counts</span><strong>${credited}</strong></div>
    <div class="kpi"><span>Uncredited rows</span><strong>${uncredited}</strong></div>
    <div class="kpi"><span>Daily team goal</span><strong>${DAILY_GOAL * ASSIGNMENTS.length}</strong></div>
    <div class="kpi"><span>Team production</span><strong>${((credited / (DAILY_GOAL * ASSIGNMENTS.length)) * 100).toFixed(1)}%</strong></div>`;

  $("productionCards").innerHTML = ASSIGNMENTS.map(({ name, aisles }) => {
    const total = state.employeeTotals[name];
    const percent = (total / DAILY_GOAL) * 100;
    const breakdown = aisles.map((aisle) => `${aisle}: ${state.aisleTotals[aisle]}`).join(" • ");
    return `<article class="summary-card">
      <div class="summary-card-top">
        <div><strong>${name}</strong><span>${breakdown}</span></div>
        <b>${total}</b>
      </div>
      <div class="meter"><span style="width:${Math.min(percent, 100)}%"></span></div>
      <div class="percent-row"><span>${percent.toFixed(1)}%</span><small>${total - DAILY_GOAL >= 0 ? "+" : ""}${total - DAILY_GOAL} vs goal</small></div>
    </article>`;
  }).join("");

  const details = $("reviewDetails");
  if (uncredited) {
    details.classList.remove("hidden");
    $("reviewCount").textContent = uncredited;
    $("reviewList").innerHTML = state.uncreditedRows
      .slice(0, 300)
      .map((item) => `<div><span>Row ${item.row}</span><strong>${escapeHtml(item.batch)}</strong><small>Not A-00-00</small></div>`)
      .join("");
  } else {
    details.classList.add("hidden");
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function loadHistory() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
  } catch {
    return [];
  }
}

function saveHistory(history) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(history));
}

function currentRecord() {
  return {
    id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now()),
    savedAt: new Date().toISOString(),
    aisleTotals: { ...state.aisleTotals },
    employeeTotals: { ...state.employeeTotals },
    uncreditedCount: state.uncreditedRows.length,
  };
}

function saveSnapshot() {
  const history = loadHistory();
  const record = currentRecord();
  history.unshift(record);
  saveHistory(history.slice(0, 100));
  showMessage(`Results saved at ${new Date(record.savedAt).toLocaleString()}.`);
  renderHistory();
}

function renderHistory() {
  const records = loadHistory();
  $("historyBody").innerHTML = records.length
    ? records.map((record) => {
        const teamTotal = Object.values(record.employeeTotals || {}).reduce((sum, count) => sum + count, 0);
        return `<tr>
          <td>${new Date(record.savedAt).toLocaleString()}</td>
          <td>${teamTotal}</td>
          <td>${record.uncreditedCount || 0}</td>
          <td><button data-delete-id="${record.id}" type="button">Delete</button></td>
        </tr>`;
      }).join("")
    : `<tr><td colspan="4" class="empty-row">No saved results yet.</td></tr>`;
}

function showMessage(text, isError = false) {
  const message = $("saveMessage");
  message.textContent = text;
  message.classList.remove("hidden", "error");
  if (isError) message.classList.add("error");
}

function downloadSummary() {
  const employeeRows = ASSIGNMENTS.map(({ name, aisles }) => ({
    Employee: name,
    "Assigned Aisles": aisles.join("-"),
    "Aisle Breakdown": aisles.map((aisle) => `${aisle}: ${state.aisleTotals[aisle]}`).join(" | "),
    "Cycle Counts": state.employeeTotals[name],
    "Production %": state.employeeTotals[name] / DAILY_GOAL,
    "Daily Goal": DAILY_GOAL,
  }));

  const aisleRows = AISLES.map((aisle) => ({
    Aisle: aisle,
    "Cycle Counts": state.aisleTotals[aisle],
  }));

  const workbook = XLSX.utils.book_new();
  const productionSheet = XLSX.utils.json_to_sheet(employeeRows);
  employeeRows.forEach((row, index) => {
    productionSheet[`E${index + 2}`].z = "0.0%";
  });
  productionSheet["!cols"] = [{ wch: 16 }, { wch: 18 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 12 }];

  XLSX.utils.book_append_sheet(workbook, productionSheet, "Employee Production");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(aisleRows), "Aisle Totals");

  if (state.uncreditedRows.length) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(state.uncreditedRows), "Uncredited Rows");
  }

  const stamp = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 16);
  XLSX.writeFile(workbook, `Cycle_Count_Production_${stamp}.xlsx`);
}

renderHistory();

sourceFile.addEventListener("change", async () => {
  const file = sourceFile.files[0];
  if (!file) return;
  try {
    state.workbook = await readWorkbook(file);
    setOptions(sourceSheet, state.workbook.SheetNames.map((name) => ({ value: name, label: name })));
    loadSelectedSheet();
    $("sourceControls").classList.remove("hidden");
    $("sourceStatus").textContent = `${file.name} loaded`;
    $("sourceStatus").classList.add("success");
  } catch (error) {
    console.error(error);
    $("sourceStatus").textContent = "Could not read file";
  }
});

sourceSheet.addEventListener("change", loadSelectedSheet);
batchColumn.addEventListener("change", calculateProduction);
$("saveSnapshotBtn").addEventListener("click", saveSnapshot);
$("downloadSummaryBtn").addEventListener("click", downloadSummary);
$("historyBody").addEventListener("click", (event) => {
  const button = event.target.closest("button[data-delete-id]");
  if (!button) return;
  saveHistory(loadHistory().filter((record) => record.id !== button.dataset.deleteId));
  renderHistory();
});
