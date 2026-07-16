const DAILY_GOAL = 200;
const AISLES = "ABCDEFGHIJKL".split("");

const ASSIGNMENTS = [
  { name: "Carico", aisles: ["A", "B"] },
  { name: "Ernie", aisles: ["C", "D"] },
  { name: "Cherish", aisles: ["E", "F"] },
  { name: "Layne", aisles: ["G", "H"] },
  { name: "Madison", aisles: ["I", "J"] },
  { name: "Antoine", aisles: ["K", "L"] },
];

const aisleToEmployee = Object.fromEntries(
  ASSIGNMENTS.flatMap(({ name, aisles }) => aisles.map((aisle) => [aisle, name]))
);

const state = {
  sourceWorkbook: null,
  sourceRows: [],
  sourceHeaderIndex: 0,
  aisleTotals: null,
  employeeTotals: null,
  alreadyCountedRows: [],
  trackerWorkbook: null,
  trackerFileName: "",
};

const $ = (id) => document.getElementById(id);
const sourceFile = $("sourceFile");
const sourceSheet = $("sourceSheet");
const binColumn = $("binColumn");
const trackerFile = $("trackerFile");
const targetSheet = $("targetSheet");
const headerRow = $("headerRow");
const nameColumn = $("nameColumn");
const countColumn = $("countColumn");
const percentColumn = $("percentColumn");

function normalize(value) {
  return String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9%]+/g, " ").trim();
}

function readWorkbook(file) {
  return file.arrayBuffer().then((buffer) => XLSX.read(buffer, {
    type: "array",
    cellStyles: true,
    cellDates: true,
    bookVBA: true,
  }));
}

function worksheetToMatrix(workbook, sheetName) {
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
    if (String(value) === String(selectedValue)) option.selected = true;
    select.appendChild(option);
  });
}

function findLikelyHeaderRow(matrix, keywords, maxRows = 25) {
  let best = { index: 0, score: -1 };
  matrix.slice(0, maxRows).forEach((row, index) => {
    const cells = row.map(normalize);
    const score = keywords.reduce(
      (total, keyword) => total + (cells.some((cell) => cell.includes(keyword)) ? 1 : 0),
      0
    );
    if (score > best.score) best = { index, score };
  });
  return best.index;
}

function detectColumn(headers, candidates) {
  const normalized = headers.map(normalize);
  const exact = normalized.findIndex((header) => candidates.includes(header));
  if (exact >= 0) return exact;
  return normalized.findIndex((header) => candidates.some((candidate) => header.includes(candidate)));
}

function columnLabel(index) {
  return XLSX.utils.encode_col(index);
}

// Only a created-count location such as C-09-12 receives automatic aisle credit.
// Anything else is held for manual review in the Already Cycle Counted List.
function parseCreatedCountLocation(value) {
  const location = String(value ?? "").trim().toUpperCase();
  const match = location.match(/^([A-L])-([0-9]{2})-([0-9]{2})$/);
  return match ? { location, aisle: match[1] } : null;
}

function loadSourceSheet() {
  const matrix = worksheetToMatrix(state.sourceWorkbook, sourceSheet.value);
  const headerIndex = findLikelyHeaderRow(matrix, ["bin", "location", "aisle"]);
  const headers = matrix[headerIndex] || [];
  const options = headers.map((header, index) => ({
    value: index,
    label: `${columnLabel(index)} — ${header || "(blank header)"}`,
  }));

  let detected = detectColumn(headers, ["bin #", "bin", "bin number", "location", "location number"]);
  if (detected < 0) detected = 0;
  setOptions(binColumn, options, detected);

  state.sourceHeaderIndex = headerIndex;
  state.sourceRows = matrix.slice(headerIndex + 1);
  calculateProduction();
}

sourceFile.addEventListener("change", async () => {
  const file = sourceFile.files[0];
  if (!file) return;

  try {
    state.sourceWorkbook = await readWorkbook(file);
    setOptions(
      sourceSheet,
      state.sourceWorkbook.SheetNames.map((name) => ({ value: name, label: name }))
    );
    loadSourceSheet();
    $("sourceControls").classList.remove("hidden");
    $("sourceStatus").textContent = `${file.name} — calculated automatically`;
    $("sourceStatus").classList.add("success");
  } catch (error) {
    showMessage("Could not read that report. Confirm it is a valid Excel file.", true);
    console.error(error);
  }
});

sourceSheet.addEventListener("change", loadSourceSheet);
binColumn.addEventListener("change", calculateProduction);
$("calculateBtn").addEventListener("click", calculateProduction);

function calculateProduction() {
  if (!state.sourceRows.length || binColumn.value === "") return;

  const selectedColumn = Number(binColumn.value);
  const aisleTotals = Object.fromEntries(AISLES.map((aisle) => [aisle, 0]));
  const alreadyCountedRows = [];
  let processed = 0;

  state.sourceRows.forEach((row, rowOffset) => {
    const rawLocation = row[selectedColumn];
    const location = String(rawLocation ?? "").trim();
    if (!location) return;

    processed += 1;
    const createdCount = parseCreatedCountLocation(location);

    if (createdCount) {
      aisleTotals[createdCount.aisle] += 1;
    } else {
      alreadyCountedRows.push({
        row: state.sourceHeaderIndex + rowOffset + 2,
        location,
        reason: "Not in created-count format A-00-00",
      });
    }
  });

  const employeeTotals = Object.fromEntries(
    ASSIGNMENTS.map(({ name, aisles }) => [
      name,
      aisles.reduce((sum, aisle) => sum + aisleTotals[aisle], 0),
    ])
  );

  state.aisleTotals = aisleTotals;
  state.employeeTotals = employeeTotals;
  state.alreadyCountedRows = alreadyCountedRows;
  renderResults(processed);
}

function renderResults(processed) {
  const credited = Object.values(state.aisleTotals).reduce((sum, count) => sum + count, 0);
  const held = state.alreadyCountedRows.length;
  $("recordSummary").textContent = `${credited} credited created counts + ${held} awaiting manual credit = ${processed} rows`;

  const aisleCards = AISLES.map((aisle) => `
    <article class="summary-card">
      <div class="summary-card-top">
        <div><strong>Aisle ${aisle}</strong><span>${aisleToEmployee[aisle]}</span></div>
        <b>${state.aisleTotals[aisle]}</b>
      </div>
    </article>
  `).join("");

  const reviewCard = `
    <article class="summary-card">
      <div class="summary-card-top">
        <div><strong>Already Cycle Counted List</strong><span>No automatic employee credit</span></div>
        <b>${held}</b>
      </div>
    </article>
  `;

  const employeeCards = ASSIGNMENTS.map(({ name, aisles }) => {
    const count = state.employeeTotals[name];
    const percent = (count / DAILY_GOAL) * 100;
    const breakdown = aisles.map((aisle) => `${aisle}: ${state.aisleTotals[aisle]}`).join(" • ");
    return `
      <article class="summary-card">
        <div class="summary-card-top">
          <div><strong>${name}</strong><span>${breakdown}</span></div>
          <b>${count}</b>
        </div>
        <div class="meter"><span style="width:${Math.min(percent, 100)}%"></span></div>
        <div class="percent-row"><span>${percent.toFixed(1)}%</span><small>${count - DAILY_GOAL >= 0 ? "+" : ""}${count - DAILY_GOAL} vs goal</small></div>
      </article>
    `;
  }).join("");

  $("summaryCards").innerHTML = aisleCards + reviewCard + employeeCards;

  const details = $("unassignedDetails");
  if (held) {
    details.classList.remove("hidden");
    details.querySelector("summary").innerHTML = `<span id="unassignedCount">${held}</span> rows on Already Cycle Counted List`;
    $("unassignedList").innerHTML = state.alreadyCountedRows
      .slice(0, 250)
      .map((item) => `<div><span>Row ${item.row}</span><strong>${escapeHtml(item.location)}</strong><small>Manual credit</small></div>`)
      .join("");
  } else {
    details.classList.add("hidden");
  }

  $("resultsSection").classList.remove("hidden");
  $("trackerSection").classList.remove("hidden");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

$("downloadSummaryBtn").addEventListener("click", () => {
  if (!state.employeeTotals || !state.aisleTotals) return;

  const aisleRows = AISLES.map((aisle) => ({
    Aisle: aisle,
    Employee: aisleToEmployee[aisle],
    "Credited Created Counts": state.aisleTotals[aisle],
  }));

  const employeeRows = ASSIGNMENTS.map(({ name, aisles }) => ({
    Employee: name,
    "Assigned Aisles": aisles.join("-"),
    "Credited Cycle Counts": state.employeeTotals[name],
    "Production %": state.employeeTotals[name] / DAILY_GOAL,
    "Daily Goal": DAILY_GOAL,
  }));

  const reviewRows = state.alreadyCountedRows.map((item) => ({
    "Source Row": item.row,
    "Bin / Location": item.location,
    Status: "Already cycle counted — assign credit manually",
    "Credit To": "",
  }));

  const workbook = XLSX.utils.book_new();
  const aisleSheet = XLSX.utils.json_to_sheet(aisleRows);
  const employeeSheet = XLSX.utils.json_to_sheet(employeeRows);
  const reviewSheet = XLSX.utils.json_to_sheet(reviewRows.length ? reviewRows : [{ Status: "No rows awaiting manual credit" }]);

  for (let row = 2; row <= employeeRows.length + 1; row += 1) employeeSheet[`D${row}`].z = "0.0%";
  aisleSheet["!cols"] = [{ wch: 10 }, { wch: 18 }, { wch: 24 }];
  employeeSheet["!cols"] = [{ wch: 16 }, { wch: 18 }, { wch: 22 }, { wch: 14 }, { wch: 12 }];
  reviewSheet["!cols"] = [{ wch: 12 }, { wch: 22 }, { wch: 42 }, { wch: 18 }];

  XLSX.utils.book_append_sheet(workbook, aisleSheet, "Aisle Totals");
  XLSX.utils.book_append_sheet(workbook, employeeSheet, "Employee Production");
  XLSX.utils.book_append_sheet(workbook, reviewSheet, "Already Cycle Counted List");
  XLSX.writeFile(workbook, `Cycle_Count_Production_${new Date().toISOString().slice(0, 10)}.xlsx`);
});

trackerFile.addEventListener("change", async () => {
  const file = trackerFile.files[0];
  if (!file) return;

  try {
    state.trackerWorkbook = await readWorkbook(file);
    state.trackerFileName = file.name;
    setOptions(
      targetSheet,
      state.trackerWorkbook.SheetNames.map((name) => ({ value: name, label: name }))
    );
    prepareTrackerSheet();
    $("trackerControls").classList.remove("hidden");
    $("trackerStatus").textContent = file.name;
    $("trackerStatus").classList.add("success");
  } catch (error) {
    showMessage("Could not read the tracker workbook.", true);
    console.error(error);
  }
});

targetSheet.addEventListener("change", prepareTrackerSheet);
headerRow.addEventListener("change", prepareTrackerColumns);

function prepareTrackerSheet() {
  if (!state.trackerWorkbook) return;
  const matrix = worksheetToMatrix(state.trackerWorkbook, targetSheet.value);
  const detectedHeader = findLikelyHeaderRow(matrix, ["name", "employee", "cycle count", "production"]);
  const options = Array.from({ length: Math.min(Math.max(matrix.length, 1), 40) }, (_, index) => ({
    value: index,
    label: `Row ${index + 1}`,
  }));
  setOptions(headerRow, options, detectedHeader);
  prepareTrackerColumns();
}

function prepareTrackerColumns() {
  if (!state.trackerWorkbook) return;
  const matrix = worksheetToMatrix(state.trackerWorkbook, targetSheet.value);
  const rowIndex = Number(headerRow.value || 0);
  const headers = matrix[rowIndex] || [];
  const maxColumns = Math.max(headers.length, 12);
  const options = Array.from({ length: maxColumns }, (_, index) => ({
    value: index,
    label: `${columnLabel(index)} — ${headers[index] || "(blank header)"}`,
  }));
  const optionsWithNone = [{ value: "", label: "Do not update" }, ...options];

  const detectedName = Math.max(0, detectColumn(headers, ["name", "employee", "employee name"]));
  let detectedCount = detectColumn(headers, ["cycle counts", "cycle count", "counts", "count total"]);
  if (detectedCount < 0) detectedCount = Math.min(detectedName + 1, maxColumns - 1);
  const detectedPercent = detectColumn(headers, ["production %", "production percent", "percentage", "%"]);

  setOptions(nameColumn, options, detectedName);
  setOptions(countColumn, options, detectedCount);
  setOptions(percentColumn, optionsWithNone, detectedPercent >= 0 ? detectedPercent : "");
}

$("updateTrackerBtn").addEventListener("click", () => {
  if (!state.trackerWorkbook || !state.employeeTotals) return;

  const sheetName = targetSheet.value;
  const sheet = state.trackerWorkbook.Sheets[sheetName];
  const headerIndex = Number(headerRow.value);
  const nameIndex = Number(nameColumn.value);
  const countIndex = Number(countColumn.value);
  const percentIndex = percentColumn.value === "" ? null : Number(percentColumn.value);
  const addMissing = $("addMissingRows").checked;
  const range = sheet["!ref"]
    ? XLSX.utils.decode_range(sheet["!ref"])
    : { s: { r: 0, c: 0 }, e: { r: headerIndex, c: countIndex } };
  const matched = new Set();
  const updates = [];

  for (let row = headerIndex + 1; row <= range.e.r; row += 1) {
    const nameCell = sheet[XLSX.utils.encode_cell({ r: row, c: nameIndex })];
    const cellName = normalize(nameCell?.v);
    if (!cellName) continue;

    const assignment = ASSIGNMENTS.find(({ name }) => {
      const target = normalize(name);
      return cellName === target || cellName.includes(target) || target.includes(cellName);
    });
    if (!assignment) continue;

    writeResultCells(sheet, row, countIndex, percentIndex, state.employeeTotals[assignment.name]);
    matched.add(assignment.name);
    updates.push(assignment.name);
  }

  if (addMissing) {
    let nextRow = range.e.r + 1;
    ASSIGNMENTS.forEach(({ name }) => {
      if (matched.has(name)) return;
      setCell(sheet, nextRow, nameIndex, name, "s");
      writeResultCells(sheet, nextRow, countIndex, percentIndex, state.employeeTotals[name]);
      matched.add(name);
      updates.push(name);
      nextRow += 1;
    });
    range.e.r = Math.max(range.e.r, nextRow - 1);
  }

  range.e.c = Math.max(range.e.c, countIndex, percentIndex ?? 0, nameIndex);
  sheet["!ref"] = XLSX.utils.encode_range(range);

  if (!updates.length) {
    showMessage("No employee names matched. Check the selected tracker columns.", true);
    return;
  }

  const stem = state.trackerFileName.replace(/\.(xlsx|xls)$/i, "");
  const outputName = `${stem}_UPDATED_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(state.trackerWorkbook, outputName, { cellStyles: true, bookVBA: true });
  showMessage(`Updated ${updates.length} employee rows. ${state.alreadyCountedRows.length} rows still need manual credit.`);
});

function setCell(sheet, row, column, value, type = "n") {
  const address = XLSX.utils.encode_cell({ r: row, c: column });
  const existing = sheet[address] || {};
  sheet[address] = { ...existing, v: value, t: type };
}

function writeResultCells(sheet, row, countIndex, percentIndex, count) {
  setCell(sheet, row, countIndex, count, "n");
  if (percentIndex !== null) {
    setCell(sheet, row, percentIndex, count / DAILY_GOAL, "n");
    sheet[XLSX.utils.encode_cell({ r: row, c: percentIndex })].z = "0.0%";
  }
}

function showMessage(text, isError = false) {
  const message = $("updateMessage");
  message.textContent = text;
  message.classList.remove("hidden", "error");
  if (isError) message.classList.add("error");
}
