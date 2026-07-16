const DAILY_GOAL = 200;
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
const aisleOwner = Object.fromEntries(productionPeople.flatMap((person) => person.aisles.map((aisle) => [aisle, person.name])));

const state = {
  sourceWorkbook: null, sourceRows: [], sourceHeaderIndex: 0,
  alreadyWorkbook: null, alreadyRows: [], alreadyHeaderIndex: 0,
  aisleTotals: Object.fromEntries(AISLES.map((a) => [a, 0])),
  createdTotals: Object.fromEntries(productionPeople.map((p) => [p.name, 0])),
  initialsTotals: Object.fromEntries(PEOPLE.map((p) => [p.name, 0])),
  finalTotals: Object.fromEntries(PEOPLE.map((p) => [p.name, 0])),
  reviewRows: [], trackerWorkbook: null, trackerFileName: "",
};

const $ = (id) => document.getElementById(id);
const normalize = (value) => String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9%]+/g, " ").trim();
const readWorkbook = (file) => file.arrayBuffer().then((buffer) => XLSX.read(buffer, { type: "array", cellStyles: true, cellDates: true, bookVBA: true }));
const matrixFrom = (workbook, sheetName) => XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", blankrows: false, raw: false });
const columnLabel = (index) => XLSX.utils.encode_col(index);

function setOptions(select, options, selected = "") {
  select.innerHTML = "";
  options.forEach(({ value, label }) => {
    const option = document.createElement("option");
    option.value = value; option.textContent = label;
    if (String(value) === String(selected)) option.selected = true;
    select.appendChild(option);
  });
}
function findHeader(matrix, keywords) {
  let best = { index: 0, score: -1 };
  matrix.slice(0, 30).forEach((row, index) => {
    const cells = row.map(normalize);
    const score = keywords.reduce((sum, keyword) => sum + (cells.some((cell) => cell.includes(keyword)) ? 1 : 0), 0);
    if (score > best.score) best = { index, score };
  });
  return best.index;
}
function detectColumn(headers, candidates) {
  const normalized = headers.map(normalize);
  let index = normalized.findIndex((header) => candidates.includes(header));
  if (index < 0) index = normalized.findIndex((header) => candidates.some((candidate) => header.includes(candidate)));
  return index;
}
function columnOptions(headers, includeNone = false) {
  const options = headers.map((header, index) => ({ value: index, label: `${columnLabel(index)} — ${header || "(blank header)"}` }));
  return includeNone ? [{ value: "", label: "Do not update" }, ...options] : options;
}
function parseCreatedCount(value) {
  const text = String(value ?? "").trim().toUpperCase();
  const match = text.match(/^([A-L])-([0-9]{2})-([0-9]{2})$/);
  return match ? { aisle: match[1], location: text } : null;
}
function parseInitials(value) {
  const text = String(value ?? "").trim().toUpperCase();
  const match = text.match(/\b(CH|EH|CC|LM|MJ|AH|GR|DW)\b/);
  return match ? match[1] : null;
}

$("sourceFile").addEventListener("change", async (event) => {
  const file = event.target.files[0]; if (!file) return;
  try {
    state.sourceWorkbook = await readWorkbook(file);
    setOptions($("sourceSheet"), state.sourceWorkbook.SheetNames.map((name) => ({ value: name, label: name })));
    loadSourceSheet();
    $("sourceControls").classList.remove("hidden");
    $("sourceStatus").textContent = `${file.name} — loaded`;
    $("sourceStatus").classList.add("success");
  } catch (error) { showMessage("Could not read the Cycle Count Detail file.", true); console.error(error); }
});
$("sourceSheet").addEventListener("change", loadSourceSheet);
$("binColumn").addEventListener("change", calculateProduction);
$("calculateBtn").addEventListener("click", calculateProduction);

function loadSourceSheet() {
  const matrix = matrixFrom(state.sourceWorkbook, $("sourceSheet").value);
  const headerIndex = findHeader(matrix, ["bin", "location"]);
  const headers = matrix[headerIndex] || [];
  let detected = detectColumn(headers, ["bin #", "bin", "bin number", "location"]); if (detected < 0) detected = 0;
  setOptions($("binColumn"), columnOptions(headers), detected);
  state.sourceHeaderIndex = headerIndex;
  state.sourceRows = matrix.slice(headerIndex + 1);
  calculateProduction();
}

$("alreadyFile").addEventListener("change", async (event) => {
  const file = event.target.files[0]; if (!file) return;
  try {
    state.alreadyWorkbook = await readWorkbook(file);
    setOptions($("alreadySheet"), state.alreadyWorkbook.SheetNames.map((name) => ({ value: name, label: name })));
    loadAlreadySheet();
    $("alreadyControls").classList.remove("hidden");
    $("alreadyStatus").textContent = `${file.name} — initials loaded`;
    $("alreadyStatus").classList.add("success");
  } catch (error) { showMessage("Could not read the Already Cycle Counted List.", true); console.error(error); }
});
$("alreadySheet").addEventListener("change", loadAlreadySheet);
$("initialsColumn").addEventListener("change", calculateProduction);
$("alreadyBinColumn").addEventListener("change", calculateProduction);
$("recalculateBtn").addEventListener("click", calculateProduction);

function loadAlreadySheet() {
  const matrix = matrixFrom(state.alreadyWorkbook, $("alreadySheet").value);
  const headerIndex = findHeader(matrix, ["initial", "employee", "counted by", "bin", "location"]);
  const headers = matrix[headerIndex] || [];
  let initials = detectColumn(headers, ["initials", "initial", "counted by", "employee initials", "employee"]);
  let bin = detectColumn(headers, ["bin #", "bin", "location", "item"]);
  if (initials < 0) initials = 0; if (bin < 0) bin = Math.min(1, Math.max(headers.length - 1, 0));
  setOptions($("initialsColumn"), columnOptions(headers), initials);
  setOptions($("alreadyBinColumn"), columnOptions(headers), bin);
  state.alreadyHeaderIndex = headerIndex;
  state.alreadyRows = matrix.slice(headerIndex + 1);
  calculateProduction();
}

function calculateProduction() {
  const aisleTotals = Object.fromEntries(AISLES.map((a) => [a, 0]));
  const createdTotals = Object.fromEntries(productionPeople.map((p) => [p.name, 0]));
  const initialsTotals = Object.fromEntries(PEOPLE.map((p) => [p.name, 0]));
  const reviewRows = [];
  let sourceProcessed = 0;

  if (state.sourceRows.length && $("binColumn").value !== "") {
    const binIndex = Number($("binColumn").value);
    state.sourceRows.forEach((row, offset) => {
      const location = String(row[binIndex] ?? "").trim(); if (!location) return;
      sourceProcessed += 1;
      const created = parseCreatedCount(location);
      if (created) aisleTotals[created.aisle] += 1;
      else reviewRows.push({ source: "Cycle Count Detail", row: state.sourceHeaderIndex + offset + 2, location, initials: "", status: "Awaiting initials list" });
    });
  }
  productionPeople.forEach((person) => { createdTotals[person.name] = person.aisles.reduce((sum, aisle) => sum + aisleTotals[aisle], 0); });

  let initialsProcessed = 0;
  if (state.alreadyRows.length && $("initialsColumn").value !== "") {
    const initialsIndex = Number($("initialsColumn").value);
    const binIndex = Number($("alreadyBinColumn").value);
    state.alreadyRows.forEach((row, offset) => {
      const rawInitials = row[initialsIndex];
      const location = String(row[binIndex] ?? "").trim();
      if (!String(rawInitials ?? "").trim() && !location) return;
      initialsProcessed += 1;
      const initials = parseInitials(rawInitials);
      const person = initials ? initialsMap[initials] : null;
      if (person) initialsTotals[person.name] += 1;
      else reviewRows.push({ source: "Already Cycle Counted List", row: state.alreadyHeaderIndex + offset + 2, location, initials: String(rawInitials ?? ""), status: "Unknown or blank initials" });
    });
  }

  const finalTotals = Object.fromEntries(PEOPLE.map((person) => [person.name, (createdTotals[person.name] || 0) + (initialsTotals[person.name] || 0)]));
  Object.assign(state, { aisleTotals, createdTotals, initialsTotals, finalTotals, reviewRows });
  renderResults(sourceProcessed, initialsProcessed);
}

function renderResults(sourceProcessed, initialsProcessed) {
  const createdCredited = Object.values(state.createdTotals).reduce((sum, value) => sum + value, 0);
  const initialsCredited = Object.values(state.initialsTotals).reduce((sum, value) => sum + value, 0);
  $("recordSummary").textContent = `${createdCredited} created-count credits + ${initialsCredited} initials credits; ${state.reviewRows.length} need review`;

  const aisleCards = AISLES.map((aisle) => `<article class="summary-card"><div class="summary-card-top"><div><strong>Aisle ${aisle}</strong><span>${aisleOwner[aisle]}</span></div><b>${state.aisleTotals[aisle]}</b></div></article>`).join("");
  const productionCards = productionPeople.map((person) => {
    const total = state.finalTotals[person.name];
    const percent = total / DAILY_GOAL * 100;
    return `<article class="summary-card"><div class="summary-card-top"><div><strong>${person.name}</strong><span>Created: ${state.createdTotals[person.name]} • Initials ${person.initials}: ${state.initialsTotals[person.name]}</span></div><b>${total}</b></div><div class="meter"><span style="width:${Math.min(percent, 100)}%"></span></div><div class="percent-row"><span>${percent.toFixed(1)}%</span><small>${total - DAILY_GOAL >= 0 ? "+" : ""}${total - DAILY_GOAL} vs goal</small></div></article>`;
  }).join("");
  const separateCards = PEOPLE.filter((p) => !p.production).map((person) => `<article class="summary-card"><div class="summary-card-top"><div><strong>${person.name}</strong><span>${person.initials} • tracked separately</span></div><b>${state.initialsTotals[person.name]}</b></div></article>`).join("");
  $("summaryCards").innerHTML = aisleCards + productionCards + separateCards;

  const details = $("unassignedDetails");
  if (state.reviewRows.length) {
    details.classList.remove("hidden");
    $("unassignedCount").textContent = state.reviewRows.length;
    $("unassignedList").innerHTML = state.reviewRows.slice(0, 300).map((item) => `<div><span>${item.source} row ${item.row}</span><strong>${escapeHtml(item.location || "No location")}</strong><small>${escapeHtml(item.initials || item.status)}</small></div>`).join("");
  } else details.classList.add("hidden");
  if (sourceProcessed || initialsProcessed) { $("resultsSection").classList.remove("hidden"); $("trackerSection").classList.remove("hidden"); }
}

function escapeHtml(value) { return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#039;"); }

$("downloadSummaryBtn").addEventListener("click", () => {
  const workbook = XLSX.utils.book_new();
  const employeeRows = PEOPLE.map((person) => ({ Employee: person.name, Initials: person.initials, "Assigned Aisles": person.aisles.join("-"), "Created Counts": state.createdTotals[person.name] || 0, "Initials List Counts": state.initialsTotals[person.name] || 0, "Total Counts": state.finalTotals[person.name] || 0, "Production %": person.production ? state.finalTotals[person.name] / DAILY_GOAL : "Tracked separately" }));
  const aisleRows = AISLES.map((aisle) => ({ Aisle: aisle, Employee: aisleOwner[aisle], "Created Counts": state.aisleTotals[aisle] }));
  const reviewRows = state.reviewRows.length ? state.reviewRows : [{ Status: "No rows need review" }];
  const employeeSheet = XLSX.utils.json_to_sheet(employeeRows); const aisleSheet = XLSX.utils.json_to_sheet(aisleRows); const reviewSheet = XLSX.utils.json_to_sheet(reviewRows);
  productionPeople.forEach((_, index) => { const cell = employeeSheet[`G${index + 2}`]; if (cell) cell.z = "0.0%"; });
  XLSX.utils.book_append_sheet(workbook, employeeSheet, "Employee Production");
  XLSX.utils.book_append_sheet(workbook, aisleSheet, "Created Count Aisles");
  XLSX.utils.book_append_sheet(workbook, reviewSheet, "Needs Review");
  XLSX.writeFile(workbook, `Cycle_Count_Production_${new Date().toISOString().slice(0, 10)}.xlsx`);
});

$("trackerFile").addEventListener("change", async (event) => {
  const file = event.target.files[0]; if (!file) return;
  try {
    state.trackerWorkbook = await readWorkbook(file); state.trackerFileName = file.name;
    setOptions($("targetSheet"), state.trackerWorkbook.SheetNames.map((name) => ({ value: name, label: name })));
    prepareTracker(); $("trackerControls").classList.remove("hidden"); $("trackerStatus").textContent = file.name; $("trackerStatus").classList.add("success");
  } catch (error) { showMessage("Could not read the tracker workbook.", true); }
});
$("targetSheet").addEventListener("change", prepareTracker);
$("headerRow").addEventListener("change", prepareTrackerColumns);
function prepareTracker() {
  const matrix = matrixFrom(state.trackerWorkbook, $("targetSheet").value);
  const header = findHeader(matrix, ["name", "employee", "cycle count", "production"]);
  setOptions($("headerRow"), Array.from({ length: Math.min(Math.max(matrix.length, 1), 40) }, (_, i) => ({ value: i, label: `Row ${i + 1}` })), header);
  prepareTrackerColumns();
}
function prepareTrackerColumns() {
  const matrix = matrixFrom(state.trackerWorkbook, $("targetSheet").value); const headers = matrix[Number($("headerRow").value || 0)] || [];
  const options = columnOptions(Array.from({ length: Math.max(headers.length, 12) }, (_, i) => headers[i] || ""));
  let name = detectColumn(headers, ["name", "employee", "employee name"]); if (name < 0) name = 0;
  let count = detectColumn(headers, ["cycle counts", "cycle count", "counts"]); if (count < 0) count = Math.min(name + 1, options.length - 1);
  const percent = detectColumn(headers, ["production %", "production percent", "percentage", "%"]);
  setOptions($("nameColumn"), options, name); setOptions($("countColumn"), options, count); setOptions($("percentColumn"), [{ value: "", label: "Do not update" }, ...options], percent >= 0 ? percent : "");
}
$("updateTrackerBtn").addEventListener("click", () => {
  if (!state.trackerWorkbook) return;
  const sheet = state.trackerWorkbook.Sheets[$("targetSheet").value]; const header = Number($("headerRow").value); const nameCol = Number($("nameColumn").value); const countCol = Number($("countColumn").value); const percentCol = $("percentColumn").value === "" ? null : Number($("percentColumn").value);
  const range = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : { s: { r: 0, c: 0 }, e: { r: header, c: countCol } }; const matched = new Set(); let updates = 0;
  for (let row = header + 1; row <= range.e.r; row += 1) {
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: nameCol })]; const name = normalize(cell?.v); if (!name) continue;
    const person = productionPeople.find((p) => name === normalize(p.name) || name.includes(normalize(p.name))); if (!person) continue;
    setCell(sheet, row, countCol, state.finalTotals[person.name]); if (percentCol !== null) { setCell(sheet, row, percentCol, state.finalTotals[person.name] / DAILY_GOAL); sheet[XLSX.utils.encode_cell({ r: row, c: percentCol })].z = "0.0%"; }
    matched.add(person.name); updates += 1;
  }
  if ($("addMissingRows").checked) {
    let row = range.e.r + 1;
    productionPeople.forEach((person) => { if (matched.has(person.name)) return; setCell(sheet, row, nameCol, person.name, "s"); setCell(sheet, row, countCol, state.finalTotals[person.name]); if (percentCol !== null) { setCell(sheet, row, percentCol, state.finalTotals[person.name] / DAILY_GOAL); sheet[XLSX.utils.encode_cell({ r: row, c: percentCol })].z = "0.0%"; } row += 1; updates += 1; });
    range.e.r = Math.max(range.e.r, row - 1);
  }
  range.e.c = Math.max(range.e.c, nameCol, countCol, percentCol ?? 0); sheet["!ref"] = XLSX.utils.encode_range(range);
  if (!updates) return showMessage("No production employee names matched the selected tracker columns.", true);
  const stem = state.trackerFileName.replace(/\.(xlsx|xls)$/i, ""); const output = `${stem}_UPDATED_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(state.trackerWorkbook, output, { cellStyles: true, bookVBA: true }); showMessage(`Updated ${updates} production rows and downloaded ${output}. Greg and Denise remain separate.`);
});
function setCell(sheet, row, column, value, type = "n") { const address = XLSX.utils.encode_cell({ r: row, c: column }); sheet[address] = { ...(sheet[address] || {}), v: value, t: type }; }
function showMessage(text, isError = false) { const message = $("updateMessage"); message.textContent = text; message.classList.remove("hidden", "error"); if (isError) message.classList.add("error"); }
