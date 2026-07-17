/**
 * Cycle Count Production - Complete Multi-Branch Implementation
 * Full localStorage persistence, employee management, and dynamic aisle assignment
 */

const DAILY_GOAL = 200;
const BRANCH_STORAGE_KEY = "cycleCountProduction.branches.v1";
const SELECTED_BRANCH_STORAGE_KEY = "cycleCountProduction.selectedBranch.v1";
const SNAPSHOTS_STORAGE_KEY = "cycleCountProduction.snapshots.v3";

// Default data for first-time users
const DEFAULT_BRANCHES = [
  {
    id: "default-branch",
    name: "Default Branch",
    expectedInventoryFilename: "",
    assignments: [
      { id: "employee-1", name: "Carico", startAisle: "A", endAisle: "B" },
      { id: "employee-2", name: "Ernie", startAisle: "C", endAisle: "D" },
      { id: "employee-3", name: "Cherish", startAisle: "E", endAisle: "F" },
      { id: "employee-4", name: "Layne", startAisle: "G", endAisle: "H" },
      { id: "employee-5", name: "Madison", startAisle: "I", endAisle: "J" },
      { id: "employee-6", name: "Antoine", startAisle: "K", endAisle: "L" },
    ],
  },
];

// Application state
const state = {
  workbook: null,
  rows: [],
  headerIndex: 0,
  aisleTotals: {},
  employeeTotals: {},
  uncreditedRows: [],
  uploadedFileName: "",
};

let branches = [];
let selectedBranchId = null;

// DOM utilities
const $ = (id) => document.getElementById(id);

/**
 * BRANCH STORAGE & RETRIEVAL
 */

function loadBranches() {
  try {
    const data = localStorage.getItem(BRANCH_STORAGE_KEY);
    return data ? JSON.parse(data) : null;
  } catch (e) {
    console.error("Failed to load branches:", e);
    return null;
  }
}

function saveBranches() {
  try {
    localStorage.setItem(BRANCH_STORAGE_KEY, JSON.stringify(branches));
  } catch (e) {
    console.error("Failed to save branches:", e);
  }
}

function getSelectedBranch() {
  return branches.find((b) => b.id === selectedBranchId) || null;
}

function selectBranch(branchId) {
  if (branches.find((b) => b.id === branchId)) {
    selectedBranchId = branchId;
    try {
      localStorage.setItem(SELECTED_BRANCH_STORAGE_KEY, branchId);
    } catch (e) {
      console.error("Failed to save selected branch:", e);
    }
    return true;
  }
  return false;
}

function getAssignments() {
  const branch = getSelectedBranch();
  return branch ? branch.assignments : [];
}

function getAssignedAisles() {
  const assignments = getAssignments();
  const aisles = new Set();
  assignments.forEach((assignment) => {
    const range = expandAisleRange(assignment.startAisle, assignment.endAisle);
    range.forEach((a) => aisles.add(a));
  });
  return Array.from(aisles).sort();
}

/**
 * AISLE RANGE UTILITIES
 */

function expandAisleRange(start, end) {
  const startCode = start.charCodeAt(0);
  const endCode = end.charCodeAt(0);
  const aisles = [];
  for (let code = startCode; code <= endCode; code++) {
    aisles.push(String.fromCharCode(code));
  }
  return aisles;
}

function formatAisleRange(start, end) {
  if (start === end) {
    return start;
  }
  return `${start}–${end}`;
}

/**
 * BRANCH MANAGEMENT
 */

function renderBranchDropdown() {
  const dropdown = $("branchSelect");
  dropdown.innerHTML = branches
    .map(
      (branch) =>
        `<option value="${branch.id}" ${branch.id === selectedBranchId ? "selected" : ""}>${escapeHtml(branch.name)}</option>`
    )
    .join("");
}

function addBranch(name, expectedFilename = "") {
  if (!name || !name.trim()) {
    showBranchMessage("Branch name cannot be empty", true);
    return false;
  }

  if (branches.some((b) => b.name.toLowerCase() === name.toLowerCase())) {
    showBranchMessage("Branch name already exists", true);
    return false;
  }

  const newBranch = {
    id: `branch-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
    name: name.trim(),
    expectedInventoryFilename: expectedFilename.trim(),
    assignments: [],
  };

  branches.push(newBranch);
  saveBranches();
  selectBranch(newBranch.id);
  return true;
}

function renameBranch(branchId, newName) {
  if (!newName || !newName.trim()) {
    showBranchMessage("Branch name cannot be empty", true);
    return false;
  }

  if (branches.some((b) => b.id !== branchId && b.name.toLowerCase() === newName.toLowerCase())) {
    showBranchMessage("Branch name already exists", true);
    return false;
  }

  const branch = branches.find((b) => b.id === branchId);
  if (branch) {
    branch.name = newName.trim();
    saveBranches();
    return true;
  }

  return false;
}

function deleteBranch(branchId) {
  if (branches.length === 1) {
    showBranchMessage("Cannot delete the last remaining branch", true);
    return false;
  }

  branches = branches.filter((b) => b.id !== branchId);

  if (selectedBranchId === branchId) {
    selectedBranchId = branches[0]?.id || null;
    if (selectedBranchId) {
      localStorage.setItem(SELECTED_BRANCH_STORAGE_KEY, selectedBranchId);
    }
  }

  saveBranches();
  return true;
}

/**
 * EMPLOYEE ASSIGNMENT MANAGEMENT
 */

function addAssignment(name, startAisle, endAisle) {
  const validation = validateAssignment(name, startAisle, endAisle);
  if (!validation.valid) {
    showTeamMessage(validation.error, true);
    return false;
  }

  const branch = getSelectedBranch();
  if (!branch) return false;

  const newAssignment = {
    id: `employee-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
    name: name.trim(),
    startAisle: startAisle.toUpperCase(),
    endAisle: endAisle.toUpperCase(),
  };

  branch.assignments.push(newAssignment);
  saveBranches();
  return true;
}

function updateAssignment(assignmentId, name, startAisle, endAisle) {
  const validation = validateAssignment(name, startAisle, endAisle, assignmentId);
  if (!validation.valid) {
    showTeamMessage(validation.error, true);
    return false;
  }

  const branch = getSelectedBranch();
  if (!branch) return false;

  const assignment = branch.assignments.find((a) => a.id === assignmentId);
  if (!assignment) return false;

  assignment.name = name.trim();
  assignment.startAisle = startAisle.toUpperCase();
  assignment.endAisle = endAisle.toUpperCase();
  saveBranches();
  return true;
}

function deleteAssignment(assignmentId) {
  const branch = getSelectedBranch();
  if (!branch) return false;

  const assignment = branch.assignments.find((a) => a.id === assignmentId);
  if (!assignment) return false;

  if (!confirm(`Remove ${assignment.name} from the team?`)) {
    return false;
  }

  branch.assignments = branch.assignments.filter((a) => a.id !== assignmentId);
  saveBranches();
  return true;
}

/**
 * VALIDATION
 */

function validateAssignment(name, startAisle, endAisle, excludeId = null) {
  if (!name || !name.trim()) {
    return { valid: false, error: "Employee name cannot be empty" };
  }

  if (!startAisle || !endAisle) {
    return { valid: false, error: "Both aisles are required" };
  }

  const start = startAisle.toUpperCase();
  const end = endAisle.toUpperCase();

  if (!/^[A-Z]$/.test(start) || !/^[A-Z]$/.test(end)) {
    return { valid: false, error: "Aisles must be A-Z" };
  }

  if (start > end) {
    return { valid: false, error: "Starting aisle cannot be after ending aisle" };
  }

  const branch = getSelectedBranch();
  if (!branch) {
    return { valid: false, error: "No active branch" };
  }

  // Check for duplicate names (excluding the current assignment if editing)
  if (
    branch.assignments.some(
      (a) => a.id !== excludeId && a.name.toLowerCase() === name.toLowerCase()
    )
  ) {
    return { valid: false, error: `${name} already exists in this branch` };
  }

  // Check for overlapping aisles (excluding the current assignment if editing)
  const newRange = expandAisleRange(start, end);
  for (const assignment of branch.assignments) {
    if (excludeId && assignment.id === excludeId) continue;

    const existingRange = expandAisleRange(assignment.startAisle, assignment.endAisle);
    const hasOverlap = newRange.some((aisle) => existingRange.includes(aisle));

    if (hasOverlap) {
      return {
        valid: false,
        error: `Aisle range overlaps with ${assignment.name}`,
      };
    }
  }

  return { valid: true };
}

/**
 * UI RENDERING
 */

function renderAssignments() {
  const table = $("teamBody");
  const assignments = getAssignments();

  if (assignments.length === 0) {
    table.innerHTML = `<tr><td colspan="4" class="empty-row">No team members. Add one to get started.</td></tr>`;
    renderAssignmentGrid();
    return;
  }

  table.innerHTML = assignments
    .map(
      (assignment) => `
    <tr>
      <td>${escapeHtml(assignment.name)}</td>
      <td>${formatAisleRange(assignment.startAisle, assignment.endAisle)}</td>
      <td><button class="team-edit-btn secondary small" data-employee-id="${assignment.id}">Edit</button></td>
      <td><button class="team-delete-btn secondary small danger" data-employee-id="${assignment.id}">Delete</button></td>
    </tr>
  `
    )
    .join("");

  // Attach handlers
  table.querySelectorAll(".team-edit-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const assignment = assignments.find((a) => a.id === btn.dataset.employeeId);
      if (assignment) {
        openAssignmentForm(assignment);
      }
    });
  });

  table.querySelectorAll(".team-delete-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      if (deleteAssignment(btn.dataset.employeeId)) {
        showTeamMessage("Employee removed");
        clearProductionState();
        renderAssignments();
        renderResults();
      }
    });
  });

  renderAssignmentGrid();
}

function renderAssignmentGrid() {
  const grid = $("assignmentGrid");
  const assignments = getAssignments();

  if (assignments.length === 0) {
    grid.innerHTML = `<div class="empty-assignments">No team members assigned. Add team members to get started.</div>`;
    return;
  }

  grid.innerHTML = assignments
    .map(
      (assignment) =>
        `<div>
      <strong>${escapeHtml(assignment.name)}</strong>
      <span>${formatAisleRange(assignment.startAisle, assignment.endAisle)}</span>
    </div>`
    )
    .join("");
}

function updateUIFromBranch() {
  const branch = getSelectedBranch();
  if (!branch) return;

  const assignments = branch.assignments;
  const dailyGoal = DAILY_GOAL;
  const teamGoal = dailyGoal * assignments.length;

  // Update daily goal display
  $("dailyGoalDisplay").textContent = dailyGoal;
  $("dailyGoalUnit").textContent = `counts = 100% (${assignments.length} employees)`;

  // Update expected filename hint
  const hint = $("expectedFilename");
  if (branch.expectedInventoryFilename) {
    hint.innerHTML = `<strong>Expected file:</strong> ${escapeHtml(branch.expectedInventoryFilename)}<br><span>The app uses the Batch column and credits created-count batches such as C-09-12.</span>`;
  } else {
    hint.textContent = "The app uses the Batch column and credits created-count batches such as C-09-12.";
  }

  renderBranchDropdown();
  renderAssignments();
}

/**
 * FORM MANAGEMENT
 */

function openAssignmentForm(assignment = null) {
  const modal = $("teamMemberModal");
  const form = $("teamMemberForm");
  const nameInput = $("employeeNameInput");
  const startInput = $("startAisleInput");
  const endInput = $("endAisleInput");

  if (assignment) {
    nameInput.value = assignment.name;
    startInput.value = assignment.startAisle;
    endInput.value = assignment.endAisle;
    form.dataset.mode = "edit";
    form.dataset.assignmentId = assignment.id;
  } else {
    nameInput.value = "";
    startInput.value = "";
    endInput.value = "";
    form.dataset.mode = "add";
    form.dataset.assignmentId = "";
  }

  modal.showModal();
}

function saveAssignment() {
  const form = $("teamMemberForm");
  const name = $("employeeNameInput").value.trim();
  const startAisle = $("startAisleInput").value.trim();
  const endAisle = $("endAisleInput").value.trim();

  if (!name || !startAisle || !endAisle) {
    showTeamMessage("All fields are required", true);
    return false;
  }

  let success;
  if (form.dataset.mode === "add") {
    success = addAssignment(name, startAisle, endAisle);
    if (success) {
      showTeamMessage(`${name} added to team`);
    }
  } else {
    success = updateAssignment(form.dataset.assignmentId, name, startAisle, endAisle);
    if (success) {
      showTeamMessage(`${name} updated`);
    }
  }

  if (success) {
    clearProductionState();
    renderAssignments();
    renderResults();
    $("teamMemberModal").close();
  }

  return success;
}

function resetAssignmentForm() {
  const form = $("teamMemberForm");
  form.reset();
  form.dataset.mode = "add";
  form.dataset.assignmentId = "";
}

/**
 * MESSAGE DISPLAY
 */

function showBranchMessage(text, isError = false) {
  const message = $("branchMessage");
  if (!message) return;
  message.textContent = text;
  message.classList.remove("hidden", "error");
  if (isError) message.classList.add("error");
  setTimeout(() => message.classList.add("hidden"), 4000);
}

function showTeamMessage(text, isError = false) {
  const message = $("teamMessage");
  if (!message) return;
  message.textContent = text;
  message.classList.remove("hidden", "error");
  if (isError) message.classList.add("error");
  setTimeout(() => message.classList.add("hidden"), 4000);
}

function showSaveMessage(text, isError = false) {
  const message = $("saveMessage");
  if (!message) return;
  message.textContent = text;
  message.classList.remove("hidden", "error");
  if (isError) message.classList.add("error");
  setTimeout(() => message.classList.add("hidden"), 4000);
}

/**
 * PRODUCTION CALCULATIONS
 */

function clearProductionState() {
  state.workbook = null;
  state.rows = [];
  state.headerIndex = 0;
  state.aisleTotals = {};
  state.employeeTotals = {};
  state.uncreditedRows = [];
  state.uploadedFileName = "";
  const fileInput = $("sourceFile");
  if (fileInput) fileInput.value = "";
  const resultsSection = $("resultsSection");
  if (resultsSection) resultsSection.classList.add("hidden");
  const sourceControls = $("sourceControls");
  if (sourceControls) sourceControls.classList.add("hidden");
  const sourceStatus = $("sourceStatus");
  if (sourceStatus) {
    sourceStatus.textContent = "Waiting for file";
    sourceStatus.classList.remove("success");
  }
}

function normalize(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9#%]+/g, " ")
    .trim();
}

function readWorkbook(file) {
  return file.arrayBuffer().then((buffer) =>
    XLSX.read(buffer, {
      type: "array",
      cellDates: true,
      cellStyles: true,
    })
  );
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
  return normalized.findIndex((header) =>
    candidates.some((candidate) => header.includes(candidate))
  );
}

function parseCreatedBatch(value) {
  const text = String(value ?? "").trim().toUpperCase();
  const match = text.match(/^([A-Z])-\d{2}-\d{2}$/);
  return match ? match[1] : null;
}

function loadSelectedSheet() {
  const sourceSheet = $("sourceSheet");
  const batchColumn = $("batchColumn");
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
  const batchColumn = $("batchColumn");
  const assignments = getAssignments();
  const assignedAisles = getAssignedAisles();

  // Initialize aisle totals with only assigned aisles
  state.aisleTotals = Object.fromEntries(assignedAisles.map((aisle) => [aisle, 0]));

  // Initialize employee totals
  state.employeeTotals = Object.fromEntries(assignments.map((a) => [a.name, 0]));

  state.uncreditedRows = [];

  if (state.rows.length && batchColumn.value !== "") {
    const column = Number(batchColumn.value);
    state.rows.forEach((row, offset) => {
      const batch = String(row[column] ?? "").trim();
      if (!batch) return;

      const aisle = parseCreatedBatch(batch);

      if (aisle && assignedAisles.includes(aisle)) {
        // Aisle is assigned to someone
        state.aisleTotals[aisle] += 1;
      } else if (aisle) {
        // Aisle is valid letter but not assigned
        state.uncreditedRows.push({
          row: state.headerIndex + offset + 2,
          batch,
          reason: `Aisle ${aisle} not assigned to any employee`,
        });
      } else {
        // Invalid batch format
        state.uncreditedRows.push({
          row: state.headerIndex + offset + 2,
          batch,
          reason: "Not a valid batch format (expected A-00-00)",
        });
      }
    });
  }

  // Calculate employee totals
  assignments.forEach((assignment) => {
    const aisles = expandAisleRange(assignment.startAisle, assignment.endAisle);
    const total = aisles.reduce((sum, aisle) => sum + (state.aisleTotals[aisle] || 0), 0);
    state.employeeTotals[assignment.name] = total;
  });

  renderResults();
}

function renderResults() {
  if (!state.workbook) return;

  const assignments = getAssignments();
  const credited = Object.values(state.aisleTotals).reduce((sum, count) => sum + count, 0);
  const uncredited = state.uncreditedRows.length;
  const teamGoal = DAILY_GOAL * assignments.length;

  const resultsSection = $("resultsSection");
  if (resultsSection) resultsSection.classList.remove("hidden");

  const recordSummary = $("recordSummary");
  if (recordSummary) recordSummary.textContent = `${credited} credited • ${uncredited} uncredited`;

  const kpiStrip = $("kpiStrip");
  if (kpiStrip) {
    kpiStrip.innerHTML = `
      <div class="kpi"><span>Credited created counts</span><strong>${credited}</strong></div>
      <div class="kpi"><span>Uncredited rows</span><strong>${uncredited}</strong></div>
      <div class="kpi"><span>Daily team goal</span><strong>${teamGoal}</strong></div>
      <div class="kpi"><span>Team production</span><strong>${teamGoal > 0 ? ((credited / teamGoal) * 100).toFixed(1) : 0}%</strong></div>`;
  }

  const productionCards = $("productionCards");
  if (productionCards) {
    if (assignments.length === 0) {
      productionCards.innerHTML =
        `<div class="empty-assignments">No team members to calculate production. Add team members first.</div>`;
    } else {
      productionCards.innerHTML = assignments
        .map((assignment) => {
          const total = state.employeeTotals[assignment.name] || 0;
          const percent = (total / DAILY_GOAL) * 100;
          const aisles = expandAisleRange(assignment.startAisle, assignment.endAisle);
          const breakdown = aisles.map((aisle) => `${aisle}: ${state.aisleTotals[aisle] || 0}`).join(" • ");
          return `<article class="summary-card">
            <div class="summary-card-top">
              <div><strong>${escapeHtml(assignment.name)}</strong><span>${breakdown}</span></div>
              <b>${total}</b>
            </div>
            <div class="meter"><span style="width:${Math.min(percent, 100)}%"></span></div>
            <div class="percent-row"><span>${percent.toFixed(1)}%</span><small>${total - DAILY_GOAL >= 0 ? "+" : ""}${total - DAILY_GOAL} vs goal</small></div>
          </article>`;
        })
        .join("");
    }
  }

  const reviewDetails = $("reviewDetails");
  if (reviewDetails) {
    if (uncredited) {
      reviewDetails.classList.remove("hidden");
      const reviewCount = $("reviewCount");
      if (reviewCount) reviewCount.textContent = uncredited;
      const reviewList = $("reviewList");
      if (reviewList) {
        reviewList.innerHTML = state.uncreditedRows
          .slice(0, 300)
          .map(
            (item) =>
              `<div><span>Row ${item.row}</span><strong>${escapeHtml(item.batch)}</strong><small>${escapeHtml(item.reason || "Unassigned")}</small></div>`
          )
          .join("");
      }
    } else {
      reviewDetails.classList.add("hidden");
    }
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

/**
 * HISTORY / SNAPSHOTS
 */

function loadHistory() {
  try {
    return JSON.parse(localStorage.getItem(SNAPSHOTS_STORAGE_KEY) || "[]");
  } catch {
    return [];
  }
}

function saveHistory(history) {
  localStorage.setItem(SNAPSHOTS_STORAGE_KEY, JSON.stringify(history));
}

function currentRecord() {
  const branch = getSelectedBranch();
  return {
    id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now()),
    savedAt: new Date().toISOString(),
    branchId: branch?.id,
    branchName: branch?.name,
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
  showSaveMessage(`Results saved at ${new Date(record.savedAt).toLocaleString()}.`);
  renderHistory();
}

function renderHistory() {
  const records = loadHistory();
  const historyBody = $("historyBody");
  if (!historyBody) return;

  historyBody.innerHTML = records.length
    ? records
        .map((record) => {
          const teamTotal = Object.values(record.employeeTotals || {}).reduce(
            (sum, count) => sum + count,
            0
          );
          const branchLabel = record.branchName ? ` (${record.branchName})` : "";
          return `<tr>
            <td>${new Date(record.savedAt).toLocaleString()}${branchLabel}</td>
            <td>${teamTotal}</td>
            <td>${record.uncreditedCount || 0}</td>
            <td><button data-delete-id="${record.id}" type="button">Delete</button></td>
          </tr>`;
        })
        .join("")
    : `<tr><td colspan="4" class="empty-row">No saved results yet.</td></tr>`;
}

/**
 * DOWNLOAD FUNCTIONALITY
 */

function downloadSummary() {
  const assignments = getAssignments();
  const assignedAisles = getAssignedAisles();

  if (assignments.length === 0) {
    showSaveMessage("Cannot download summary with no team members", true);
    return;
  }

  const employeeRows = assignments.map((assignment) => {
    const total = state.employeeTotals[assignment.name] || 0;
    const aisles = expandAisleRange(assignment.startAisle, assignment.endAisle);
    return {
      Employee: assignment.name,
      "Assigned Aisles": formatAisleRange(assignment.startAisle, assignment.endAisle),
      "Aisle Breakdown": aisles.map((aisle) => `${aisle}: ${state.aisleTotals[aisle] || 0}`).join(" | "),
      "Cycle Counts": total,
      "Production %": total / DAILY_GOAL,
      "Daily Goal": DAILY_GOAL,
    };
  });

  const aisleRows = assignedAisles.map((aisle) => ({
    Aisle: aisle,
    "Cycle Counts": state.aisleTotals[aisle] || 0,
  }));

  const workbook = XLSX.utils.book_new();
  const productionSheet = XLSX.utils.json_to_sheet(employeeRows);
  employeeRows.forEach((row, index) => {
    productionSheet[`E${index + 2}`].z = "0.0%";
  });
  productionSheet["!cols"] = [
    { wch: 16 },
    { wch: 18 },
    { wch: 24 },
    { wch: 14 },
    { wch: 14 },
    { wch: 12 },
  ];

  XLSX.utils.book_append_sheet(workbook, productionSheet, "Employee Production");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(aisleRows), "Aisle Totals");

  if (state.uncreditedRows.length) {
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(state.uncreditedRows),
      "Uncredited Rows"
    );
  }

  const stamp = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 16);
  const branch = getSelectedBranch();
  const branchName = branch?.name?.replace(/[^a-z0-9]/gi, "_") || "Export";
  XLSX.writeFile(workbook, `Cycle_Count_${branchName}_${stamp}.xlsx`);
}

/**
 * EVENT LISTENERS SETUP
 */

function setupEventListeners() {
  // Branch management
  const branchSelect = $("branchSelect");
  if (branchSelect) {
    branchSelect.addEventListener("change", () => {
      selectBranch(branchSelect.value);
      clearProductionState();
      updateUIFromBranch();
    });
  }

  const addBranchBtn = $("addBranchBtn");
  if (addBranchBtn) {
    addBranchBtn.addEventListener("click", () => {
      const branchModal = $("branchModal");
      const branchNameInput = $("branchNameInput");
      const expectedFilenameInput = $("expectedFilenameInput");
      const branchForm = $("branchForm");

      branchNameInput.value = "";
      expectedFilenameInput.value = "";
      branchForm.dataset.mode = "add";
      branchModal.showModal();
    });
  }

  const renameBranchBtn = $("renameBranchBtn");
  if (renameBranchBtn) {
    renameBranchBtn.addEventListener("click", () => {
      const branch = getSelectedBranch();
      if (!branch) return;

      const branchModal = $("branchModal");
      const branchNameInput = $("branchNameInput");
      const expectedFilenameInput = $("expectedFilenameInput");
      const branchForm = $("branchForm");
      const branchModalTitle = $("branchModalTitle");

      branchNameInput.value = branch.name;
      expectedFilenameInput.value = branch.expectedInventoryFilename || "";
      branchForm.dataset.mode = "edit";
      branchForm.dataset.branchId = branch.id;
      branchModalTitle.textContent = "Edit Branch";
      branchModal.showModal();
    });
  }

  const deleteBranchBtn = $("deleteBranchBtn");
  if (deleteBranchBtn) {
    deleteBranchBtn.addEventListener("click", () => {
      const branch = getSelectedBranch();
      if (!branch) return;

      if (confirm(`Delete branch "${branch.name}"? This cannot be undone.`)) {
        if (deleteBranch(branch.id)) {
          showBranchMessage(`Branch deleted. Switched to ${getSelectedBranch().name}`);
          clearProductionState();
          updateUIFromBranch();
        }
      }
    });
  }

  const branchForm = $("branchForm");
  if (branchForm) {
    branchForm.addEventListener("submit", (e) => {
      e.preventDefault();

      const branchNameInput = $("branchNameInput");
      const expectedFilenameInput = $("expectedFilenameInput");
      const name = branchNameInput.value.trim();
      const expectedFilename = expectedFilenameInput.value.trim();

      if (!name) {
        showBranchMessage("Branch name cannot be empty", true);
        return;
      }

      let success;
      const branchModal = $("branchModal");

      if (branchForm.dataset.mode === "add") {
        success = addBranch(name, expectedFilename);
        if (success) {
          showBranchMessage(`Branch "${name}" created`);
        }
      } else {
        const branchId = branchForm.dataset.branchId;
        success = renameBranch(branchId, name);
        if (success) {
          const branch = branches.find((b) => b.id === branchId);
          if (branch) {
            branch.expectedInventoryFilename = expectedFilename;
            saveBranches();
          }
          showBranchMessage("Branch updated");
        }
      }

      if (success) {
        clearProductionState();
        updateUIFromBranch();
        branchModal.close();
      }
    });
  }

  const branchCancel = $("branchCancel");
  if (branchCancel) {
    branchCancel.addEventListener("click", () => {
      $("branchModal").close();
    });
  }

  // Team member management
  const addTeamMemberBtn = $("addTeamMemberBtn");
  if (addTeamMemberBtn) {
    addTeamMemberBtn.addEventListener("click", () => {
      openAssignmentForm();
    });
  }

  const teamMemberForm = $("teamMemberForm");
  if (teamMemberForm) {
    teamMemberForm.addEventListener("submit", (e) => {
      e.preventDefault();
      saveAssignment();
    });
  }

  const teamMemberCancel = $("teamMemberCancel");
  if (teamMemberCancel) {
    teamMemberCancel.addEventListener("click", () => {
      resetAssignmentForm();
      $("teamMemberModal").close();
    });
  }

  // File upload
  const sourceFile = $("sourceFile");
  if (sourceFile) {
    sourceFile.addEventListener("change", async () => {
      const file = sourceFile.files[0];
      if (!file) return;

      try {
        state.workbook = await readWorkbook(file);
        state.uploadedFileName = file.name;

        const sourceSheet = $("sourceSheet");
        setOptions(
          sourceSheet,
          state.workbook.SheetNames.map((name) => ({ value: name, label: name }))
        );

        loadSelectedSheet();

        const sourceControls = $("sourceControls");
        if (sourceControls) sourceControls.classList.remove("hidden");

        const sourceStatus = $("sourceStatus");
        if (sourceStatus) {
          sourceStatus.textContent = `${file.name} loaded`;
          sourceStatus.classList.add("success");
        }
      } catch (error) {
        console.error(error);
        const sourceStatus = $("sourceStatus");
        if (sourceStatus) {
          sourceStatus.textContent = "Could not read file";
          sourceStatus.classList.remove("success");
        }
      }
    });
  }

  const sourceSheet = $("sourceSheet");
  if (sourceSheet) {
    sourceSheet.addEventListener("change", loadSelectedSheet);
  }

  const batchColumn = $("batchColumn");
  if (batchColumn) {
    batchColumn.addEventListener("change", calculateProduction);
  }

  const saveSnapshotBtn = $("saveSnapshotBtn");
  if (saveSnapshotBtn) {
    saveSnapshotBtn.addEventListener("click", saveSnapshot);
  }

  const downloadSummaryBtn = $("downloadSummaryBtn");
  if (downloadSummaryBtn) {
    downloadSummaryBtn.addEventListener("click", downloadSummary);
  }

  const historyBody = $("historyBody");
  if (historyBody) {
    historyBody.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-delete-id]");
      if (!button) return;
      saveHistory(loadHistory().filter((record) => record.id !== button.dataset.deleteId));
      renderHistory();
    });
  }
}

/**
 * INITIALIZATION
 */

function initializeApp() {
  // Load branches from localStorage
  const savedBranches = loadBranches();
  branches = savedBranches || DEFAULT_BRANCHES;
  saveBranches();

  // Restore selected branch
  const savedBranchId = localStorage.getItem(SELECTED_BRANCH_STORAGE_KEY);
  if (savedBranchId && branches.find((b) => b.id === savedBranchId)) {
    selectedBranchId = savedBranchId;
  } else {
    selectedBranchId = branches[0]?.id || null;
  }

  // Initialize UI
  renderHistory();
  updateUIFromBranch();
  setupEventListeners();
}

// Start when DOM is ready
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initializeApp);
} else {
  initializeApp();
}
