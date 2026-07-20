"use strict";

/*
 * Flexible aisle assignments and date-filtered reference counts.
 * Loaded after app.js and attendance.js so the existing saved-data model stays intact.
 */
(() => {
  const REFERENCE_NAME = "already cycle counted numbers";

  function normalizeName(value) {
    return String(value || "").trim().toLowerCase();
  }

  function isReferenceAssignment(assignmentOrName) {
    const name = typeof assignmentOrName === "string"
      ? assignmentOrName
      : assignmentOrName?.name;
    return normalizeName(name) === REFERENCE_NAME;
  }

  function normalizeAisleToken(value) {
    return String(value || "")
      .trim()
      .toUpperCase()
      .replace(/\s+/g, "");
  }

  function normalizeDateValue(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      const year = value.getFullYear();
      const month = String(value.getMonth() + 1).padStart(2, "0");
      const day = String(value.getDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    }

    const text = String(value || "").trim();
    if (!text) return "";

    const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (iso) {
      return `${iso[1]}-${String(iso[2]).padStart(2, "0")}-${String(iso[3]).padStart(2, "0")}`;
    }

    const slash = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
    if (slash) {
      const year = slash[3].length === 2 ? `20${slash[3]}` : slash[3];
      return `${year}-${String(slash[1]).padStart(2, "0")}-${String(slash[2]).padStart(2, "0")}`;
    }

    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? "" : normalizeDateValue(parsed);
  }

  function displayDate(value) {
    const normalized = normalizeDateValue(value);
    if (!normalized) return "Choose date";
    const [year, month, day] = normalized.split("-");
    return `${month}/${day}/${year}`;
  }

  function expandFlexibleRange(startValue, endValue) {
    const start = normalizeAisleToken(startValue);
    const end = normalizeAisleToken(endValue || startValue);
    if (!start) return [];
    if (!end || start === end) return [start];

    if (/^[A-Z]$/.test(start) && /^[A-Z]$/.test(end) && start <= end) {
      const values = [];
      for (let code = start.charCodeAt(0); code <= end.charCodeAt(0); code += 1) {
        values.push(String.fromCharCode(code));
      }
      return values;
    }

    const startMatch = start.match(/^(.*?)(\d+)$/);
    const endMatch = end.match(/^(.*?)(\d+)$/);
    if (startMatch && endMatch && startMatch[1] === endMatch[1]) {
      const first = Number(startMatch[2]);
      const last = Number(endMatch[2]);
      if (first <= last && last - first <= 500) {
        const width = Math.max(startMatch[2].length, endMatch[2].length);
        return Array.from({ length: last - first + 1 }, (_, index) =>
          `${startMatch[1]}${String(first + index).padStart(width, "0")}`
        );
      }
    }

    return [start, end];
  }

  expandAisleRange = expandFlexibleRange;

  formatAisleRange = function formatFlexibleAisleRange(startValue, endValue) {
    const start = normalizeAisleToken(startValue);
    const end = normalizeAisleToken(endValue || startValue);
    return !end || start === end ? start : `${start}–${end}`;
  };

  parseCreatedBatch = function parseFlexibleCreatedBatch(value) {
    const text = String(value || "").trim().toUpperCase();
    const datedBatch = text.match(/^(.+?)-\d{2}-\d{2}$/);
    return datedBatch ? normalizeAisleToken(datedBatch[1]) : null;
  };

  getAssignedAisles = function getFlexibleAssignedAisles() {
    const assigned = new Set();
    getAssignments()
      .filter((assignment) => !isReferenceAssignment(assignment))
      .forEach((assignment) => {
        expandFlexibleRange(assignment.startAisle, assignment.endAisle)
          .forEach((aisle) => assigned.add(aisle));
      });
    return Array.from(assigned).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  };

  validateAssignment = function validateFlexibleAssignment(
    name,
    startAisle,
    endAisle,
    excludedAssignmentId = null
  ) {
    const branch = getSelectedBranch();
    if (!branch) return { valid: false, error: "No branch is currently selected." };

    const normalizedName = String(name || "").trim();
    if (!normalizedName) return { valid: false, error: "Employee name is required." };

    const duplicateName = branch.assignments.some(
      (assignment) => assignment.id !== excludedAssignmentId &&
        normalizeName(assignment.name) === normalizeName(normalizedName)
    );
    if (duplicateName) {
      return { valid: false, error: `${normalizedName} already exists in this branch.` };
    }

    if (isReferenceAssignment(normalizedName)) {
      const referenceDate = normalizeDateValue(startAisle);
      if (!referenceDate) {
        return { valid: false, error: "Choose the date for the already cycle counted numbers." };
      }
      return {
        valid: true,
        name: normalizedName,
        startAisle: referenceDate,
        endAisle: referenceDate,
      };
    }

    const start = normalizeAisleToken(startAisle);
    const end = normalizeAisleToken(endAisle || startAisle);
    if (!start || !end) {
      return { valid: false, error: "Starting and ending aisles are required." };
    }
    if (start.length > 30 || end.length > 30) {
      return { valid: false, error: "Assigned aisle values must be 30 characters or fewer." };
    }
    if (!/^[A-Z0-9#_-]+$/.test(start) || !/^[A-Z0-9#_-]+$/.test(end)) {
      return { valid: false, error: "Aisles may use letters, numbers, #, underscores, and hyphens." };
    }

    const proposed = new Set(expandFlexibleRange(start, end));
    for (const assignment of branch.assignments) {
      if (assignment.id === excludedAssignmentId || isReferenceAssignment(assignment)) continue;
      const overlap = expandFlexibleRange(assignment.startAisle, assignment.endAisle)
        .some((aisle) => proposed.has(aisle));
      if (overlap) {
        return { valid: false, error: `The aisle assignment overlaps with ${assignment.name}.` };
      }
    }

    return { valid: true, name: normalizedName, startAisle: start, endAisle: end };
  };

  function configureTeamModal() {
    const nameInput = $("employeeNameInput");
    const startInput = $("startAisleInput");
    const endInput = $("endAisleInput");
    if (!nameInput || !startInput || !endInput) return;

    const startLabel = startInput.closest("label")?.querySelector("span");
    const endLabel = endInput.closest("label")?.querySelector("span");
    const reference = isReferenceAssignment(nameInput.value);

    startInput.maxLength = reference ? 10 : 30;
    endInput.maxLength = reference ? 10 : 30;
    startInput.type = reference ? "date" : "text";
    endInput.type = reference ? "date" : "text";

    if (reference) {
      if (startLabel) startLabel.textContent = "Reference Date *";
      if (endLabel) endLabel.textContent = "Reference Date (same day)";
      const normalized = normalizeDateValue(startInput.value);
      startInput.value = normalized;
      endInput.value = normalized;
      endInput.disabled = true;
      endInput.required = false;
    } else {
      if (startLabel) startLabel.textContent = "Starting Aisle *";
      if (endLabel) endLabel.textContent = "Ending Aisle *";
      endInput.disabled = false;
      endInput.required = true;
    }
  }

  const originalOpenTeamMemberForm = openTeamMemberForm;
  openTeamMemberForm = function openFlexibleTeamMemberForm(assignment = null) {
    originalOpenTeamMemberForm(assignment);
    window.setTimeout(configureTeamModal, 0);
  };

  document.addEventListener("input", (event) => {
    if (event.target?.id === "employeeNameInput") configureTeamModal();
    if (event.target?.id === "startAisleInput" && isReferenceAssignment($("employeeNameInput")?.value)) {
      const endInput = $("endAisleInput");
      if (endInput) endInput.value = event.target.value;
    }
  });

  function findReferenceDateColumn() {
    if (!state.workbook) return -1;
    const sheetName = $("sourceSheet")?.value;
    if (!sheetName) return -1;
    const matrix = workbookMatrix(state.workbook, sheetName);
    const headers = matrix[state.headerIndex] || [];
    return detectColumn(headers, [
      "count date",
      "created date",
      "date counted",
      "counted date",
      "date",
    ]);
  }

  function applyReferenceDateTotals() {
    const referenceAssignments = getAssignments().filter(isReferenceAssignment);
    if (!state.workbook || referenceAssignments.length === 0) return;

    const dateColumn = findReferenceDateColumn();
    referenceAssignments.forEach((assignment) => {
      const wantedDate = normalizeDateValue(assignment.startAisle);
      const total = dateColumn < 0 || !wantedDate
        ? 0
        : state.rows.reduce((count, row) =>
            count + (normalizeDateValue(row[dateColumn]) === wantedDate ? 1 : 0),
          0);
      state.employeeTotals[assignment.name] = total;
    });
  }

  const originalCalculateProduction = calculateProduction;
  calculateProduction = function calculateFlexibleProduction() {
    originalCalculateProduction();
    applyReferenceDateTotals();
    renderResults();
  };

  const originalRenderAssignments = renderAssignments;
  renderAssignments = function renderFlexibleAssignments() {
    originalRenderAssignments();
    const assignments = getAssignments();
    const rows = Array.from(document.querySelectorAll("#teamBody tr"));

    rows.forEach((row, index) => {
      const assignment = assignments[index];
      if (!assignment || !isReferenceAssignment(assignment)) return;
      const aisleCell = row.children[1];
      if (aisleCell) aisleCell.textContent = `Date: ${displayDate(assignment.startAisle)}`;
      const attendance = row.querySelector(".attendance-select");
      if (attendance) {
        attendance.disabled = true;
        attendance.title = "Reference totals do not use attendance status.";
      }
    });
  };

  const originalRenderAssignmentGrid = renderAssignmentGrid;
  renderAssignmentGrid = function renderFlexibleAssignmentGrid() {
    originalRenderAssignmentGrid();
    const assignments = getAssignments();
    const cards = Array.from(document.querySelectorAll("#assignmentGrid > div"));
    cards.forEach((card, index) => {
      const assignment = assignments[index];
      if (!assignment || !isReferenceAssignment(assignment)) return;
      const value = card.querySelector("span");
      if (value) value.textContent = `Date: ${displayDate(assignment.startAisle)}`;
    });
  };

  const originalRenderResults = renderResults;
  renderResults = function renderFlexibleResults() {
    originalRenderResults();
    const referenceAssignments = getAssignments().filter(isReferenceAssignment);
    if (referenceAssignments.length === 0) return;

    const dateColumn = findReferenceDateColumn();
    document.querySelectorAll("#productionCards .summary-card").forEach((card) => {
      const name = card.querySelector("strong")?.textContent?.trim();
      const assignment = referenceAssignments.find((item) => item.name === name);
      if (!assignment) return;
      const breakdown = card.querySelector(".summary-card-top span");
      if (breakdown) {
        breakdown.textContent = dateColumn < 0
          ? "Date column not found in report"
          : `Count date: ${displayDate(assignment.startAisle)}`;
      }
    });
  };
})();
