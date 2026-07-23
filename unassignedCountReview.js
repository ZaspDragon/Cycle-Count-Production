"use strict";

(() => {
  const STORAGE_KEY = "cycleCountProduction.manualCountAssignments.v1";
  const reviewState = { rows: [], appliedByEmployee: {} };

  function readSaved() {
    try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}") || {}; }
    catch { return {}; }
  }

  function writeSaved(value) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(value));
  }

  function reportKey() {
    const branch = String(getSelectedBranch()?.name || "branch").trim().toUpperCase();
    const date = typeof acGetReportCountDate === "function" ? acGetReportCountDate(state.workbook) : null;
    const dateKey = date instanceof Date && !Number.isNaN(date.getTime())
      ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`
      : "undated";
    return `${branch}|${dateKey}`;
  }

  function detailRows() {
    const rows = [];
    if (!state.workbook) return rows;

    state.workbook.SheetNames.forEach((sheetName) => {
      const matrix = workbookMatrix(state.workbook, sheetName);
      let currentItem = "";
      let columns = null;

      matrix.forEach((row, index) => {
        const first = acNormalizeItem(row?.[0]);
        if (/^\d{4,}$/.test(first) && normalizeText(row?.[0]) !== "total") currentItem = first;

        const countDate = detectColumn(row, ["count date"]);
        if (countDate >= 0) {
          columns = {
            countDate,
            bin: detectColumn(row, ["bin #", "bin"]),
            batch: detectColumn(row, ["batch"]),
            times: detectColumn(row, ["times counted"]),
          };
          return;
        }

        if (!currentItem || !columns || columns.bin < 0 || columns.batch < 0) return;
        const bin = acNormalizeBin(row?.[columns.bin]);
        const batch = String(row?.[columns.batch] ?? "").trim();
        const dateValue = row?.[columns.countDate];
        const count = Math.max(1, Math.round(Number(row?.[columns.times]) || 1));
        if (!bin || !batch || /^batch$/i.test(batch) || !dateValue) return;
        if (!/[A-Z]/.test(bin) || !(/\d|CAGE/.test(bin))) return;

        const id = `${currentItem}|${bin}|${batch}|${sheetName}|${index + 1}`;
        rows.push({ id, itemNumber: currentItem, bin, batch, countDate: dateValue, count });
      });
    });
    return rows;
  }

  function initialsByItem() {
    const map = new Map();
    (alreadyCountedState.rows || []).forEach((row) => {
      if (!map.has(row.itemNumber)) map.set(row.itemNumber, new Set());
      map.get(row.itemNumber).add(acNormalizeInitials(row.initials));
    });
    return map;
  }

  function aisleCandidates(value) {
    const text = String(value || "").trim().toUpperCase();
    if (!text || /^(OH\d+|BATCH|DW)\b/.test(text)) return [];

    const firstToken = text.split(/[-\s_/]+/).find(Boolean) || "";
    const candidates = [firstToken];
    const leadingLetters = firstToken.match(/^[A-Z]+/)?.[0];
    if (leadingLetters && !candidates.includes(leadingLetters)) candidates.push(leadingLetters);
    if (leadingLetters?.length > 1 && !candidates.includes(leadingLetters[0])) {
      candidates.push(leadingLetters[0]);
    }
    return candidates;
  }

  function employeeForAisle(aisle) {
    const normalized = String(aisle || "").trim().toUpperCase();
    if (!normalized) return "";

    const matches = getAssignments().filter((assignment) => {
      if (/^(batches?|variance reports?)$/i.test(String(assignment.name || "").trim())) return false;
      const assigned = typeof expandAisleRange === "function"
        ? expandAisleRange(assignment.startAisle, assignment.endAisle)
        : [assignment.startAisle, assignment.endAisle];
      return assigned.map((value) => String(value || "").trim().toUpperCase()).includes(normalized);
    });

    return matches.length === 1 ? matches[0].name : "";
  }

  function automaticEmployee(row) {
    for (const aisle of aisleCandidates(row.batch)) {
      const employee = employeeForAisle(aisle);
      if (employee) return { employee, source: "batch" };
    }
    for (const aisle of aisleCandidates(row.bin)) {
      const employee = employeeForAisle(aisle);
      if (employee) return { employee, source: "bin" };
    }
    return { employee: "", source: "" };
  }

  function buildRows() {
    const itemInitials = initialsByItem();
    const saved = readSaved()[reportKey()] || {};
    reviewState.rows = detailRows()
      .filter((row) => !itemInitials.has(row.itemNumber))
      .map((row) => {
        const automatic = automaticEmployee(row);
        const savedEmployee = saved[row.id] || "";
        return {
          ...row,
          assignedEmployee: savedEmployee || automatic.employee,
          assignmentSource: savedEmployee ? "manual" : automatic.source,
        };
      });
  }

  function manualTotals() {
    return reviewState.rows.reduce((totals, row) => {
      if (!row.assignedEmployee) return totals;
      totals[row.assignedEmployee] = (totals[row.assignedEmployee] || 0) + row.count;
      return totals;
    }, {});
  }

  function applyManualTotals() {
    getAssignments().forEach((assignment) => {
      const previous = Number(reviewState.appliedByEmployee[assignment.name] || 0);
      if (previous) state.employeeTotals[assignment.name] = Math.max(0, Number(state.employeeTotals[assignment.name] || 0) - previous);
    });
    const next = manualTotals();
    getAssignments().forEach((assignment) => {
      const add = Number(next[assignment.name] || 0);
      if (add) state.employeeTotals[assignment.name] = Number(state.employeeTotals[assignment.name] || 0) + add;
    });
    reviewState.appliedByEmployee = next;
  }

  function persistAssignments() {
    const all = readSaved();
    all[reportKey()] = Object.fromEntries(
      reviewState.rows
        .filter((row) => row.assignmentSource === "manual" && row.assignedEmployee)
        .map((row) => [row.id, row.assignedEmployee])
    );
    writeSaved(all);
  }

  function renderReview() {
    const section = $("unassignedCountReviewSection");
    const body = $("unassignedCountReviewBody");
    const summary = $("unassignedCountReviewSummary");
    if (!section || !body || !summary) return;

    const unresolvedRows = reviewState.rows.filter((row) => !row.assignedEmployee);
    const unresolvedCounts = unresolvedRows.reduce((sum, row) => sum + row.count, 0);
    const autoRows = reviewState.rows.filter((row) => row.assignedEmployee && row.assignmentSource !== "manual");
    const autoCounts = autoRows.reduce((sum, row) => sum + row.count, 0);

    summary.textContent = `${autoRows.length} rows / ${autoCounts} counts auto-assigned by aisle • ${unresolvedRows.length} rows / ${unresolvedCounts} counts need review`;
    section.classList.toggle("hidden", unresolvedRows.length === 0);

    const employeeOptions = getAssignments()
      .filter((assignment) => !/^(batches?|variance reports?)$/i.test(assignment.name.trim()))
      .map((assignment) => `<option value="${escapeHtml(assignment.name)}">${escapeHtml(assignment.name)}</option>`)
      .join("");

    body.innerHTML = unresolvedRows.slice(0, 1000).map((row) => `
      <tr>
        <td>${escapeHtml(row.itemNumber)}</td>
        <td>${escapeHtml(row.bin)}</td>
        <td>${escapeHtml(row.batch)}</td>
        <td>${row.count}</td>
        <td>${escapeHtml(String(row.countDate))}</td>
        <td>
          <select data-review-id="${escapeHtml(row.id)}">
            <option value="">Unassigned</option>
            ${employeeOptions}
          </select>
        </td>
      </tr>
    `).join("") || '<tr><td colspan="6">Every row was assigned automatically by aisle.</td></tr>';

    body.querySelectorAll("select[data-review-id]").forEach((select) => {
      const row = reviewState.rows.find((item) => item.id === select.dataset.reviewId);
      if (row) select.value = row.assignedEmployee || "";
      select.addEventListener("change", () => {
        if (!row) return;
        row.assignedEmployee = select.value;
        row.assignmentSource = select.value ? "manual" : "";
        persistAssignments();
        applyManualTotals();
        renderResults();
        if (typeof acRenderUnassignedProductionCard === "function") acRenderUnassignedProductionCard();
        renderReview();
      });
    });
  }

  function refreshReview() {
    if (!state.workbook || !alreadyCountedState.applied) return;
    buildRows();
    applyManualTotals();
    renderResults();
    if (typeof acRenderUnassignedProductionCard === "function") acRenderUnassignedProductionCard();
    renderReview();
  }

  const previousMatch = acMatchFiles;
  acMatchFiles = function matchWithUnassignedReview() {
    previousMatch();
    window.setTimeout(refreshReview, 0);
  };

  document.addEventListener("change", (event) => {
    if (event.target?.id === "branchSelect") window.setTimeout(refreshReview, 0);
  });
})();