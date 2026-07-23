"use strict";

(() => {
  if (window.__ownershipPriorityFixInstalled) return;
  window.__ownershipPriorityFixInstalled = true;

  const SUPPORT = /^(batches?|variance reports?)$/i;
  const normalize = (value) => String(value ?? "").trim().toUpperCase();
  const normalizeInitials = (value) => typeof acNormalizeInitials === "function"
    ? acNormalizeInitials(value)
    : String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");

  function employees() {
    return getAssignments().filter((assignment) => !SUPPORT.test(String(assignment?.name || "").trim()));
  }

  function initialsOwners() {
    const employeeByInitials = new Map();
    employees().forEach((assignment) => {
      const initials = normalizeInitials(typeof acGetInitials === "function" ? acGetInitials(assignment) : assignment.initials);
      if (initials) employeeByInitials.set(initials, assignment.name);
    });

    const ownerByItem = new Map();
    (alreadyCountedState?.rows || []).forEach((row) => {
      const item = typeof acNormalizeItem === "function" ? acNormalizeItem(row.itemNumber) : String(row.itemNumber ?? "").trim();
      const initials = normalizeInitials(row.initials);
      if (!item || !initials || ownerByItem.has(item)) return;
      ownerByItem.set(item, { initials, employee: employeeByInitials.get(initials) || "" });
    });
    return ownerByItem;
  }

  function aisleOwner(value) {
    const text = normalize(value);
    if (!text || /^(OH\d+|BATCH|DW)\b/.test(text)) return "";
    const token = text.split(/[-\s_/]+/).find(Boolean) || "";
    const candidates = [token];
    const letters = token.match(/^[A-Z]+/)?.[0] || "";
    if (letters && !candidates.includes(letters)) candidates.push(letters);
    if (letters.length > 1 && !candidates.includes(letters[0])) candidates.push(letters[0]);

    for (const candidate of candidates) {
      const matches = employees().filter((assignment) => {
        const range = typeof expandAisleRange === "function"
          ? expandAisleRange(assignment.startAisle, assignment.endAisle)
          : [assignment.startAisle, assignment.endAisle];
        return range.map(normalize).includes(normalize(candidate));
      });
      if (matches.length === 1) return matches[0].name;
    }
    return "";
  }

  function detailRows() {
    const output = [];
    if (!state.workbook) return output;

    state.workbook.SheetNames.forEach((sheetName) => {
      const matrix = workbookMatrix(state.workbook, sheetName);
      let currentItem = "";
      let columns = null;

      matrix.forEach((row) => {
        const first = typeof acNormalizeItem === "function" ? acNormalizeItem(row?.[0]) : String(row?.[0] ?? "").trim();
        if (/^\d{4,}$/.test(first) && normalizeText(row?.[0]) !== "total") currentItem = first;

        const countDateColumn = detectColumn(row, ["count date"]);
        if (countDateColumn >= 0) {
          columns = {
            countDate: countDateColumn,
            bin: detectColumn(row, ["bin #", "bin"]),
            batch: detectColumn(row, ["batch"]),
            times: detectColumn(row, ["times counted"]),
          };
          return;
        }

        if (!currentItem || !columns || columns.bin < 0 || columns.batch < 0) return;
        const bin = typeof acNormalizeBin === "function" ? acNormalizeBin(row?.[columns.bin]) : normalize(row?.[columns.bin]);
        const batch = String(row?.[columns.batch] ?? "").trim();
        const dateValue = row?.[columns.countDate];
        const count = Math.max(1, Math.round(Number(row?.[columns.times]) || 1));
        if (!bin || !dateValue || !/[A-Z]/.test(bin) || !(/\d|CAGE/.test(bin))) return;
        output.push({ item: currentItem, bin, batch, count });
      });
    });
    return output;
  }

  function recalculateOwnership() {
    if (!state.workbook || !alreadyCountedState?.applied) return;

    const ownerByItem = initialsOwners();
    const totals = Object.fromEntries(employees().map((assignment) => [assignment.name, 0]));
    let variance = 0;
    let batches = 0;

    detailRows().forEach((row) => {
      const listedOwner = ownerByItem.get(row.item);
      if (listedOwner) {
        if (listedOwner.initials === "dw") variance += row.count;
        else if (listedOwner.employee) totals[listedOwner.employee] += row.count;
        else batches += row.count;
        return;
      }

      const employee = aisleOwner(row.batch) || aisleOwner(row.bin);
      if (employee) totals[employee] += row.count;
      else batches += row.count;
    });

    employees().forEach((assignment) => {
      state.employeeTotals[assignment.name] = Number(totals[assignment.name] || 0);
      assignment.__alreadyCountedApplied = 0;
      assignment.__ownershipPriorityTotal = Number(totals[assignment.name] || 0);
    });

    state.ownershipPriorityTotals = totals;
    state.ownershipPriorityVariance = variance;
    state.ownershipPriorityBatches = batches;
    state.dailyBatchesTotal = batches;

    window.pcVarianceTotal = () => Number(state.ownershipPriorityVariance || 0);
    window.pcBatchesTotal = () => Number(state.ownershipPriorityBatches || 0);
    window.rrGetBatchesTotal = window.pcBatchesTotal;
    window.acGetUnassignedBatchTotal = window.pcBatchesTotal;

    renderResults();
    if (typeof acRenderUnassignedProductionCard === "function") acRenderUnassignedProductionCard();
  }

  const previousMatch = acMatchFiles;
  acMatchFiles = function matchFilesWithFinalOwnership() {
    previousMatch();
    window.setTimeout(recalculateOwnership, 0);
    window.setTimeout(recalculateOwnership, 250);
  };
})();