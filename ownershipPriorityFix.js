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

  function isAbsent(assignment) {
    return typeof window.isCycleCountAssignmentAbsent === "function"
      ? window.isCycleCountAssignmentAbsent(assignment)
      : false;
  }

  function initialsOwners() {
    const employeeByInitials = new Map();
    employees().forEach((assignment) => {
      const initials = normalizeInitials(typeof acGetInitials === "function" ? acGetInitials(assignment) : assignment.initials);
      if (initials) employeeByInitials.set(initials, assignment);
    });

    const ownerByItem = new Map();
    (alreadyCountedState?.rows || []).forEach((row) => {
      const item = typeof acNormalizeItem === "function" ? acNormalizeItem(row.itemNumber) : String(row.itemNumber ?? "").trim();
      const initials = normalizeInitials(row.initials);
      if (!item || !initials) return;
      const assignment = employeeByInitials.get(initials) || null;
      ownerByItem.set(item, {
        initials,
        employee: assignment?.name || "",
        assignment,
      });
    });
    return ownerByItem;
  }

  function aisleOwner(value) {
    const text = normalize(value);
    if (!text || /^(OH\d+|BATCH|DW)\b/.test(text)) return null;
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
      if (matches.length === 1) return matches[0];
    }
    return null;
  }

  function detailRows() {
    const output = [];
    const seenLocations = new Set();
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
          };
          return;
        }

        if (!currentItem || !columns || columns.bin < 0) return;
        const bin = typeof acNormalizeBin === "function" ? acNormalizeBin(row?.[columns.bin]) : normalize(row?.[columns.bin]);
        const batch = columns.batch >= 0 ? String(row?.[columns.batch] ?? "").trim() : "";
        const dateValue = columns.countDate >= 0 ? row?.[columns.countDate] : null;
        if (!bin || !/[A-Z]/.test(bin) || !(/\d|CAGE/.test(bin))) return;

        const locationKey = `${currentItem}|${bin}`;
        if (seenLocations.has(locationKey)) return;
        seenLocations.add(locationKey);
        output.push({ item: currentItem, bin, batch, count: 1, dateValue });
      });
    });

    return output;
  }

  function renderOwnershipBreakdown() {
    const cards = Array.from(document.querySelectorAll("#productionCards .summary-card"));
    const assignments = employees();
    assignments.forEach((assignment, index) => {
      const card = cards[index];
      if (!card) return;
      const source = state.ownershipPrioritySources?.[assignment.name] || { aisle: 0, alreadyCounted: 0, manual: 0 };
      const initials = normalizeInitials(typeof acGetInitials === "function" ? acGetInitials(assignment) : assignment.initials).toUpperCase();
      const label = card.querySelector(".summary-card-top span");
      if (label) {
        const manualText = source.manual ? ` • Manual: ${source.manual}` : "";
        label.textContent = `Aisle-owned: ${source.aisle || 0} • Already Counted${initials ? ` (${initials})` : ""}: ${source.alreadyCounted || 0}${manualText}`;
      }
    });
  }

  function recalculateOwnership() {
    if (!state.workbook || !alreadyCountedState?.applied) return;

    const ownerByItem = initialsOwners();
    const totals = Object.fromEntries(employees().map((assignment) => [assignment.name, 0]));
    const sources = Object.fromEntries(employees().map((assignment) => [assignment.name, { aisle: 0, alreadyCounted: 0, manual: 0 }]));
    let variance = 0;
    let batches = 0;

    detailRows().forEach((row) => {
      const manualEmployee = typeof window.getManualCycleCountOwner === "function"
        ? window.getManualCycleCountOwner(row)
        : "";
      if (manualEmployee && Object.prototype.hasOwnProperty.call(totals, manualEmployee)) {
        totals[manualEmployee] += row.count;
        sources[manualEmployee].manual += row.count;
        return;
      }

      const listedOwner = ownerByItem.get(row.item);
      if (listedOwner) {
        if (listedOwner.initials === "dw") variance += row.count;
        else if (listedOwner.assignment && !isAbsent(listedOwner.assignment)) {
          totals[listedOwner.employee] += row.count;
          sources[listedOwner.employee].alreadyCounted += row.count;
        } else {
          batches += row.count;
        }
        return;
      }

      const assignment = aisleOwner(row.batch) || aisleOwner(row.bin);
      if (assignment && !isAbsent(assignment)) {
        totals[assignment.name] += row.count;
        sources[assignment.name].aisle += row.count;
      } else {
        batches += row.count;
      }
    });

    employees().forEach((assignment) => {
      state.employeeTotals[assignment.name] = Number(totals[assignment.name] || 0);
      assignment.__alreadyCountedApplied = Number(sources[assignment.name]?.alreadyCounted || 0);
      assignment.__ownershipPriorityTotal = Number(totals[assignment.name] || 0);
    });

    state.ownershipPriorityTotals = totals;
    state.ownershipPrioritySources = sources;
    state.ownershipPriorityVariance = variance;
    state.ownershipPriorityBatches = batches;
    state.dailyBatchesTotal = batches;

    window.pcVarianceTotal = () => Number(state.ownershipPriorityVariance || 0);
    window.pcBatchesTotal = () => Number(state.ownershipPriorityBatches || 0);
    window.rrGetBatchesTotal = window.pcBatchesTotal;
    window.acGetUnassignedBatchTotal = window.pcBatchesTotal;

    renderResults();
    window.setTimeout(renderOwnershipBreakdown, 0);
    if (typeof acRenderUnassignedProductionCard === "function") acRenderUnassignedProductionCard();
  }

  window.recalculateCycleCountOwnership = recalculateOwnership;

  const previousMatch = acMatchFiles;
  acMatchFiles = function matchFilesWithFinalOwnership() {
    previousMatch();
    window.setTimeout(recalculateOwnership, 0);
    window.setTimeout(recalculateOwnership, 250);
  };
})();