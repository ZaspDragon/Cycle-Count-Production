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

    // Never silently let the last duplicate item win. Conflicting initials for
    // one item are held for aisle ownership/review instead.
    const initialsByItem = new Map();
    (alreadyCountedState?.rows || []).forEach((row) => {
      const item = typeof acNormalizeItem === "function" ? acNormalizeItem(row.itemNumber) : String(row.itemNumber ?? "").trim();
      const initials = normalizeInitials(row.initials);
      if (!item || !initials) return;
      if (!initialsByItem.has(item)) initialsByItem.set(item, new Set());
      initialsByItem.get(item).add(initials);
    });

    const ownerByItem = new Map();
    initialsByItem.forEach((initialsSet, item) => {
      if (initialsSet.size !== 1) return;
      const initials = Array.from(initialsSet)[0];
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

  function specialLocationOwner(value) {
    const token = normalize(value).split(/[-\s_/]+/).find(Boolean) || "";
    const letters = token.match(/^[A-Z]+/)?.[0] || "";
    if (!/^R(?:[A-Z])?$/.test(letters)) return null;

    const april = employees().find(
      (assignment) => String(assignment?.name || "").trim().toLowerCase() === "april"
    );
    if (april) return april;

    const rOwners = employees().filter((assignment) => {
      const range = typeof expandAisleRange === "function"
        ? expandAisleRange(assignment.startAisle, assignment.endAisle)
        : [assignment.startAisle, assignment.endAisle];
      return range.map(normalize).some((aisle) => aisle === "R" || /^R[A-Z]$/.test(aisle));
    });
    return rOwners.length === 1 ? rOwners[0] : null;
  }

  function detailRows() {
    const output = [];
    if (!state.workbook) return output;

    state.workbook.SheetNames.forEach((sheetName) => {
      const matrix = workbookMatrix(state.workbook, sheetName);
      let currentItem = "";
      let columns = null;

      matrix.forEach((row, rowIndex) => {
        const first = typeof acNormalizeItem === "function" ? acNormalizeItem(row?.[0]) : String(row?.[0] ?? "").trim();
        if (/^\d{4,}$/.test(first) && normalizeText(row?.[0]) !== "total") currentItem = first;

        const countDateColumn = detectColumn(row, ["count date"]);
        if (countDateColumn >= 0) {
          columns = {
            countDate: countDateColumn,
            bin: detectColumn(row, ["bin #", "bin"]),
            batch: detectColumn(row, ["batch"]),
            timesCounted: detectColumn(row, ["times counted"]),
          };
          return;
        }

        if (!currentItem || !columns || columns.bin < 0) return;
        const bin = typeof acNormalizeBin === "function" ? acNormalizeBin(row?.[columns.bin]) : normalize(row?.[columns.bin]);
        const batch = columns.batch >= 0 ? String(row?.[columns.batch] ?? "").trim() : "";
        const dateValue = columns.countDate >= 0 ? row?.[columns.countDate] : null;
        if (!bin || !/[A-Z]/.test(bin) || !(/\d|CAGE/.test(bin))) return;

        const rawCount = columns.timesCounted >= 0
          ? Number(String(row?.[columns.timesCounted] ?? "").replace(/,/g, "").trim())
          : 1;
        const count = Number.isFinite(rawCount) && rawCount > 0 ? Math.round(rawCount) : 1;
        output.push({
          id: `${sheetName}|${rowIndex + 1}|${currentItem}|${bin}|${batch}`,
          item: currentItem,
          bin,
          batch,
          count,
          dateValue,
        });
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
    const auditRows = [];

    detailRows().forEach((row) => {
      const manualEmployee = typeof window.getManualCycleCountOwner === "function"
        ? window.getManualCycleCountOwner(row)
        : "";
      if (manualEmployee && Object.prototype.hasOwnProperty.call(totals, manualEmployee)) {
        totals[manualEmployee] += row.count;
        sources[manualEmployee].manual += row.count;
        auditRows.push({ ...row, employee: manualEmployee, reason: "manual override" });
        return;
      }

      // R, RA, RB, RC, RD, RE, RF and RG are location-owned by April.
      // This explicit warehouse rule intentionally outranks item initials.
      const specialOwner = specialLocationOwner(row.bin);
      if (specialOwner && !isAbsent(specialOwner)) {
        totals[specialOwner.name] += row.count;
        sources[specialOwner.name].aisle += row.count;
        auditRows.push({ ...row, employee: specialOwner.name, reason: "R-location rule" });
        return;
      }

      const listedOwner = ownerByItem.get(row.item);
      if (listedOwner) {
        if (listedOwner.initials === "dw") {
          variance += row.count;
          auditRows.push({ ...row, employee: "Variance Reports", reason: "DW initials" });
        } else if (listedOwner.assignment && !isAbsent(listedOwner.assignment)) {
          totals[listedOwner.employee] += row.count;
          sources[listedOwner.employee].alreadyCounted += row.count;
          auditRows.push({ ...row, employee: listedOwner.employee, reason: `item initials ${listedOwner.initials.toUpperCase()}` });
        } else {
          batches += row.count;
          auditRows.push({ ...row, employee: "", reason: "initials not mapped to an active employee" });
        }
        return;
      }

      const assignment = aisleOwner(row.batch) || aisleOwner(row.bin);
      if (assignment && !isAbsent(assignment)) {
        totals[assignment.name] += row.count;
        sources[assignment.name].aisle += row.count;
        auditRows.push({ ...row, employee: assignment.name, reason: "aisle assignment" });
      } else {
        batches += row.count;
        auditRows.push({ ...row, employee: "", reason: "needs review" });
      }
    });

    employees().forEach((assignment) => {
      state.employeeTotals[assignment.name] = Number(totals[assignment.name] || 0);
      assignment.__alreadyCountedApplied = Number(sources[assignment.name]?.alreadyCounted || 0);
      assignment.__ownershipPriorityTotal = Number(totals[assignment.name] || 0);
    });

    state.ownershipAuditRows = auditRows;
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