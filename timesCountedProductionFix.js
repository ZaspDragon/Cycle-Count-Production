"use strict";

/*
 * Cycle Count Detail can contain fewer physical detail rows than the official
 * count total because a detail line may have Times Counted greater than one.
 * This patch makes every production path use that weight and reads the report's
 * true Total / Times Counted value for reconciliation.
 */

function tcpfNumber(value) {
  const number = Number(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(number) ? number : 0;
}

function tcpfDetailRecords(workbook = state.workbook) {
  const records = [];
  if (!workbook) return records;

  workbook.SheetNames.forEach((sheetName) => {
    const matrix = workbookMatrix(workbook, sheetName);
    let currentItem = "";
    let columns = null;

    matrix.forEach((row, rowIndex) => {
      const first = acNormalizeItem(row?.[0]);
      const firstLabel = normalizeText(row?.[0]);

      if (/^\d{4,}$/.test(first) && firstLabel !== "total") {
        currentItem = first;
      }

      const countDateColumn = row.findIndex(
        (cell) => normalizeText(cell) === "count date"
      );
      if (countDateColumn >= 0) {
        columns = {
          countDate: countDateColumn,
          bin: detectColumn(row, ["bin #", "bin"]),
          batch: detectColumn(row, ["batch"]),
          timesCounted: detectColumn(row, ["times counted"]),
        };
        return;
      }

      if (!currentItem || !columns) return;

      const bin = acNormalizeBin(row?.[columns.bin]);
      const batch = String(row?.[columns.batch] ?? "").trim();
      const countDate = row?.[columns.countDate];
      const rawWeight = tcpfNumber(row?.[columns.timesCounted]);
      const weight = Math.max(1, Math.round(rawWeight || 1));
      const looksLikeBin =
        bin &&
        !["BIN", "BIN#", "BATCH", "RANK"].includes(bin) &&
        /[A-Z]/.test(bin) &&
        /\d|CAGE/.test(bin);

      if (!countDate || !looksLikeBin || !batch || /^batch$/i.test(batch)) return;

      records.push({
        itemNumber: currentItem,
        bin,
        batch,
        weight,
        sheetName,
        rowNumber: rowIndex + 1,
      });
    });
  });

  return records;
}

function tcpfOfficialReportTotal() {
  if (!state.workbook) return 0;
  let total = 0;

  state.workbook.SheetNames.forEach((sheetName) => {
    const matrix = workbookMatrix(state.workbook, sheetName);
    matrix.forEach((row) => {
      if (normalizeText(row?.[0]) !== "total") return;
      row.forEach((value) => {
        const number = tcpfNumber(value);
        if (number > total && number <= 100000) total = number;
      });
    });
  });

  return Math.round(total);
}

// Use the explicit Total row first; retain the previous detector as fallback.
const tcpfPreviousOfficialDetector = rrDetectOfficialReportTotal;
rrDetectOfficialReportTotal = function detectTrueTimesCountedTotal() {
  return tcpfOfficialReportTotal() || tcpfPreviousOfficialDetector();
};

// Build item/bin credit using the detail line's Times Counted value, not one
// credit per unique bin. This is especially important for OH01 date batches.
acBuildDetailItemMap = function buildWeightedDetailItemMap(workbook) {
  const itemBins = new Map();

  tcpfDetailRecords(workbook).forEach((record) => {
    if (!itemBins.has(record.itemNumber)) {
      itemBins.set(record.itemNumber, new Map());
    }
    const bins = itemBins.get(record.itemNumber);
    bins.set(record.bin, (bins.get(record.bin) || 0) + record.weight);
  });

  return itemBins;
};

acMatchFiles = function matchFilesWithTimesCounted() {
  if (!alreadyCountedState.workbook || !state.workbook) {
    acSetStatus("Upload both workbooks to calculate location credit.", true);
    return;
  }

  const detailMap = acBuildDetailItemMap(state.workbook);
  const alreadyRows = acFindAlreadyCountedRows(alreadyCountedState.workbook);
  const seenItemInitials = new Set();
  const matchedRows = [];
  const unmatchedRows = [];
  let duplicates = 0;

  alreadyRows.forEach((row) => {
    const dedupeKey = `${row.itemNumber}|${row.initials}`;
    if (seenItemInitials.has(dedupeKey)) {
      duplicates += 1;
      return;
    }
    seenItemInitials.add(dedupeKey);

    const binWeights = detailMap.get(row.itemNumber);
    if (!binWeights || binWeights.size === 0) {
      unmatchedRows.push({ ...row, reason: "No locations found in Cycle Count Detail" });
      return;
    }

    const locationCount = [...binWeights.values()].reduce(
      (sum, value) => sum + (Number(value) || 0),
      0
    );

    matchedRows.push({
      ...row,
      bins: [...binWeights.keys()].sort(),
      uniqueLocationCount: binWeights.size,
      locationCount,
    });
  });

  alreadyCountedState.rows = alreadyRows;
  alreadyCountedState.matchedRows = matchedRows;
  alreadyCountedState.unmatchedRows = unmatchedRows;
  alreadyCountedState.duplicateCount = duplicates;
  alreadyCountedState.totalsByInitials = matchedRows.reduce((totals, row) => {
    totals[row.initials] = (totals[row.initials] || 0) + row.locationCount;
    return totals;
  }, {});
  alreadyCountedState.applied = true;

  acPopulateInitialsFilter();
  acApplyCreditsToProduction();
  acRenderPreview();

  const weightedTotal = matchedRows.reduce(
    (sum, row) => sum + row.locationCount,
    0
  );
  const uniqueTotal = matchedRows.reduce(
    (sum, row) => sum + (row.uniqueLocationCount || 0),
    0
  );
  acSetStatus(
    `${matchedRows.length} items matched • ${weightedTotal} cycle counts credited across ${uniqueTotal} unique locations`,
    false
  );
};

// Standard aisle batches also receive their Times Counted weight. The existing
// calculation already grants one, so only the additional amount is applied.
const tcpfPreviousCalculateProduction = calculateProduction;
calculateProduction = function calculateWeightedProduction() {
  tcpfPreviousCalculateProduction();

  const assignments = getAssignments();
  const assignedAisles = getAssignedAisles();
  let changed = false;

  tcpfDetailRecords().forEach((record) => {
    const aisle = parseCreatedBatch(record.batch);
    const extra = Math.max(0, record.weight - 1);
    if (!extra || !aisle || !assignedAisles.includes(aisle)) return;

    state.aisleTotals[aisle] = (Number(state.aisleTotals[aisle]) || 0) + extra;
    const assignment = assignments.find((item) =>
      expandAisleRange(item.startAisle, item.endAisle).includes(aisle)
    );
    if (assignment) {
      state.employeeTotals[assignment.name] =
        (Number(state.employeeTotals[assignment.name]) || 0) + extra;
    }
    changed = true;
  });

  if (changed && state.workbook) {
    renderResults();
    if (typeof acRenderUnassignedProductionCard === "function") {
      acRenderUnassignedProductionCard();
    }
  }
};

// Branch batch reconciliation must also sum the weighted detail total.
selectedBranchDetailRowTotal = function weightedSelectedBranchDetailTotal() {
  return tcpfDetailRecords().reduce(
    (total, record) =>
      total + (selectedBranchIsBatch(record.batch) ? record.weight : 0),
    0
  );
};
