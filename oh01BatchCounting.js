"use strict";

/*
 * OH01-only Batches rule.
 *
 * - A row counts as an OH01 batch only when its Batch value begins with OH01.
 * - A row whose Batch value is exactly "Batch" is a label/header and never counts.
 * - Already Cycle Counted matches redistribute OH01 production to employees or
 *   Variance Reports. Batches receives only the remaining OH01 production needed
 *   to reconcile the official report total.
 */

function oh01IsBatch(value) {
  const text = String(value ?? "").trim().toUpperCase();
  return /^OH01(?:$|[\s\-_/].*)/.test(text);
}

function oh01DetailRowTotal() {
  const batchSelect = $("batchColumn");
  if (!batchSelect || batchSelect.value === "" || !Array.isArray(state.rows)) {
    return 0;
  }

  const columnIndex = Number(batchSelect.value);
  return state.rows.reduce((total, row) => {
    const batch = String(row?.[columnIndex] ?? "").trim();
    if (/^batch$/i.test(batch)) return total;
    return total + (oh01IsBatch(batch) ? 1 : 0);
  }, 0);
}

function oh01MatchedLocationTotal() {
  return (alreadyCountedState.matchedRows || []).reduce(
    (sum, row) => sum + (Number(row.locationCount) || 0),
    0
  );
}

function oh01UnmatchedRowTotal() {
  return Math.max(0, oh01DetailRowTotal() - oh01MatchedLocationTotal());
}

const oh01OriginalRemoveMetadataWarnings = rrRemoveMetadataWarnings;

pcBatchesTotal = function oh01OnlyBatchesTotal() {
  const officialTotal = rrGetOfficialReportTotal();
  const namedTotal = pcNamedEmployeeTotal();
  const varianceTotal = pcVarianceTotal();
  const unmatchedOh01Total = oh01UnmatchedRowTotal();

  if (officialTotal > 0) {
    return Math.max(
      unmatchedOh01Total,
      officialTotal - namedTotal - varianceTotal
    );
  }

  return unmatchedOh01Total;
};

rrGetBatchesTotal = pcBatchesTotal;
acGetUnassignedBatchTotal = pcBatchesTotal;

rrRemoveMetadataWarnings = function keepOnlyRealOh01Rows() {
  oh01OriginalRemoveMetadataWarnings();
  state.uncreditedRows = state.uncreditedRows.filter((item) => {
    const batch = String(item?.batch ?? "").trim();
    if (/^batch$/i.test(batch)) return false;
    return !oh01IsBatch(batch);
  });
};

const oh01OriginalRenderCards = acRenderUnassignedProductionCard;
acRenderUnassignedProductionCard = function renderOh01OnlyBatchesCard() {
  oh01OriginalRenderCards();

  const card = $("productionCards")?.querySelector(
    "[data-unassigned-batches-card]"
  );
  if (!card) return;

  const description = card.querySelector(".summary-card-top span");
  if (description) {
    description.textContent =
      `OH01 production not credited to a named employee: ${pcBatchesTotal()}`;
  }
};
