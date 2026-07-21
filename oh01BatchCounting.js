"use strict";

/*
 * Treat every real detail-row batch beginning with OH01 as production.
 * Already Cycle Counted matches redistribute those OH01 rows to named
 * employees, Variance Reports, or Batches. Only unmatched OH01 rows remain
 * as additional Batches, preventing the same count from being added twice.
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
    const batch = row?.[columnIndex];
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

pcBatchesTotal = function correctedBatchesTotal() {
  const officialTotal = rrGetOfficialReportTotal();
  const namedTotal = pcNamedEmployeeTotal();
  const varianceTotal = pcVarianceTotal();
  const explicitInitialsTotal = pcExplicitBatchTotal();
  const unmatchedOh01Total = oh01UnmatchedRowTotal();
  const directBatchesTotal = explicitInitialsTotal + unmatchedOh01Total;

  if (officialTotal > 0) {
    return Math.max(
      directBatchesTotal,
      officialTotal - namedTotal - varianceTotal
    );
  }

  return directBatchesTotal;
};

rrGetBatchesTotal = pcBatchesTotal;
acGetUnassignedBatchTotal = pcBatchesTotal;

rrRemoveMetadataWarnings = function keepOh01AsCountedBatches() {
  oh01OriginalRemoveMetadataWarnings();
  state.uncreditedRows = state.uncreditedRows.filter(
    (item) => !oh01IsBatch(item?.batch)
  );
};

const oh01OriginalRenderCards = acRenderUnassignedProductionCard;
acRenderUnassignedProductionCard = function renderCorrectedOh01Cards() {
  oh01OriginalRenderCards();

  const card = $("productionCards")?.querySelector(
    "[data-unassigned-batches-card]"
  );
  if (!card) return;

  const description = card.querySelector(".summary-card-top span");
  if (description) {
    const explicitInitialsTotal = pcExplicitBatchTotal();
    const unmatchedOh01Total = oh01UnmatchedRowTotal();
    description.textContent =
      `Unassigned initials: ${explicitInitialsTotal} • ` +
      `Unmatched OH01 rows: ${unmatchedOh01Total}`;
  }
};
