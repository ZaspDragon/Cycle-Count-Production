"use strict";

/*
 * Prevent OH01 detail rows from being counted twice.
 * Already Cycle Counted matches redistribute OH01 rows to employees,
 * Variance Reports, or Batches. Only unmatched OH01 rows remain as extra Batches.
 */

function podcMatchedLocationTotal() {
  return (alreadyCountedState.matchedRows || []).reduce(
    (sum, row) => sum + (Number(row.locationCount) || 0),
    0
  );
}

function podcUnmatchedOh01Total() {
  return Math.max(0, oh01DetailRowTotal() - podcMatchedLocationTotal());
}

pcBatchesTotal = function correctedBatchesTotal() {
  const explicitUnassignedInitials = pcExplicitBatchTotal();
  const unmatchedOh01 = podcUnmatchedOh01Total();
  const officialTotal = rrGetOfficialReportTotal();
  const namedTotal = pcNamedEmployeeTotal();
  const varianceTotal = pcVarianceTotal();
  const directTotal = explicitUnassignedInitials + unmatchedOh01;

  if (officialTotal > 0) {
    return Math.max(
      directTotal,
      officialTotal - namedTotal - varianceTotal
    );
  }

  return directTotal;
};

rrGetBatchesTotal = pcBatchesTotal;
acGetUnassignedBatchTotal = pcBatchesTotal;

const podcOriginalRenderCards = acRenderUnassignedProductionCard;
acRenderUnassignedProductionCard = function renderCorrectedBatchesCard() {
  podcOriginalRenderCards();

  const card = $("productionCards")?.querySelector(
    "[data-unassigned-batches-card]"
  );
  if (!card) return;

  const description = card.querySelector(".summary-card-top span");
  if (description) {
    const explicit = pcExplicitBatchTotal();
    const unmatched = podcUnmatchedOh01Total();
    description.textContent = `Unassigned initials: ${explicit} • Unmatched OH01 rows: ${unmatched}`;
  }
};
