"use strict";

/*
 * Treat every real detail-row batch beginning with OH01 as production.
 * These rows are assigned to the production person named "Batches" unless
 * another employee-specific rule has already identified the count.
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

const oh01OriginalBatchesTotal = pcBatchesTotal;
const oh01OriginalRemoveMetadataWarnings = rrRemoveMetadataWarnings;

pcBatchesTotal = function batchesIncludingOh01() {
  const officialTotal = rrGetOfficialReportTotal();
  const namedTotal = pcNamedEmployeeTotal();
  const varianceTotal = pcVarianceTotal();
  const explicitInitialsTotal = pcExplicitBatchTotal();
  const oh01Total = oh01DetailRowTotal();
  const directlyIdentifiedBatches = explicitInitialsTotal + oh01Total;

  if (officialTotal > 0) {
    return Math.max(
      directlyIdentifiedBatches,
      officialTotal - namedTotal - varianceTotal
    );
  }

  return Math.max(directlyIdentifiedBatches, oh01OriginalBatchesTotal());
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
acRenderUnassignedProductionCard = function renderOh01AwareCards() {
  oh01OriginalRenderCards();

  const card = $("productionCards")?.querySelector(
    "[data-unassigned-batches-card]"
  );
  const oh01Total = oh01DetailRowTotal();

  if (card && oh01Total > 0) {
    const description = card.querySelector(".summary-card-top span");
    if (description) {
      const existing = description.textContent.trim();
      description.textContent = existing
        ? `${existing} • OH01 rows: ${oh01Total}`
        : `OH01 rows: ${oh01Total}`;
    }
  }
};
