"use strict";

(() => {
  if (window.__officialReportTotal1128FixV2) return;
  window.__officialReportTotal1128FixV2 = true;

  function numeric(value) {
    const number = Number(String(value ?? "").replace(/,/g, "").trim());
    return Number.isFinite(number) ? number : 0;
  }

  function workbookTotal() {
    if (!state.workbook) return 0;
    let result = 0;

    state.workbook.SheetNames.forEach((sheetName) => {
      workbookMatrix(state.workbook, sheetName).forEach((row) => {
        const labels = row.map((cell) => normalizeText(cell));
        const isTotalRow = labels.some((label) => label === "total");
        if (!isTotalRow) return;

        row.forEach((cell) => {
          const number = numeric(cell);
          if (number > result && number <= 100000) result = number;
        });
      });
    });

    return Math.round(result);
  }

  function namedEmployeeTotal() {
    return Object.values(state.employeeTotals || {}).reduce(
      (sum, value) => sum + (Number(value) || 0),
      0
    );
  }

  function varianceTotal() {
    if (typeof pcVarianceTotal === "function") {
      return Number(pcVarianceTotal()) || 0;
    }
    return 0;
  }

  function reconciledBatchesTotal() {
    const official = workbookTotal();
    const named = namedEmployeeTotal();
    const variances = varianceTotal();
    if (official > 0) return Math.max(0, official - named - variances);
    return 0;
  }

  const previousOfficialGetter = rrGetOfficialReportTotal;
  rrGetOfficialReportTotal = function correctedOfficialReportTotal() {
    return workbookTotal() || previousOfficialGetter();
  };

  rrGetBatchesTotal = reconciledBatchesTotal;
  acGetUnassignedBatchTotal = reconciledBatchesTotal;
  if (typeof pcBatchesTotal !== "undefined") {
    pcBatchesTotal = reconciledBatchesTotal;
  }

  function refreshToOfficialTotal() {
    const official = workbookTotal();
    if (!official) return;

    const input = $("officialReportTotal");
    if (input) input.value = String(official);

    renderResults();
    if (typeof acRenderUnassignedProductionCard === "function") {
      acRenderUnassignedProductionCard();
    }

    const summary = $("recordSummary");
    if (summary) {
      summary.textContent =
        `${official} total production • ${varianceTotal()} Variance Reports • ` +
        `${reconciledBatchesTotal()} Batches`;
    }

    const firstKpi = $("kpiStrip")?.querySelector(".kpi strong");
    if (firstKpi) firstKpi.textContent = String(official);
  }

  $("sourceFile")?.addEventListener("change", () => {
    window.setTimeout(refreshToOfficialTotal, 500);
  });

  $("alreadyCountedFile")?.addEventListener("change", () => {
    window.setTimeout(refreshToOfficialTotal, 500);
  });

  $("matchAlreadyCountedBtn")?.addEventListener("click", () => {
    window.setTimeout(refreshToOfficialTotal, 150);
  });
})();
