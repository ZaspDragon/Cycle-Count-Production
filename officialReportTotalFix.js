"use strict";

(() => {
  if (window.__officialReportTotal1128Fix) return;
  window.__officialReportTotal1128Fix = true;

  function numeric(value) {
    const n = Number(String(value ?? "").replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : 0;
  }

  function workbookTotal() {
    if (!state.workbook) return 0;
    let result = 0;
    state.workbook.SheetNames.forEach((sheetName) => {
      workbookMatrix(state.workbook, sheetName).forEach((row) => {
        if (normalizeText(row?.[0]) !== "total") return;
        row.forEach((cell) => {
          const n = numeric(cell);
          if (n > result && n <= 100000) result = n;
        });
      });
    });
    return Math.round(result);
  }

  const oldGetter = rrGetOfficialReportTotal;
  rrGetOfficialReportTotal = function correctedOfficialReportTotal() {
    return workbookTotal() || oldGetter();
  };

  function refreshToOfficialTotal() {
    const total = workbookTotal();
    if (!total) return;
    const input = $("officialReportTotal");
    if (input) input.value = String(total);
    renderResults();
    if (typeof acRenderUnassignedProductionCard === "function") {
      acRenderUnassignedProductionCard();
    }
  }

  $("sourceFile")?.addEventListener("change", () => {
    window.setTimeout(refreshToOfficialTotal, 400);
  });

  $("matchAlreadyCountedBtn")?.addEventListener("click", () => {
    window.setTimeout(refreshToOfficialTotal, 100);
  });
})();
