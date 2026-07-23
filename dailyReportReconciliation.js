"use strict";

/*
 * Permanent daily reconciliation.
 * Reads the uploaded report's own Total Cycle Counts value every day and makes:
 * named employees + Variance Reports + Batches = official report total.
 * Nothing here is tied to a date or a hard-coded total.
 */
(() => {
  if (window.__dailyReportReconciliationInstalled) return;
  window.__dailyReportReconciliationInstalled = true;

  const numberValue = (value) => {
    const n = Number(String(value ?? "").replace(/[$,]/g, "").trim());
    return Number.isFinite(n) ? n : 0;
  };

  function officialReportTotal() {
    if (!state.workbook) return 0;
    const candidates = [];

    state.workbook.SheetNames.forEach((sheetName) => {
      const matrix = workbookMatrix(state.workbook, sheetName);

      matrix.forEach((row, rowIndex) => {
        const labels = row.map((cell) => normalizeText(cell));

        // Preferred layout: header contains "Total Cycle Counts" and a later
        // row labelled Total contains the daily value in that same column.
        const totalCycleColumn = labels.findIndex((label) =>
          label.includes("total cycle count")
        );
        if (totalCycleColumn >= 0) {
          for (let i = rowIndex + 1; i < Math.min(matrix.length, rowIndex + 30); i += 1) {
            const nextRow = matrix[i] || [];
            const isTotal = nextRow.some((cell) => normalizeText(cell) === "total");
            if (!isTotal) continue;
            const value = numberValue(nextRow[totalCycleColumn]);
            if (value > 0) candidates.push(value);
          }
        }

        // Also support reports where the Total label and total value share a row.
        if (labels.some((label) => label === "total")) {
          row.forEach((cell) => {
            const value = numberValue(cell);
            if (value > 0 && value <= 100000) candidates.push(value);
          });
        }
      });
    });

    return candidates.length ? Math.round(Math.max(...candidates)) : 0;
  }

  function varianceTotal() {
    return typeof pcVarianceTotal === "function"
      ? Number(pcVarianceTotal()) || 0
      : 0;
  }

  function namedTotal() {
    return getAssignments().reduce((sum, assignment) => {
      const name = String(assignment?.name || "").trim().toLowerCase();
      if (["batches", "batch", "variance reports", "variance report"].includes(name)) {
        return sum;
      }
      return sum + (Number(state.employeeTotals?.[assignment.name]) || 0);
    }, 0);
  }

  function batchesTotal() {
    const official = officialReportTotal();
    if (!official) {
      return typeof pcBatchesTotal === "function" ? Number(pcBatchesTotal()) || 0 : 0;
    }
    return Math.max(0, official - namedTotal() - varianceTotal());
  }

  function ensureBatchesCard(total) {
    const cards = $("productionCards");
    if (!cards) return;

    let card = cards.querySelector("[data-unassigned-batches-card]");
    if (total <= 0) {
      card?.remove();
      return;
    }

    if (!card) {
      card = document.createElement("article");
      card.className = "summary-card";
      card.dataset.unassignedBatchesCard = "true";
      cards.appendChild(card);
    }

    card.innerHTML = `
      <div class="summary-card-top">
        <div>
          <strong>Batches</strong>
          <span>Official report remainder not assigned to a named employee</span>
        </div>
        <b>${total}</b>
      </div>
      <div class="percent-row">
        <span>Support total</span>
        <small>${total} added to overall completion • no individual goal</small>
      </div>
    `;
  }

  function applyDailyReconciliation() {
    if (!state.workbook) return;

    const official = officialReportTotal();
    if (!official) return;

    const named = namedTotal();
    const variance = varianceTotal();
    const batches = Math.max(0, official - named - variance);

    // Keep every older helper pointed at the same daily calculation.
    window.rrGetOfficialReportTotal = officialReportTotal;
    window.rrGetNamedEmployeeTotal = namedTotal;
    window.rrGetBatchesTotal = batchesTotal;
    window.acGetUnassignedBatchTotal = batchesTotal;
    window.pcBatchesTotal = batchesTotal;

    ensureBatchesCard(batches);

    const kpis = Array.from(document.querySelectorAll("#kpiStrip .kpi"));
    if (kpis[0]) {
      const label = kpis[0].querySelector("span");
      const value = kpis[0].querySelector("strong");
      if (label) label.textContent = "Overall Completed Counts";
      if (value) value.textContent = String(official);
    }

    const requiredGoal = Number(kpis[2]?.querySelector("strong")?.textContent) || 0;
    if (kpis[3]) {
      const value = kpis[3].querySelector("strong");
      if (value) {
        value.textContent = requiredGoal > 0
          ? `${((official / requiredGoal) * 100).toFixed(1)}%`
          : "0.0%";
      }
    }

    const summary = $("recordSummary");
    if (summary) {
      summary.textContent =
        `${official} total production • ${variance} Variance Reports • ${batches} Batches`;
    }

    state.dailyOfficialReportTotal = official;
    state.dailyBatchesTotal = batches;
  }

  const runSoon = () => {
    window.setTimeout(applyDailyReconciliation, 0);
    window.setTimeout(applyDailyReconciliation, 150);
    window.setTimeout(applyDailyReconciliation, 600);
  };

  ["sourceFile", "alreadyCountedFile", "matchAlreadyCountedBtn", "branchSelect"].forEach((id) => {
    $(id)?.addEventListener(id === "matchAlreadyCountedBtn" ? "click" : "change", runSoon);
  });

  // Run after every normal render without recursively calling renderResults.
  const previousRenderResults = renderResults;
  renderResults = function renderResultsWithDailyReconciliation() {
    previousRenderResults();
    applyDailyReconciliation();
  };
})();
