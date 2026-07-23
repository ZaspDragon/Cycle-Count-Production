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

  const previousRenderResults = renderResults;
  renderResults = function renderResultsWithDailyReconciliation() {
    previousRenderResults();
    applyDailyReconciliation();
  };
})();

/* Add a visible ownership-review table and load its assignment logic last. */
(() => {
  if (!document.getElementById("unassignedCountReviewSection")) {
    const results = document.getElementById("resultsSection");
    if (results) {
      results.insertAdjacentHTML("afterend", `
        <section id="unassignedCountReviewSection" class="card hidden">
          <div class="section-heading">
            <div>
              <h2>Unassigned Count Review</h2>
              <small>Assign unmatched count rows to the employee who completed them. Saved by branch and report date.</small>
            </div>
            <span id="unassignedCountReviewSummary" class="status"></span>
          </div>
          <div class="table-wrap">
            <table class="history-table">
              <thead>
                <tr><th>Item</th><th>Bin</th><th>Batch</th><th>Counts</th><th>Count Date</th><th>Assign To</th></tr>
              </thead>
              <tbody id="unassignedCountReviewBody"></tbody>
            </table>
          </div>
        </section>
      `);
    }
  }

  if (!window.__unassignedCountReviewLoader) {
    window.__unassignedCountReviewLoader = true;
    const script = document.createElement("script");
    script.src = `unassignedCountReview.js?v=20260723-1-${Date.now()}`;
    document.body.appendChild(script);
  }
})();
