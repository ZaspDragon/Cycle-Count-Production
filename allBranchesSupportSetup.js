"use strict";

/*
 * Apply the same Batches and Variance Reports setup to every company branch.
 * Support totals stay branch-local, remain visible as cards, add to overall
 * completed counts, and never add a 200-count goal.
 */
(() => {
  function branchCode() {
    return String(getSelectedBranch()?.name || "Branch").trim().toUpperCase();
  }

  function refreshUniversalSupportCards() {
    if (!state.workbook) return;

    if (typeof acRenderUnassignedProductionCard === "function") {
      acRenderUnassignedProductionCard();
    }

    const code = branchCode();
    const batchesCard = document.querySelector(
      "#productionCards [data-unassigned-batches-card]"
    );
    const varianceCard = document.querySelector(
      "#productionCards [data-variance-reports-card]"
    );

    if (batchesCard) {
      batchesCard.dataset.supportTotal = "true";
      const description = batchesCard.querySelector(".summary-card-top span");
      if (description) {
        description.textContent = `${code} unassigned batch counts included in overall completion`;
      }
      batchesCard.querySelector(".meter")?.remove();
      const row = batchesCard.querySelector(".percent-row");
      if (row) {
        row.innerHTML =
          "<span>Support total</span><small>No individual 200-count requirement</small>";
      }
    }

    if (varianceCard) {
      varianceCard.dataset.supportTotal = "true";
      const description = varianceCard.querySelector(".summary-card-top span");
      if (description) {
        description.textContent = `${code} variance-report counts included in overall completion`;
      }
      varianceCard.querySelector(".meter")?.remove();
      const row = varianceCard.querySelector(".percent-row");
      if (row) {
        row.innerHTML =
          "<span>Support total</span><small>No individual 200-count requirement</small>";
      }
    }
  }

  const previousRenderResults = renderResults;
  renderResults = function renderResultsForEveryBranch() {
    previousRenderResults();
    refreshUniversalSupportCards();
  };

  const previousSelectBranch = selectBranch;
  selectBranch = function selectBranchWithSupportSetup(branchId) {
    const selected = previousSelectBranch(branchId);
    if (selected) {
      window.setTimeout(() => {
        if (state.workbook) renderResults();
      }, 0);
    }
    return selected;
  };

  document.addEventListener("change", (event) => {
    if (event.target?.id === "branchSelect") {
      window.setTimeout(refreshUniversalSupportCards, 0);
    }
  });
})();

/* Always load the current official-total reconciliation after all wrappers. */
(() => {
  const script = document.createElement("script");
  script.src = `officialReportTotalFix.js?v=20260723-4-${Date.now()}`;
  document.body.appendChild(script);
})();
