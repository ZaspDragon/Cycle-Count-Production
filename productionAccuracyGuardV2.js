"use strict";

(() => {
  function numberFromText(value) {
    const match = String(value ?? "").replace(/,/g, "").match(/-?\d+(?:\.\d+)?/);
    return match ? Number(match[0]) || 0 : 0;
  }

  function officialTotal() {
    return Number(state.dailyOfficialReportTotal || 0) ||
      (typeof rrGetOfficialReportTotal === "function" ? Number(rrGetOfficialReportTotal()) || 0 : 0);
  }

  function namedTotal() {
    return getAssignments().reduce((sum, assignment) => {
      if (/^(batches?|variance reports?)$/i.test(String(assignment?.name || "").trim())) return sum;
      return sum + (Number(state.employeeTotals?.[assignment.name]) || 0);
    }, 0);
  }

  function varianceTotal() {
    return Number(state.ownershipPriorityVariance || 0) ||
      (typeof pcVarianceTotal === "function" ? Number(pcVarianceTotal()) || 0 : 0);
  }

  function displayedBatchTotal() {
    const card = document.querySelector("#productionCards [data-unassigned-batches-card]");
    const displayed = numberFromText(card?.querySelector(".summary-card-top b")?.textContent);
    const official = officialTotal();
    const named = namedTotal();
    const variance = varianceTotal();
    const remainder = official > 0 ? Math.max(0, official - named - variance) : 0;
    return Math.max(displayed, Number(state.dailyBatchesTotal || 0), remainder);
  }

  function isAbsent(assignment) {
    return typeof window.isCycleCountAssignmentAbsent === "function" &&
      window.isCycleCountAssignmentAbsent(assignment);
  }

  function ensurePanel() {
    let panel = document.getElementById("productionAccuracyPanel");
    if (!panel) {
      panel = document.createElement("div");
      panel.id = "productionAccuracyPanel";
      panel.className = "message";
      panel.style.marginTop = "1rem";
      document.getElementById("resultsSection")?.appendChild(panel);
    }
    return panel;
  }

  function check() {
    if (!state.workbook) return;
    const panel = ensurePanel();
    if (!panel) return;

    const official = officialTotal();
    const named = namedTotal();
    const variance = varianceTotal();
    const batches = displayedBatchTotal();
    const absentWithCredit = getAssignments()
      .filter((assignment) => isAbsent(assignment) && (Number(state.employeeTotals?.[assignment.name]) || 0) > 0)
      .map((assignment) => `${assignment.name}: ${Number(state.employeeTotals[assignment.name]) || 0}`);

    const issues = [];
    if (absentWithCredit.length) issues.push(`Absent employees still have credit: ${absentWithCredit.join(", ")}`);
    if (official > 0 && named + variance + batches !== official) {
      issues.push(`Totals do not reconcile: named ${named} + variance ${variance} + needs review ${batches} = ${named + variance + batches}, but report total is ${official}`);
    }

    if (issues.length) {
      panel.classList.remove("success");
      panel.classList.add("error");
      panel.innerHTML = `<strong>Accuracy check failed.</strong><br>${issues.map(escapeHtml).join("<br>")}`;
    } else {
      panel.classList.remove("error");
      panel.classList.add("success");
      panel.innerHTML = `<strong>Accuracy check passed.</strong> Named employees ${named} + variance ${variance} + needs review ${batches} = ${official || named + variance + batches}.`;
    }
  }

  window.runProductionAccuracyCheck = check;
  const runLater = () => {
    window.setTimeout(check, 100);
    window.setTimeout(check, 700);
    window.setTimeout(check, 1600);
  };

  ["sourceFile", "alreadyCountedFile", "branchSelect", "matchAlreadyCountedBtn"].forEach((id) => {
    document.getElementById(id)?.addEventListener("change", runLater);
    document.getElementById(id)?.addEventListener("click", runLater);
  });

  const observer = new MutationObserver(runLater);
  const results = document.getElementById("resultsSection");
  if (results) observer.observe(results, { childList: true, subtree: true, characterData: true });
})();