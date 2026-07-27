"use strict";

(() => {
  if (window.__productionAccuracyGuardLoaded) return;
  window.__productionAccuracyGuardLoaded = true;

  function isAbsent(assignment) {
    return typeof window.isCycleCountAssignmentAbsent === "function"
      ? window.isCycleCountAssignmentAbsent(assignment)
      : false;
  }

  function getOfficialTotal() {
    return typeof rrGetOfficialReportTotal === "function"
      ? Number(rrGetOfficialReportTotal()) || 0
      : 0;
  }

  function getVarianceTotal() {
    return typeof pcVarianceTotal === "function"
      ? Number(pcVarianceTotal()) || 0
      : Number(state.ownershipPriorityVariance || 0);
  }

  function getBatchTotal() {
    return typeof pcBatchesTotal === "function"
      ? Number(pcBatchesTotal()) || 0
      : Number(state.ownershipPriorityBatches || 0);
  }

  function employeeTotal() {
    return getAssignments().reduce((sum, assignment) => {
      if (/^(batches?|variance reports?)$/i.test(String(assignment?.name || "").trim())) return sum;
      return sum + (Number(state.employeeTotals?.[assignment.name]) || 0);
    }, 0);
  }

  function ensurePanel() {
    let panel = document.getElementById("productionAccuracyPanel");
    if (panel) return panel;
    const results = document.getElementById("resultsSection");
    if (!results) return null;
    panel = document.createElement("div");
    panel.id = "productionAccuracyPanel";
    panel.className = "message";
    panel.style.marginTop = "1rem";
    results.appendChild(panel);
    return panel;
  }

  function runAccuracyCheck() {
    if (!state.workbook) return;

    const issues = [];
    const absentWithCredit = [];

    getAssignments().forEach((assignment) => {
      const total = Number(state.employeeTotals?.[assignment.name]) || 0;
      if (isAbsent(assignment) && total > 0) {
        absentWithCredit.push(`${assignment.name}: ${total}`);
      }
    });

    if (absentWithCredit.length) {
      issues.push(`Absent employees still have credit: ${absentWithCredit.join(", ")}`);
    }

    const official = getOfficialTotal();
    const named = employeeTotal();
    const variance = getVarianceTotal();
    const batches = getBatchTotal();
    const reconciled = named + variance + batches;

    if (official > 0 && reconciled !== official) {
      issues.push(`Totals do not reconcile: named ${named} + variance ${variance} + batches ${batches} = ${reconciled}, but report total is ${official}`);
    }

    const duplicateInitials = new Map();
    getAssignments().forEach((assignment) => {
      const initials = typeof acGetInitials === "function" ? acGetInitials(assignment) : assignment.initials;
      const key = String(initials || "").trim().toLowerCase();
      if (!key) return;
      if (!duplicateInitials.has(key)) duplicateInitials.set(key, []);
      duplicateInitials.get(key).push(assignment.name);
    });
    duplicateInitials.forEach((names, initials) => {
      if (names.length > 1) issues.push(`Initials ${initials.toUpperCase()} are assigned to multiple employees: ${names.join(", ")}`);
    });

    const panel = ensurePanel();
    if (!panel) return;

    if (issues.length) {
      panel.classList.remove("success");
      panel.classList.add("error");
      panel.innerHTML = `<strong>Accuracy check failed.</strong><br>${issues.map((issue) => escapeHtml(issue)).join("<br>")}`;
    } else {
      panel.classList.remove("error");
      panel.classList.add("success");
      panel.innerHTML = `<strong>Accuracy check passed.</strong> Every credited count has one owner, absent employees have zero, and totals reconcile to ${official || reconciled}.`;
    }
  }

  window.runProductionAccuracyCheck = runAccuracyCheck;

  const previousRenderResults = renderResults;
  renderResults = function renderResultsWithAccuracyGuard() {
    previousRenderResults();
    window.setTimeout(runAccuracyCheck, 0);
  };

  ["sourceFile", "alreadyCountedFile", "branchSelect", "matchAlreadyCountedBtn"].forEach((id) => {
    document.getElementById(id)?.addEventListener("change", () => window.setTimeout(runAccuracyCheck, 400));
    document.getElementById(id)?.addEventListener("click", () => window.setTimeout(runAccuracyCheck, 400));
  });
})();