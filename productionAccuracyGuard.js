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
      : Number(state.dailyOfficialReportTotal || 0);
  }

  function getVarianceTotal() {
    return Number(state.ownershipPriorityVariance || 0) ||
      (typeof pcVarianceTotal === "function" ? Number(pcVarianceTotal()) || 0 : 0);
  }

  function employeeTotal() {
    return getAssignments().reduce((sum, assignment) => {
      if (/^(batches?|variance reports?)$/i.test(String(assignment?.name || "").trim())) return sum;
      return sum + (Number(state.employeeTotals?.[assignment.name]) || 0);
    }, 0);
  }

  function getBatchTotal(official, named, variance) {
    const detected = Math.max(
      Number(state.ownershipPriorityBatches || 0),
      Number(state.dailyBatchesTotal || 0),
      typeof pcBatchesTotal === "function" ? Number(pcBatchesTotal()) || 0 : 0,
      typeof rrGetBatchesTotal === "function" ? Number(rrGetBatchesTotal()) || 0 : 0
    );
    const requiredRemainder = official > 0 ? Math.max(0, official - named - variance) : 0;
    return Math.max(detected, requiredRemainder);
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
      if (isAbsent(assignment) && total > 0) absentWithCredit.push(`${assignment.name}: ${total}`);
    });

    if (absentWithCredit.length) {
      issues.push(`Absent employees still have credit: ${absentWithCredit.join(", ")}`);
    }

    const official = getOfficialTotal();
    const named = employeeTotal();
    const variance = getVarianceTotal();
    const batches = getBatchTotal(official, named, variance);
    const reconciled = named + variance + batches;

    state.dailyBatchesTotal = batches;

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
      panel.innerHTML = `<strong>Accuracy check passed.</strong> Named employees ${named} + variance ${variance} + needs-review ${batches} = ${official || reconciled}. No absent employee has credit.`;
    }
  }

  window.runProductionAccuracyCheck = runAccuracyCheck;

  const previousRenderResults = renderResults;
  renderResults = function renderResultsWithAccuracyGuard() {
    previousRenderResults();
    window.setTimeout(runAccuracyCheck, 50);
    window.setTimeout(runAccuracyCheck, 500);
  };

  ["sourceFile", "alreadyCountedFile", "branchSelect", "matchAlreadyCountedBtn"].forEach((id) => {
    document.getElementById(id)?.addEventListener("change", () => window.setTimeout(runAccuracyCheck, 700));
    document.getElementById(id)?.addEventListener("click", () => window.setTimeout(runAccuracyCheck, 700));
  });
})();