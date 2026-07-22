"use strict";

/*
 * Batches and Variance Reports are support totals.
 * Their counts contribute to the overall completed cycle-count total, but they
 * do not carry an individual 200-count production goal and do not increase the
 * team's required goal.
 */
(() => {
  const SUPPORT_NAMES = new Set([
    "batches",
    "batch",
    "variance reports",
    "variance report",
  ]);
  const REFERENCE_NAME = "already cycle counted numbers";
  const DEFAULT_GOAL = 200;

  function normalizeName(value) {
    return String(value || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
  }

  function isSupportAssignment(assignmentOrName) {
    const name = typeof assignmentOrName === "string"
      ? assignmentOrName
      : assignmentOrName?.name;
    return SUPPORT_NAMES.has(normalizeName(name));
  }

  function isReferenceAssignment(assignmentOrName) {
    const name = typeof assignmentOrName === "string"
      ? assignmentOrName
      : assignmentOrName?.name;
    return normalizeName(name) === REFERENCE_NAME;
  }

  function isGoalAssignment(assignment) {
    return Boolean(
      assignment &&
      !isSupportAssignment(assignment) &&
      !isReferenceAssignment(assignment)
    );
  }

  function goalFor(assignment) {
    if (!isGoalAssignment(assignment)) return 0;
    const value = Number(assignment?.dailyGoal);
    return Number.isFinite(value) && value > 0 ? Math.round(value) : DEFAULT_GOAL;
  }

  function isAbsent(assignment) {
    const attendance = typeof getAttendanceForAssignment === "function"
      ? getAttendanceForAssignment(assignment)
      : null;
    return attendance?.status === "absent";
  }

  function applySupportGoalDisplay() {
    if (!state.workbook) return;

    const assignments = getAssignments();
    const assignmentByName = new Map(
      assignments.map((assignment) => [assignment.name, assignment])
    );

    let requiredGoal = 0;
    let productiveEmployeeCounts = 0;
    let supportCounts = 0;

    assignments.forEach((assignment) => {
      const total = Number(state.employeeTotals[assignment.name] || 0);
      if (isSupportAssignment(assignment)) {
        supportCounts += total;
        assignment.dailyGoal = 0;
        return;
      }
      if (!isGoalAssignment(assignment) || isAbsent(assignment)) return;
      requiredGoal += goalFor(assignment);
      productiveEmployeeCounts += total;
    });

    const overallCompleted = productiveEmployeeCounts + supportCounts;
    const overallPercent = requiredGoal > 0
      ? ((overallCompleted / requiredGoal) * 100).toFixed(1)
      : "0.0";

    document.querySelectorAll("#productionCards .summary-card").forEach((card) => {
      const name = card.querySelector("strong")?.textContent?.trim();
      const assignment = assignmentByName.get(name);
      if (!assignment || !isSupportAssignment(assignment)) return;

      card.dataset.supportTotal = "true";
      const total = Number(state.employeeTotals[assignment.name] || 0);
      const meter = card.querySelector(".meter");
      if (meter) meter.style.display = "none";

      const percentRow = card.querySelector(".percent-row");
      if (percentRow) {
        percentRow.innerHTML = `
          <span>Support total</span>
          <small>${total} added to overall completion • no 200-count goal</small>
        `;
      }

      const breakdown = card.querySelector(".summary-card-top span");
      if (breakdown) {
        breakdown.textContent = "Counts toward overall cycle-count completion only";
      }
    });

    const kpis = Array.from(document.querySelectorAll("#kpiStrip .kpi"));
    if (kpis[0]) {
      const label = kpis[0].querySelector("span");
      const value = kpis[0].querySelector("strong");
      if (label) label.textContent = "Overall Completed Counts";
      if (value) value.textContent = String(overallCompleted);
    }
    if (kpis[2]) {
      const label = kpis[2].querySelector("span");
      const value = kpis[2].querySelector("strong");
      if (label) label.textContent = "Required Team Goal";
      if (value) value.textContent = String(requiredGoal);
    }
    if (kpis[3]) {
      const label = kpis[3].querySelector("span");
      const value = kpis[3].querySelector("strong");
      if (label) label.textContent = "Overall Goal Completion";
      if (value) value.textContent = `${overallPercent}%`;
    }

    const summary = $("recordSummary");
    if (summary) {
      summary.textContent =
        `${overallCompleted} completed toward ${requiredGoal} required` +
        ` • ${supportCounts} support counts (Batches/Variance Reports)`;
    }
  }

  const previousRenderResults = renderResults;
  renderResults = function renderResultsWithSupportTotals() {
    previousRenderResults();
    applySupportGoalDisplay();
  };

  const previousRenderAssignments = renderAssignments;
  renderAssignments = function renderAssignmentsWithSupportLabels() {
    previousRenderAssignments();
    const assignments = getAssignments();
    Array.from(document.querySelectorAll("#teamBody tr")).forEach((row, index) => {
      const assignment = assignments[index];
      if (!assignment || !isSupportAssignment(assignment) || !row.children[1]) return;
      row.querySelectorAll("small").forEach((small) => small.remove());
      const note = document.createElement("small");
      note.style.display = "block";
      note.style.color = "#64748b";
      note.textContent = "Support total — no individual 200-count goal";
      row.children[1].appendChild(note);
    });
  };

  const previousCreateCurrentRecord = createCurrentRecord;
  createCurrentRecord = function createCurrentRecordWithoutSupportGoals() {
    const record = previousCreateCurrentRecord();
    const assignments = getAssignments();
    const requiredGoal = assignments.reduce((sum, assignment) => {
      if (!isGoalAssignment(assignment)) return sum;
      const attendanceEntry = record.attendance?.[assignment.id];
      return attendanceEntry?.status === "absent"
        ? sum
        : sum + goalFor(assignment);
    }, 0);
    const supportCounts = assignments.reduce(
      (sum, assignment) =>
        sum + (isSupportAssignment(assignment)
          ? Number(state.employeeTotals[assignment.name] || 0)
          : 0),
      0
    );
    return {
      ...record,
      dailyTeamGoal: requiredGoal,
      supportCounts,
      goalExcludedAssignments: assignments
        .filter(isSupportAssignment)
        .map((assignment) => assignment.name),
    };
  };

  // Normalize existing saved support assignments without deleting any counts.
  getAssignments().forEach((assignment) => {
    if (isSupportAssignment(assignment)) assignment.dailyGoal = 0;
  });
  saveBranches();
})();