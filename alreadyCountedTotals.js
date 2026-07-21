"use strict";

/*
 * Production-total presentation and employee matching for Already Cycle Counted.
 * The initials filter only changes the preview; all matched employee credits remain
 * included in production totals.
 */

function acEmployeeForInitials(initials) {
  const normalized = acNormalizeInitials(initials);
  return getAssignments().find(
    (assignment) => acGetInitials(assignment) === normalized
  ) || null;
}

acApplyCreditsToProduction = function applyAllAlreadyCountedCredits() {
  const totalsByEmployee = {};
  const assignments = getAssignments();

  assignments.forEach((assignment) => {
    const initials = acGetInitials(assignment);
    totalsByEmployee[assignment.name] = Number(
      alreadyCountedState.totalsByInitials[initials] || 0
    );
  });

  alreadyCountedState.totalsByEmployee = totalsByEmployee;

  assignments.forEach((assignment) => {
    const current = Number(state.employeeTotals[assignment.name] || 0);
    const previousExtra = Number(assignment.__alreadyCountedApplied || 0);
    const nextExtra = Number(totalsByEmployee[assignment.name] || 0);

    state.employeeTotals[assignment.name] =
      Math.max(0, current - previousExtra) + nextExtra;
    assignment.__alreadyCountedApplied = nextExtra;
  });

  if (state.workbook) renderResults();
};

acRenderPreview = function renderAlreadyCountedWithEmployeeTotals() {
  const preview = $("alreadyCountedPreview");
  if (!preview) return;

  const filter = alreadyCountedState.selectedInitials;
  const rows = alreadyCountedState.matchedRows.filter(
    (row) => filter === "all" || row.initials === filter
  );
  const filteredLocations = rows.reduce(
    (sum, row) => sum + row.locationCount,
    0
  );

  const employeeRows = getAssignments().map((assignment) => {
    const initials = acGetInitials(assignment);
    const alreadyCounted = Number(
      alreadyCountedState.totalsByInitials[initials] || 0
    );
    const productionTotal = Number(
      state.employeeTotals[assignment.name] || 0
    );
    const regularProduction = Math.max(
      0,
      productionTotal - alreadyCounted
    );

    return {
      name: assignment.name,
      initials,
      regularProduction,
      alreadyCounted,
      productionTotal,
    };
  });

  const unassignedInitials = Object.entries(
    alreadyCountedState.totalsByInitials
  )
    .filter(([initials]) => !acEmployeeForInitials(initials))
    .map(([initials, total]) => ({ initials, total }));

  preview.classList.remove("hidden");
  preview.innerHTML = `
    <div class="already-counted-kpis">
      <div><span>Matched items</span><strong>${rows.length}</strong></div>
      <div><span>Unique locations</span><strong>${filteredLocations}</strong></div>
      <div><span>Unmatched items</span><strong>${alreadyCountedState.unmatchedRows.length}</strong></div>
      <div><span>Duplicates removed</span><strong>${alreadyCountedState.duplicateCount}</strong></div>
    </div>

    <div class="table-wrap already-counted-table-wrap">
      <table class="history-table">
        <thead>
          <tr><th>Employee</th><th>Initials</th><th>Item</th><th>Locations</th><th>Bins</th></tr>
        </thead>
        <tbody>
          ${rows.slice(0, 500).map((row) => {
            const employee = acEmployeeForInitials(row.initials);
            return `
              <tr>
                <td>${escapeHtml(employee?.name || "Not assigned")}</td>
                <td>${escapeHtml(row.initials.toUpperCase())}</td>
                <td>${escapeHtml(row.itemNumber)}</td>
                <td>${row.locationCount}</td>
                <td>${escapeHtml(row.bins.join(", "))}</td>
              </tr>
            `;
          }).join("") || '<tr><td colspan="5" class="empty-row">No matching items for this filter.</td></tr>'}
        </tbody>
      </table>
    </div>

    <section class="already-counted-production-summary">
      <div class="section-heading">
        <div>
          <h3>Already Counted Added to Production</h3>
          <small>Initials are matched to the employee name, and these locations are included in the production total.</small>
        </div>
      </div>
      <div class="table-wrap">
        <table class="history-table">
          <thead>
            <tr>
              <th>Employee</th>
              <th>Initials</th>
              <th>Regular Production</th>
              <th>Already Counted Added</th>
              <th>Production Total</th>
            </tr>
          </thead>
          <tbody>
            ${employeeRows.map((employee) => `
              <tr>
                <td><strong>${escapeHtml(employee.name)}</strong></td>
                <td>${escapeHtml(employee.initials.toUpperCase() || "—")}</td>
                <td>${employee.regularProduction}</td>
                <td><strong>+${employee.alreadyCounted}</strong></td>
                <td><strong>${employee.productionTotal}</strong></td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
      ${unassignedInitials.length ? `
        <div class="message error">
          These initials have location credit but are not matched to an employee:
          ${unassignedInitials.map((entry) => `${escapeHtml(entry.initials.toUpperCase())} (${entry.total})`).join(", ")}.
          Enter those initials beside the correct employee above.
        </div>
      ` : ""}
    </section>
  `;
};
