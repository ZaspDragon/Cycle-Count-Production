"use strict";

/*
 * Daily attendance extension for Cycle Count Production.
 *
 * Attendance is stored separately from branch/team configuration so marking a
 * person absent never deletes their aisle assignment. The storage key includes
 * the local calendar date, which makes every new day start as Present.
 */
(() => {
  const ATTENDANCE_KEY_PREFIX = "cycleCountProduction.attendance.v1";
  const VALID_STATUSES = new Set(["present", "absent"]);

  function localDateKey(date = new Date()) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function attendanceStorageKey() {
    const branch = getSelectedBranch();
    return `${ATTENDANCE_KEY_PREFIX}.${branch?.id || "no-branch"}.${localDateKey()}`;
  }

  function readAttendance() {
    try {
      const stored = JSON.parse(localStorage.getItem(attendanceStorageKey()) || "{}");
      return stored && typeof stored === "object" ? stored : {};
    } catch (error) {
      console.error("Unable to read daily attendance.", error);
      return {};
    }
  }

  function writeAttendance(attendance) {
    try {
      localStorage.setItem(attendanceStorageKey(), JSON.stringify(attendance));
      return true;
    } catch (error) {
      console.error("Unable to save daily attendance.", error);
      showTeamMessage("Attendance could not be saved in this browser.", true);
      return false;
    }
  }

  function getAssignmentStatus(assignmentId) {
    const status = readAttendance()[assignmentId];
    return VALID_STATUSES.has(status) ? status : "present";
  }

  function setAssignmentStatus(assignmentId, status) {
    if (!VALID_STATUSES.has(status)) {
      return;
    }

    const attendance = readAttendance();
    attendance[assignmentId] = status;

    if (writeAttendance(attendance)) {
      const assignment = getAssignments().find((item) => item.id === assignmentId);
      showTeamMessage(`${assignment?.name || "Employee"} marked ${status}.`);
    }
  }

  function getDailyAttendanceSnapshot() {
    return Object.fromEntries(
      getAssignments().map((assignment) => [
        assignment.id,
        {
          name: assignment.name,
          status: getAssignmentStatus(assignment.id),
        },
      ])
    );
  }

  function getPresentAssignments() {
    return getAssignments().filter(
      (assignment) => getAssignmentStatus(assignment.id) !== "absent"
    );
  }

  function injectAttendanceStyles() {
    if (document.getElementById("attendanceStyles")) {
      return;
    }

    const style = document.createElement("style");
    style.id = "attendanceStyles";
    style.textContent = `
      .attendance-select {
        min-width: 112px;
        padding: 0.5rem 0.65rem;
        border: 1px solid #cbd5e1;
        border-radius: 0.55rem;
        background: #fff;
        font: inherit;
        font-weight: 700;
      }

      .attendance-select[data-status="absent"] {
        border-color: #fca5a5;
        background: #fff1f2;
        color: #b91c1c;
      }

      .summary-card.is-absent {
        border-style: dashed;
        opacity: 0.82;
      }

      .summary-card.is-absent .meter span {
        width: 0 !important;
      }

      .absent-label {
        display: inline-flex;
        align-items: center;
        border-radius: 999px;
        padding: 0.25rem 0.6rem;
        background: #fff1f2;
        color: #b91c1c;
        font-weight: 800;
      }

      .attendance-summary {
        display: block;
        margin-top: 0.2rem;
        color: #64748b;
        font-size: 0.78rem;
      }
    `;
    document.head.appendChild(style);
  }

  const originalRenderAssignments = renderAssignments;
  renderAssignments = function renderAssignmentsWithAttendance() {
    originalRenderAssignments();

    const assignments = getAssignments();
    const rows = Array.from(document.querySelectorAll("#teamBody tr"));

    rows.forEach((row, index) => {
      const assignment = assignments[index];
      if (!assignment) {
        return;
      }

      const status = getAssignmentStatus(assignment.id);
      const cell = document.createElement("td");
      cell.innerHTML = `
        <select
          class="attendance-select"
          data-attendance-id="${escapeHtml(assignment.id)}"
          data-status="${status}"
          aria-label="Daily status for ${escapeHtml(assignment.name)}"
        >
          <option value="present" ${status === "present" ? "selected" : ""}>Present</option>
          <option value="absent" ${status === "absent" ? "selected" : ""}>Absent</option>
        </select>
      `;

      const editCell = row.children[2];
      row.insertBefore(cell, editCell || null);
    });
  };

  const originalUpdateUIFromSelectedBranch = updateUIFromSelectedBranch;
  updateUIFromSelectedBranch = function updateUIWithAttendance() {
    originalUpdateUIFromSelectedBranch();

    const goalUnit = $("dailyGoalUnit");
    if (goalUnit) {
      const presentCount = getPresentAssignments().length;
      const totalCount = getAssignments().length;
      goalUnit.textContent = `counts = 100% (${presentCount} present of ${totalCount})`;
    }
  };

  const originalRenderResults = renderResults;
  renderResults = function renderResultsWithAttendance() {
    originalRenderResults();

    if (!state.workbook) {
      return;
    }

    const presentAssignments = getPresentAssignments();
    const presentNames = new Set(presentAssignments.map((assignment) => assignment.name));
    const teamGoal = DAILY_GOAL * presentAssignments.length;
    const credited = Object.values(state.aisleTotals).reduce(
      (total, count) => total + Number(count || 0),
      0
    );
    const teamPercentage = teamGoal > 0 ? ((credited / teamGoal) * 100).toFixed(1) : "0.0";

    const kpis = Array.from(document.querySelectorAll("#kpiStrip .kpi"));
    if (kpis[2]) {
      const value = kpis[2].querySelector("strong");
      if (value) value.textContent = String(teamGoal);
    }
    if (kpis[3]) {
      const value = kpis[3].querySelector("strong");
      if (value) value.textContent = `${teamPercentage}%`;
    }

    document.querySelectorAll("#productionCards .summary-card").forEach((card) => {
      const name = card.querySelector("strong")?.textContent?.trim();
      if (!name || presentNames.has(name)) {
        return;
      }

      card.classList.add("is-absent");
      const percentRow = card.querySelector(".percent-row");
      if (percentRow) {
        percentRow.innerHTML = `
          <span class="absent-label">Absent</span>
          <small>Excluded from today’s goal</small>
        `;
      }
    });
  };

  const originalCreateCurrentRecord = createCurrentRecord;
  createCurrentRecord = function createCurrentRecordWithAttendance() {
    const record = originalCreateCurrentRecord();
    const attendance = getDailyAttendanceSnapshot();
    const absentNames = Object.values(attendance)
      .filter((entry) => entry.status === "absent")
      .map((entry) => entry.name);

    return {
      ...record,
      attendanceDate: localDateKey(),
      attendance,
      absentEmployees: absentNames,
      activeEmployeeCount: getPresentAssignments().length,
      dailyTeamGoal: DAILY_GOAL * getPresentAssignments().length,
    };
  };

  const originalRenderHistory = renderHistory;
  renderHistory = function renderHistoryWithAttendance() {
    originalRenderHistory();

    const history = loadHistory();
    const rows = Array.from(document.querySelectorAll("#historyBody tr"));

    rows.forEach((row, index) => {
      const record = history[index];
      const firstCell = row.children[0];
      if (!record || !firstCell || !Array.isArray(record.absentEmployees) || record.absentEmployees.length === 0) {
        return;
      }

      const note = document.createElement("small");
      note.className = "attendance-summary";
      note.textContent = `Absent: ${record.absentEmployees.join(", ")}`;
      firstCell.appendChild(note);
    });
  };

  document.addEventListener("change", (event) => {
    const select = event.target.closest("[data-attendance-id]");
    if (!select) {
      return;
    }

    setAssignmentStatus(select.dataset.attendanceId, select.value);
    select.dataset.status = select.value;
    updateUIFromSelectedBranch();
    renderResults();
  });

  injectAttendanceStyles();
})();
