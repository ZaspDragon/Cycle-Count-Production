"use strict";

/* Reliable browser persistence for branch rosters and team changes. */
(() => {
  const PRIMARY_KEY = "cycleCountProduction.branches.v1";
  const BACKUP_KEY = "cycleCountProduction.branches.backup.v1";
  const SELECTED_KEY = "cycleCountProduction.currentBranch.v1";

  // Maintain compatibility with the branch modal title expected by app.js.
  const branchTitle = document.getElementById("branchMemberModalTitle");
  if (branchTitle) branchTitle.id = "branchModalTitle";

  function validBranches(value) {
    return Array.isArray(value) && value.length > 0;
  }

  function parseSaved(key) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return null;
      const value = JSON.parse(raw);
      return validBranches(value) ? value : null;
    } catch (error) {
      console.error(`Unable to read saved roster from ${key}.`, error);
      return null;
    }
  }

  function persistRoster() {
    if (!validBranches(state.branches)) return false;

    try {
      const serialized = JSON.stringify(state.branches);
      localStorage.setItem(PRIMARY_KEY, serialized);
      localStorage.setItem(BACKUP_KEY, serialized);

      if (state.selectedBranchId) {
        localStorage.setItem(SELECTED_KEY, state.selectedBranchId);
      }

      return true;
    } catch (error) {
      console.error("Unable to persist team roster.", error);
      showAppError(
        "Your team change could not be saved in this browser. Check that browser storage is enabled and you are not using a private window."
      );
      return false;
    }
  }

  const primaryRoster = parseSaved(PRIMARY_KEY);
  const backupRoster = parseSaved(BACKUP_KEY);

  if (!primaryRoster && backupRoster) {
    state.branches = backupRoster.map(normalizeBranch).filter(Boolean);

    const savedSelected = localStorage.getItem(SELECTED_KEY);
    state.selectedBranchId = state.branches.some(
      (branch) => branch.id === savedSelected
    )
      ? savedSelected
      : state.branches[0]?.id || null;

    persistRoster();
    renderBranchDropdown();
    renderAssignments();
    renderAssignmentGrid();
  } else if (primaryRoster) {
    persistRoster();
  }

  const originalSaveBranches = saveBranches;
  saveBranches = function saveBranchesWithBackup() {
    const primarySaved = originalSaveBranches();
    const backupSaved = persistRoster();
    return Boolean(primarySaved && backupSaved);
  };

  function wrapRosterMutation(functionName) {
    const original = window[functionName];
    if (typeof original !== "function") return;

    window[functionName] = function persistAfterRosterMutation(...args) {
      const result = original.apply(this, args);
      if (result?.success) persistRoster();
      return result;
    };
  }

  [
    "addAssignment",
    "updateAssignment",
    "deleteAssignment",
    "addBranch",
    "updateBranch",
    "deleteSelectedBranch",
  ].forEach(wrapRosterMutation);

  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") persistRoster();
  });
  window.addEventListener("pagehide", persistRoster);
  window.addEventListener("beforeunload", persistRoster);
  window.setTimeout(persistRoster, 0);
})();