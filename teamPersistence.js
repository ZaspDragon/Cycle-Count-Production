"use strict";

/*
 * Reliable browser persistence for branch rosters.
 * Keeps a second backup of team/branch data and saves after every add, edit,
 * delete, branch change, page hide, and page close.
 */
(() => {
  const PRIMARY_KEY = "cycleCountProduction.branches.v1";
  const BACKUP_KEY = "cycleCountProduction.branches.backup.v1";
  const SELECTED_KEY = "cycleCountProduction.currentBranch.v1";

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

  // Recover from the backup only when the primary roster is unavailable.
  const primaryRoster = parseSaved(PRIMARY_KEY);
  const backupRoster = parseSaved(BACKUP_KEY);
  if (!primaryRoster && backupRoster) {
    state.branches = backupRoster
      .map(normalizeBranch)
      .filter(Boolean);

    const savedSelected = localStorage.getItem(SELECTED_KEY);
    state.selectedBranchId = state.branches.some(
      (branch) => branch.id === savedSelected
    )
      ? savedSelected
      : state.branches[0]?.id || null;

    persistRoster();
  } else if (primaryRoster) {
    // Seed or refresh the backup with the currently loaded roster.
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
      if (result?.success) {
        persistRoster();
      }
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

  // Save once after all scripts have initialized and branch defaults are added.
  window.setTimeout(persistRoster, 0);
})();