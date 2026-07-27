"use strict";

/*
 * Stable initials ownership for Already Cycle Counted production.
 * Keeps matching case-insensitive and prevents stale saved initials from
 * blocking credit for the established team mapping (including AH -> Antoine).
 */
(() => {
  if (window.__antoineInitialsCreditFixLoaded) return;
  window.__antoineInitialsCreditFixLoaded = true;

  const CANONICAL_INITIALS_BY_NAME = Object.freeze({
    carico: "ch",
    ernie: "eh",
    cherish: "cc",
    layne: "lm",
    madison: "mj",
    antoine: "ah",
    greg: "gr",
    denise: "dw",
  });

  function normalizedName(value) {
    return String(value ?? "").trim().toLowerCase();
  }

  function normalizedInitials(value) {
    return String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z]/g, "");
  }

  const previousGetInitials = acGetInitials;
  acGetInitials = function getStableEmployeeInitials(assignment) {
    const canonical = CANONICAL_INITIALS_BY_NAME[normalizedName(assignment?.name)];
    if (canonical) return canonical;
    return normalizedInitials(previousGetInitials(assignment));
  };

  function rebuildAlreadyCountedTotals() {
    if (!alreadyCountedState?.matchedRows) return;

    const totals = {};
    alreadyCountedState.matchedRows.forEach((row) => {
      const initials = normalizedInitials(row.initials);
      if (!initials) return;
      row.initials = initials;
      totals[initials] = (totals[initials] || 0) + Number(row.locationCount || 0);
    });
    alreadyCountedState.totalsByInitials = totals;
  }

  const previousApplyCredits = acApplyCreditsToProduction;
  acApplyCreditsToProduction = function applyStableInitialsCredits() {
    rebuildAlreadyCountedTotals();
    previousApplyCredits();
  };

  function migrateCanonicalInitials() {
    let changed = false;
    getAssignments().forEach((assignment) => {
      const canonical = CANONICAL_INITIALS_BY_NAME[normalizedName(assignment?.name)];
      if (!canonical || normalizedInitials(assignment.initials) === canonical) return;
      assignment.initials = canonical;
      changed = true;
    });

    if (changed) {
      saveBranches();
      if (typeof acSaveSettings === "function") acSaveSettings();
    }

    if (typeof acRenderInitialsAssignments === "function") {
      acRenderInitialsAssignments();
    }

    if (alreadyCountedState?.applied) {
      acApplyCreditsToProduction();
      if (typeof acRenderPreview === "function") acRenderPreview();
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", migrateCanonicalInitials);
  } else {
    migrateCanonicalInitials();
  }

  $("branchSelect")?.addEventListener("change", () => {
    window.setTimeout(migrateCanonicalInitials, 0);
  });
})();
