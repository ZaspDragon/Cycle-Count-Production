"use strict";

/*
 * Make Already Cycle Counted initials case-insensitive and letter-only.
 * Examples: "AB", "ab", "A.B.", and "a b" all normalize to "ab".
 */
(() => {
  acNormalizeInitials = function normalizeInitialsLettersOnly(value) {
    return String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z]/g, "");
  };

  function normalizeInitialInputs() {
    document
      .querySelectorAll("#initialsAssignmentGrid input, #employeeInitialsInput")
      .forEach((input) => {
        const normalized = acNormalizeInitials(input.value);
        if (input.value !== normalized) input.value = normalized.toUpperCase();
      });
  }

  document.addEventListener("input", (event) => {
    const input = event.target;
    if (!(input instanceof HTMLInputElement)) return;
    if (
      input.id !== "employeeInitialsInput" &&
      !input.closest("#initialsAssignmentGrid")
    ) {
      return;
    }

    const normalized = acNormalizeInitials(input.value);
    input.value = normalized.toUpperCase();
  });

  const originalRenderInitialsAssignments = acRenderInitialsAssignments;
  acRenderInitialsAssignments = function renderLetterOnlyInitialsAssignments() {
    originalRenderInitialsAssignments();
    normalizeInitialInputs();
  };
})();
