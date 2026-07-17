"use strict";

/**
 * Cycle Count Production - Branch Manager
 *
 * Handles:
 * - Branch storage
 * - Selected branch storage
 * - Employee / aisle assignments
 * - Old data migration
 * - Validation
 *
 * This file does not control the page UI.
 * app.js should handle rendering and event listeners.
 */

(function initializeBranchManagerModule(global) {
  const DEFAULT_DAILY_GOAL = 200;

  const STORAGE_KEYS = {
    branches: "cycleCountProduction.branches.v1",
    currentBranch: "cycleCountProduction.currentBranch.v1",
    legacySelectedBranch: "cycleCountProduction.selectedBranch.v1",
  };

  const DEFAULT_ASSIGNMENTS = [
    {
      name: "Carico",
      startAisle: "A",
      endAisle: "B",
    },
    {
      name: "Ernie",
      startAisle: "C",
      endAisle: "D",
    },
    {
      name: "Cherish",
      startAisle: "E",
      endAisle: "F",
    },
    {
      name: "Layne",
      startAisle: "G",
      endAisle: "H",
    },
    {
      name: "Madison",
      startAisle: "I",
      endAisle: "J",
    },
    {
      name: "Antoine",
      startAisle: "K",
      endAisle: "L",
    },
  ];

  function createId(prefix = "id") {
    if (
      typeof crypto !== "undefined" &&
      typeof crypto.randomUUID === "function"
    ) {
      return `${prefix}-${crypto.randomUUID()}`;
    }

    return `${prefix}-${Date.now()}-${Math.random()
      .toString(36)
      .slice(2, 10)}`;
  }

  function safeReadStorage(key) {
    try {
      return localStorage.getItem(key);
    } catch (error) {
      console.error(`Could not read localStorage key "${key}".`, error);
      return null;
    }
  }

  function safeWriteStorage(key, value) {
    try {
      localStorage.setItem(key, value);
      return true;
    } catch (error) {
      console.error(`Could not write localStorage key "${key}".`, error);
      return false;
    }
  }

  function normalizeAisle(value) {
    return String(value ?? "")
      .trim()
      .toUpperCase();
  }

  function normalizeAssignment(rawAssignment) {
    if (!rawAssignment || typeof rawAssignment !== "object") {
      return null;
    }

    const name = String(rawAssignment.name ?? "").trim();

    const legacyAisles = Array.isArray(rawAssignment.aisles)
      ? rawAssignment.aisles
      : [];

    const startAisle = normalizeAisle(
      rawAssignment.startAisle ?? legacyAisles[0] ?? ""
    );

    const endAisle = normalizeAisle(
      rawAssignment.endAisle ??
        legacyAisles[legacyAisles.length - 1] ??
        startAisle
    );

    if (!name || !startAisle || !endAisle) {
      return null;
    }

    return {
      id: String(rawAssignment.id || createId("employee")),
      name,
      startAisle,
      endAisle,
    };
  }

  function normalizeBranch(rawBranch) {
    if (!rawBranch || typeof rawBranch !== "object") {
      return null;
    }

    const assignmentsSource = Array.isArray(rawBranch.assignments)
      ? rawBranch.assignments
      : Array.isArray(rawBranch.employees)
        ? rawBranch.employees
        : [];

    const assignments = assignmentsSource
      .map(normalizeAssignment)
      .filter(Boolean);

    return {
      id: String(rawBranch.id || createId("branch")),
      name: String(
        rawBranch.name ||
          rawBranch.branchName ||
          "Unnamed Branch"
      ).trim(),
      expectedInventoryFilename: String(
        rawBranch.expectedInventoryFilename ??
          rawBranch.expectedFilename ??
          ""
      ).trim(),
      assignments,
      dailyGoal:
        Number(rawBranch.dailyGoal) > 0
          ? Number(rawBranch.dailyGoal)
          : DEFAULT_DAILY_GOAL,
      createdAt:
        rawBranch.createdAt ||
        new Date().toISOString(),
    };
  }

  function createDefaultBranch() {
    return {
      id: createId("branch"),
      name: "Main Branch",
      expectedInventoryFilename: "Inventory.xlsx",
      assignments: DEFAULT_ASSIGNMENTS.map((assignment) => ({
        id: createId("employee"),
        ...assignment,
      })),
      dailyGoal: DEFAULT_DAILY_GOAL,
      createdAt: new Date().toISOString(),
    };
  }

  class BranchManager {
    constructor() {
      this.branches = [];
      this.currentBranchId = null;
      this.storageError = false;

      this.load();
    }

    load() {
      const rawBranches = safeReadStorage(STORAGE_KEYS.branches);

      if (!rawBranches) {
        this.branches = [createDefaultBranch()];
        this.currentBranchId = this.branches[0].id;

        this.saveBranches();
        this.saveCurrentBranchId();

        return;
      }

      try {
        const parsedBranches = JSON.parse(rawBranches);

        if (!Array.isArray(parsedBranches)) {
          throw new Error("Stored branch data is not an array.");
        }

        const normalizedBranches = parsedBranches
          .map(normalizeBranch)
          .filter(Boolean);

        if (normalizedBranches.length === 0) {
          this.branches = [createDefaultBranch()];
        } else {
          this.branches = normalizedBranches;
        }

        this.loadCurrentBranchId();
        this.ensureValidCurrentBranch();

        /*
         * Save normalized data only after parsing succeeded.
         * This migrates old employees/expectedFilename structures.
         */
        this.saveBranches();
        this.saveCurrentBranchId();
      } catch (error) {
        console.error(
          "Stored branch data could not be parsed. Existing data was not deleted.",
          error
        );

        this.storageError = true;
        this.branches = [createDefaultBranch()];
        this.currentBranchId = this.branches[0].id;

        /*
         * Do not overwrite malformed stored data here.
         */
      }
    }

    loadCurrentBranchId() {
      const currentId =
        safeReadStorage(STORAGE_KEYS.currentBranch) ||
        safeReadStorage(STORAGE_KEYS.legacySelectedBranch);

      this.currentBranchId = currentId || null;
    }

    ensureValidCurrentBranch() {
      const currentExists = this.branches.some(
        (branch) => branch.id === this.currentBranchId
      );

      if (!currentExists) {
        this.currentBranchId = this.branches[0]?.id || null;
      }
    }

    saveBranches() {
      if (this.storageError) {
        return false;
      }

      return safeWriteStorage(
        STORAGE_KEYS.branches,
        JSON.stringify(this.branches)
      );
    }

    saveCurrentBranchId() {
      if (!this.currentBranchId || this.storageError) {
        return false;
      }

      const currentSaved = safeWriteStorage(
        STORAGE_KEYS.currentBranch,
        this.currentBranchId
      );

      /*
       * Keep the older selected-branch key synchronized.
       */
      safeWriteStorage(
        STORAGE_KEYS.legacySelectedBranch,
        this.currentBranchId
      );

      return currentSaved;
    }

    getBranches() {
      return this.branches.map((branch) => ({
        ...branch,
        assignments: branch.assignments.map((assignment) => ({
          ...assignment,
        })),
      }));
    }

    getCurrentBranch() {
      return (
        this.branches.find(
          (branch) => branch.id === this.currentBranchId
        ) || null
      );
    }

    getCurrentBranchId() {
      return this.currentBranchId;
    }

    setCurrentBranch(branchId) {
      const branchExists = this.branches.some(
        (branch) => branch.id === branchId
      );

      if (!branchExists) {
        return {
          success: false,
          error: "Branch could not be found.",
        };
      }

      this.currentBranchId = branchId;
      this.saveCurrentBranchId();

      return {
        success: true,
        branch: this.getCurrentBranch(),
      };
    }

    createBranch(name, expectedInventoryFilename = "") {
      const normalizedName = String(name ?? "").trim();

      if (!normalizedName) {
        return {
          success: false,
          error: "Branch name cannot be empty.",
        };
      }

      const duplicate = this.branches.some(
        (branch) =>
          branch.name.toLowerCase() ===
          normalizedName.toLowerCase()
      );

      if (duplicate) {
        return {
          success: false,
          error: "A branch with this name already exists.",
        };
      }

      const branch = {
        id: createId("branch"),
        name: normalizedName,
        expectedInventoryFilename: String(
          expectedInventoryFilename ?? ""
        ).trim(),
        assignments: [],
        dailyGoal: DEFAULT_DAILY_GOAL,
        createdAt: new Date().toISOString(),
      };

      this.branches.push(branch);
      this.currentBranchId = branch.id;

      this.saveBranches();
      this.saveCurrentBranchId();

      return {
        success: true,
        branch,
      };
    }

    updateBranch(
      branchId,
      {
        name,
        expectedInventoryFilename,
        dailyGoal,
      } = {}
    ) {
      const branch = this.branches.find(
        (item) => item.id === branchId
      );

      if (!branch) {
        return {
          success: false,
          error: "Branch could not be found.",
        };
      }

      if (name !== undefined) {
        const normalizedName = String(name).trim();

        if (!normalizedName) {
          return {
            success: false,
            error: "Branch name cannot be empty.",
          };
        }

        const duplicate = this.branches.some(
          (item) =>
            item.id !== branchId &&
            item.name.toLowerCase() ===
              normalizedName.toLowerCase()
        );

        if (duplicate) {
          return {
            success: false,
            error: "A branch with this name already exists.",
          };
        }

        branch.name = normalizedName;
      }

      if (expectedInventoryFilename !== undefined) {
        branch.expectedInventoryFilename = String(
          expectedInventoryFilename
        ).trim();
      }

      if (dailyGoal !== undefined) {
        const parsedGoal = Number(dailyGoal);

        if (!Number.isFinite(parsedGoal) || parsedGoal < 1) {
          return {
            success: false,
            error: "Daily goal must be at least 1.",
          };
        }

        branch.dailyGoal = parsedGoal;
      }

      this.saveBranches();

      return {
        success: true,
        branch,
      };
    }

    renameBranch(branchId, newName) {
      return this.updateBranch(branchId, {
        name: newName,
      });
    }

    deleteBranch(branchId) {
      if (this.branches.length <= 1) {
        return {
          success: false,
          error: "You must keep at least one branch.",
        };
      }

      const branchIndex = this.branches.findIndex(
        (branch) => branch.id === branchId
      );

      if (branchIndex < 0) {
        return {
          success: false,
          error: "Branch could not be found.",
        };
      }

      const deletedBranch = this.branches[branchIndex];

      this.branches.splice(branchIndex, 1);

      if (this.currentBranchId === branchId) {
        this.currentBranchId = this.branches[0]?.id || null;
      }

      this.saveBranches();
      this.saveCurrentBranchId();

      return {
        success: true,
        deletedBranch,
        currentBranch: this.getCurrentBranch(),
      };
    }

    getCurrentAssignments() {
      return this.getCurrentBranch()?.assignments || [];
    }

    getCurrentEmployees() {
      return this.getCurrentAssignments();
    }

    validateAisleRange(startAisle, endAisle) {
      const start = normalizeAisle(startAisle);
      const end = normalizeAisle(endAisle);

      if (!start || !end) {
        return {
          valid: false,
          error: "Starting and ending aisles are required.",
        };
      }

      if (!/^[A-Z]$/.test(start)) {
        return {
          valid: false,
          error: "Starting aisle must be a single letter from A through Z.",
        };
      }

      if (!/^[A-Z]$/.test(end)) {
        return {
          valid: false,
          error: "Ending aisle must be a single letter from A through Z.",
        };
      }

      if (start > end) {
        return {
          valid: false,
          error:
            "Starting aisle cannot come after the ending aisle.",
        };
      }

      return {
        valid: true,
        startAisle: start,
        endAisle: end,
      };
    }

    expandAisleRange(startAisle, endAisle) {
      const validation = this.validateAisleRange(
        startAisle,
        endAisle
      );

      if (!validation.valid) {
        return [];
      }

      const aisles = [];
      const startCode = validation.startAisle.charCodeAt(0);
      const endCode = validation.endAisle.charCodeAt(0);

      for (let code = startCode; code <= endCode; code += 1) {
        aisles.push(String.fromCharCode(code));
      }

      return aisles;
    }

    checkAisleOverlap(
      startAisle,
      endAisle,
      excludedAssignmentId = null
    ) {
      const proposedAisles = this.expandAisleRange(
        startAisle,
        endAisle
      );

      const assignments = this.getCurrentAssignments();

      for (const assignment of assignments) {
        if (assignment.id === excludedAssignmentId) {
          continue;
        }

        const existingAisles = this.expandAisleRange(
          assignment.startAisle,
          assignment.endAisle
        );

        const overlaps = proposedAisles.some((aisle) =>
          existingAisles.includes(aisle)
        );

        if (overlaps) {
          return {
            overlaps: true,
            employee: assignment.name,
            assignment,
          };
        }
      }

      return {
        overlaps: false,
      };
    }

    validateAssignment(
      name,
      startAisle,
      endAisle,
      excludedAssignmentId = null
    ) {
      const branch = this.getCurrentBranch();

      if (!branch) {
        return {
          valid: false,
          error: "No branch is currently selected.",
        };
      }

      const normalizedName = String(name ?? "").trim();

      if (!normalizedName) {
        return {
          valid: false,
          error: "Employee name is required.",
        };
      }

      const aisleValidation = this.validateAisleRange(
        startAisle,
        endAisle
      );

      if (!aisleValidation.valid) {
        return aisleValidation;
      }

      const duplicateName = branch.assignments.some(
        (assignment) =>
          assignment.id !== excludedAssignmentId &&
          assignment.name.toLowerCase() ===
            normalizedName.toLowerCase()
      );

      if (duplicateName) {
        return {
          valid: false,
          error: `${normalizedName} already exists in this branch.`,
        };
      }

      const overlap = this.checkAisleOverlap(
        aisleValidation.startAisle,
        aisleValidation.endAisle,
        excludedAssignmentId
      );

      if (overlap.overlaps) {
        return {
          valid: false,
          error: `The aisle range overlaps with ${overlap.employee}.`,
        };
      }

      return {
        valid: true,
        name: normalizedName,
        startAisle: aisleValidation.startAisle,
        endAisle: aisleValidation.endAisle,
      };
    }

    addAssignment(name, startAisle, endAisle) {
      const branch = this.getCurrentBranch();

      if (!branch) {
        return {
          success: false,
          error: "No branch is currently selected.",
        };
      }

      const validation = this.validateAssignment(
        name,
        startAisle,
        endAisle
      );

      if (!validation.valid) {
        return {
          success: false,
          error: validation.error,
        };
      }

      const assignment = {
        id: createId("employee"),
        name: validation.name,
        startAisle: validation.startAisle,
        endAisle: validation.endAisle,
      };

      branch.assignments.push(assignment);
      this.saveBranches();

      return {
        success: true,
        assignment,
      };
    }

    addEmployee(name, startAisle, endAisle) {
      return this.addAssignment(
        name,
        startAisle,
        endAisle
      );
    }

    updateAssignment(
      assignmentId,
      name,
      startAisle,
      endAisle
    ) {
      const branch = this.getCurrentBranch();

      if (!branch) {
        return {
          success: false,
          error: "No branch is currently selected.",
        };
      }

      const assignment = branch.assignments.find(
        (item) => item.id === assignmentId
      );

      if (!assignment) {
        return {
          success: false,
          error: "Employee could not be found.",
        };
      }

      const validation = this.validateAssignment(
        name,
        startAisle,
        endAisle,
        assignmentId
      );

      if (!validation.valid) {
        return {
          success: false,
          error: validation.error,
        };
      }

      assignment.name = validation.name;
      assignment.startAisle = validation.startAisle;
      assignment.endAisle = validation.endAisle;

      this.saveBranches();

      return {
        success: true,
        assignment,
      };
    }

    updateEmployee(
      employeeId,
      name,
      startAisle,
      endAisle
    ) {
      return this.updateAssignment(
        employeeId,
        name,
        startAisle,
        endAisle
      );
    }

    deleteAssignment(assignmentId) {
      const branch = this.getCurrentBranch();

      if (!branch) {
        return {
          success: false,
          error: "No branch is currently selected.",
        };
      }

      const assignmentIndex = branch.assignments.findIndex(
        (assignment) => assignment.id === assignmentId
      );

      if (assignmentIndex < 0) {
        return {
          success: false,
          error: "Employee could not be found.",
        };
      }

      const deletedAssignment =
        branch.assignments[assignmentIndex];

      branch.assignments.splice(assignmentIndex, 1);

      this.saveBranches();

      return {
        success: true,
        assignment: deletedAssignment,
      };
    }

    deleteEmployee(employeeId) {
      return this.deleteAssignment(employeeId);
    }

    getAssignedAisles() {
      const assignedAisles = new Set();

      this.getCurrentAssignments().forEach((assignment) => {
        this.expandAisleRange(
          assignment.startAisle,
          assignment.endAisle
        ).forEach((aisle) => assignedAisles.add(aisle));
      });

      return Array.from(assignedAisles).sort();
    }

    formatAisleRange(startAisle, endAisle) {
      const start = normalizeAisle(startAisle);
      const end = normalizeAisle(endAisle);

      return start === end ? start : `${start}–${end}`;
    }

    resetStorageError() {
      this.storageError = false;
    }
  }

  /*
   * Make the class available to app.js:
   *
   * const branchManager = new BranchManager();
   */
  global.BranchManager = BranchManager;
})(window);
