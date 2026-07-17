/**
 * Branch Manager
 * Handles multi-branch data management and persistence
 */

const BRANCHES_STORAGE_KEY = "cycleCountProduction.branches.v1";
const CURRENT_BRANCH_KEY = "cycleCountProduction.currentBranch.v1";
const DEFAULT_DAILY_GOAL = 200;

// Legacy default assignments for backward compatibility
const LEGACY_ASSIGNMENTS = [
  { name: "Carico", aisles: ["A", "B"] },
  { name: "Ernie", aisles: ["C", "D"] },
  { name: "Cherish", aisles: ["E", "F"] },
  { name: "Layne", aisles: ["G", "H"] },
  { name: "Madison", aisles: ["I", "J"] },
  { name: "Antoine", aisles: ["K", "L"] },
];

class BranchManager {
  constructor() {
    this.branches = this.loadBranches();
    this.currentBranchId = this.loadCurrentBranchId();
    this.ensureValidState();
  }

  /**
   * Load all branches from localStorage
   */
  loadBranches() {
    try {
      const data = localStorage.getItem(BRANCHES_STORAGE_KEY);
      return data ? JSON.parse(data) : [];
    } catch (e) {
      console.error("Failed to load branches:", e);
      return [];
    }
  }

  /**
   * Load the current branch ID from localStorage
   */
  loadCurrentBranchId() {
    try {
      return localStorage.getItem(CURRENT_BRANCH_KEY) || null;
    } catch (e) {
      return null;
    }
  }

  /**
   * Save all branches to localStorage
   */
  saveBranches() {
    try {
      localStorage.setItem(BRANCHES_STORAGE_KEY, JSON.stringify(this.branches));
    } catch (e) {
      console.error("Failed to save branches:", e);
    }
  }

  /**
   * Save the current branch ID to localStorage
   */
  saveCurrentBranchId() {
    try {
      localStorage.setItem(CURRENT_BRANCH_KEY, this.currentBranchId);
    } catch (e) {
      console.error("Failed to save current branch ID:", e);
    }
  }

  /**
   * Ensure the application has a valid state:
   * - If no branches exist, create a default branch with legacy assignments
   * - If current branch ID is invalid, select the first branch
   */
  ensureValidState() {
    if (this.branches.length === 0) {
      this.createDefaultBranch();
    }

    if (!this.currentBranchId || !this.branches.find((b) => b.id === this.currentBranchId)) {
      this.currentBranchId = this.branches[0]?.id || null;
      this.saveCurrentBranchId();
    }
  }

  /**
   * Create a default branch with legacy assignments (for new users or first load)
   */
  createDefaultBranch() {
    const defaultBranch = {
      id: this.generateId(),
      name: "Main Branch",
      expectedFilename: "Inventory.xlsx",
      employees: LEGACY_ASSIGNMENTS.map((assignment) => ({
        id: this.generateId(),
        name: assignment.name,
        startAisle: assignment.aisles[0],
        endAisle: assignment.aisles[assignment.aisles.length - 1],
      })),
      dailyGoal: DEFAULT_DAILY_GOAL,
      createdAt: new Date().toISOString(),
    };
    this.branches.push(defaultBranch);
    this.currentBranchId = defaultBranch.id;
    this.saveBranches();
    this.saveCurrentBranchId();
  }

  /**
   * Get the currently selected branch
   */
  getCurrentBranch() {
    return this.branches.find((b) => b.id === this.currentBranchId) || null;
  }

  /**
   * Set the current branch by ID
   */
  setCurrentBranch(branchId) {
    if (this.branches.find((b) => b.id === branchId)) {
      this.currentBranchId = branchId;
      this.saveCurrentBranchId();
      return true;
    }
    return false;
  }

  /**
   * Create a new branch
   */
  createBranch(name, expectedFilename = "") {
    if (this.branches.some((b) => b.name.toLowerCase() === name.toLowerCase())) {
      return { success: false, error: "Branch name already exists" };
    }

    const newBranch = {
      id: this.generateId(),
      name,
      expectedFilename,
      employees: [],
      dailyGoal: DEFAULT_DAILY_GOAL,
      createdAt: new Date().toISOString(),
    };

    this.branches.push(newBranch);
    this.saveBranches();
    return { success: true, branch: newBranch };
  }

  /**
   * Rename an existing branch
   */
  renameBranch(branchId, newName) {
    const branch = this.branches.find((b) => b.id === branchId);
    if (!branch) {
      return { success: false, error: "Branch not found" };
    }

    if (this.branches.some((b) => b.id !== branchId && b.name.toLowerCase() === newName.toLowerCase())) {
      return { success: false, error: "Branch name already exists" };
    }

    branch.name = newName;
    this.saveBranches();
    return { success: true };
  }

  /**
   * Delete a branch (must leave at least one)
   */
  deleteBranch(branchId) {
    if (this.branches.length === 1) {
      return { success: false, error: "Cannot delete the only remaining branch" };
    }

    const index = this.branches.findIndex((b) => b.id === branchId);
    if (index < 0) {
      return { success: false, error: "Branch not found" };
    }

    this.branches.splice(index, 1);

    // If we deleted the current branch, switch to the first available
    if (this.currentBranchId === branchId) {
      this.currentBranchId = this.branches[0]?.id || null;
      this.saveCurrentBranchId();
    }

    this.saveBranches();
    return { success: true };
  }

  /**
   * Update branch name and expected filename
   */
  updateBranchSettings(branchId, settings) {
    const branch = this.branches.find((b) => b.id === branchId);
    if (!branch) {
      return { success: false, error: "Branch not found" };
    }

    if (settings.name && settings.name !== branch.name) {
      if (this.branches.some((b) => b.id !== branchId && b.name.toLowerCase() === settings.name.toLowerCase())) {
        return { success: false, error: "Branch name already exists" };
      }
      branch.name = settings.name;
    }

    if (settings.expectedFilename !== undefined) {
      branch.expectedFilename = settings.expectedFilename;
    }

    if (settings.dailyGoal !== undefined) {
      branch.dailyGoal = Math.max(1, settings.dailyGoal);
    }

    this.saveBranches();
    return { success: true };
  }

  /**
   * Get all employees in the current branch
   */
  getCurrentEmployees() {
    const branch = this.getCurrentBranch();
    return branch?.employees || [];
  }

  /**
   * Add an employee to the current branch
   */
  addEmployee(name, startAisle, endAisle) {
    const branch = this.getCurrentBranch();
    if (!branch) {
      return { success: false, error: "No active branch" };
    }

    // Validation
    if (!name || !name.trim()) {
      return { success: false, error: "Employee name cannot be blank" };
    }

    if (!startAisle || !endAisle) {
      return { success: false, error: "Both aisles are required" };
    }

    const normalizedName = name.trim();
    if (branch.employees.some((e) => e.name.toLowerCase() === normalizedName.toLowerCase())) {
      return { success: false, error: `${normalizedName} already exists in this branch` };
    }

    // Validate aisles
    const validation = this.validateAisleRange(startAisle, endAisle);
    if (!validation.valid) {
      return { success: false, error: validation.error };
    }

    // Check for overlapping assignments
    const overlap = this.checkAisleOverlap(startAisle, endAisle);
    if (overlap.overlaps) {
      return { success: false, error: `Aisle range overlaps with ${overlap.employee}` };
    }

    const newEmployee = {
      id: this.generateId(),
      name: normalizedName,
      startAisle: startAisle.toUpperCase(),
      endAisle: endAisle.toUpperCase(),
    };

    branch.employees.push(newEmployee);
    this.saveBranches();
    return { success: true, employee: newEmployee };
  }

  /**
   * Update an employee in the current branch
   */
  updateEmployee(employeeId, name, startAisle, endAisle) {
    const branch = this.getCurrentBranch();
    if (!branch) {
      return { success: false, error: "No active branch" };
    }

    const employee = branch.employees.find((e) => e.id === employeeId);
    if (!employee) {
      return { success: false, error: "Employee not found" };
    }

    // Validation
    if (!name || !name.trim()) {
      return { success: false, error: "Employee name cannot be blank" };
    }

    if (!startAisle || !endAisle) {
      return { success: false, error: "Both aisles are required" };
    }

    const normalizedName = name.trim();
    if (
      branch.employees.some((e) => e.id !== employeeId && e.name.toLowerCase() === normalizedName.toLowerCase())
    ) {
      return { success: false, error: `${normalizedName} already exists in this branch` };
    }

    // Validate aisles
    const validation = this.validateAisleRange(startAisle, endAisle);
    if (!validation.valid) {
      return { success: false, error: validation.error };
    }

    // Check for overlapping assignments (excluding current employee)
    const overlap = this.checkAisleOverlap(startAisle, endAisle, employeeId);
    if (overlap.overlaps) {
      return { success: false, error: `Aisle range overlaps with ${overlap.employee}` };
    }

    employee.name = normalizedName;
    employee.startAisle = startAisle.toUpperCase();
    employee.endAisle = endAisle.toUpperCase();
    this.saveBranches();
    return { success: true };
  }

  /**
   * Delete an employee from the current branch
   */
  deleteEmployee(employeeId) {
    const branch = this.getCurrentBranch();
    if (!branch) {
      return { success: false, error: "No active branch" };
    }

    const index = branch.employees.findIndex((e) => e.id === employeeId);
    if (index < 0) {
      return { success: false, error: "Employee not found" };
    }

    branch.employees.splice(index, 1);
    this.saveBranches();
    return { success: true };
  }

  /**
   * Validate an aisle range
   */
  validateAisleRange(startAisle, endAisle) {
    const start = String(startAisle || "").trim().toUpperCase();
    const end = String(endAisle || "").trim().toUpperCase();

    if (!start || !end) {
      return { valid: false, error: "Both aisles are required" };
    }

    if (!/^[A-Z]$/.test(start)) {
      return { valid: false, error: `Invalid starting aisle: ${start}` };
    }

    if (!/^[A-Z]$/.test(end)) {
      return { valid: false, error: `Invalid ending aisle: ${end}` };
    }

    if (start > end) {
      return { valid: false, error: "Starting aisle cannot be after ending aisle" };
    }

    return { valid: true };
  }

  /**
   * Check if an aisle range overlaps with any existing assignments
   * Optionally exclude a specific employee
   */
  checkAisleOverlap(startAisle, endAisle, excludeEmployeeId = null) {
    const start = startAisle.toUpperCase();
    const end = endAisle.toUpperCase();

    const branch = this.getCurrentBranch();
    if (!branch) {
      return { overlaps: false };
    }

    for (const emp of branch.employees) {
      if (excludeEmployeeId && emp.id === excludeEmployeeId) {
        continue;
      }

      const empStart = emp.startAisle;
      const empEnd = emp.endAisle;

      // Check if ranges overlap
      if (!(end < empStart || start > empEnd)) {
        return { overlaps: true, employee: emp.name };
      }
    }

    return { overlaps: false };
  }

  /**
   * Get all aisles assigned to the current branch
   */
  getBranchAisles() {
    const branch = this.getCurrentBranch();
    if (!branch || branch.employees.length === 0) {
      return [];
    }

    const aisles = new Set();
    for (const emp of branch.employees) {
      const start = emp.startAisle.charCodeAt(0);
      const end = emp.endAisle.charCodeAt(0);
      for (let code = start; code <= end; code++) {
        aisles.add(String.fromCharCode(code));
      }
    }

    return Array.from(aisles).sort();
  }

  /**
   * Get aisles formatted as ranges for display
   */
  formatAisleRange(startAisle, endAisle) {
    if (startAisle === endAisle) {
      return startAisle;
    }
    return `${startAisle}–${endAisle}`;
  }

  /**
   * Get the daily goal for the current branch
   */
  getDailyGoal() {
    const branch = this.getCurrentBranch();
    return branch?.dailyGoal || DEFAULT_DAILY_GOAL;
  }

  /**
   * Get expected filename for current branch
   */
  getExpectedFilename() {
    const branch = this.getCurrentBranch();
    return branch?.expectedFilename || "";
  }

  /**
   * Get all branches as options for UI dropdowns
   */
  getBranchOptions() {
    return this.branches.map((b) => ({
      id: b.id,
      name: b.name,
    }));
  }

  /**
   * Generate a unique ID
   */
  generateId() {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }
}

// Export for use in app.js
const branchManager = new BranchManager();
