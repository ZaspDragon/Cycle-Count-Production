"use strict";

/*
 * Read the report weekday plus the immediately previous workday from the
 * Already Cycle Counted workbook. This allows counts entered a day early to
 * receive credit on the report day while preventing the same item/initials
 * pair from being counted twice across both tabs.
 */

function acGetReportCountDate(workbook) {
  if (!workbook) return null;
  const dates = [];

  workbook.SheetNames.forEach((sheetName) => {
    workbookMatrix(workbook, sheetName).forEach((row) => {
      row.forEach((cell) => {
        if (cell instanceof Date && !Number.isNaN(cell.getTime())) {
          dates.push(cell);
          return;
        }
        if (typeof cell !== "string") return;
        const text = cell.trim();
        if (!/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(text)) return;
        const parsed = new Date(text);
        if (!Number.isNaN(parsed.getTime())) dates.push(parsed);
      });
    });
  });

  if (!dates.length) return null;
  const counts = new Map();
  dates.forEach((date) => {
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
    counts.set(key, (counts.get(key) || 0) + 1);
  });

  const selectedKey = [...counts.entries()]
    .sort((a, b) => b[1] - a[1] || b[0].localeCompare(a[0]))[0]?.[0];
  if (!selectedKey) return null;
  const [year, month, day] = selectedKey.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function acGetWeekdaySheetName(workbook, date) {
  if (!workbook || !date) return null;
  const weekday = date.toLocaleDateString("en-US", { weekday: "long" });
  return workbook.SheetNames.find(
    (sheetName) => sheetName.trim().toLowerCase() === weekday.toLowerCase()
  ) || null;
}

function acGetPreviousWorkday(date) {
  if (!date) return null;
  const previous = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  previous.setDate(previous.getDate() - 1);
  if (previous.getDay() === 0) previous.setDate(previous.getDate() - 2);
  if (previous.getDay() === 6) previous.setDate(previous.getDate() - 1);
  return previous;
}

function acReadEveryDailyRow(workbook, sheetName) {
  const matrix = workbookMatrix(workbook, sheetName);
  const rows = [];
  let itemColumn = -1;
  let initialsColumn = -1;

  for (let rowIndex = 0; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] || [];

    if (itemColumn < 0 || initialsColumn < 0) {
      const detectedItem = detectColumn(row, ["item number", "item num", "item"]);
      const detectedInitials = detectColumn(row, ["initials", "initial"]);
      if (detectedItem >= 0 && detectedInitials >= 0) {
        itemColumn = detectedItem;
        initialsColumn = detectedInitials;
        continue;
      }
    }

    if (itemColumn < 0 || initialsColumn < 0) continue;

    const itemNumber = acNormalizeItem(row[itemColumn]);
    const initials = acNormalizeInitials(row[initialsColumn]);
    if (!itemNumber || !/^\d+$/.test(itemNumber) || !initials) continue;

    rows.push({
      itemNumber,
      initials,
      sheetName,
      rowNumber: rowIndex + 1,
    });
  }

  return rows;
}

acFindAlreadyCountedRows = function findAlreadyCountedRowsForTwoDayWindow(workbook) {
  const reportDate = acGetReportCountDate(state.workbook);
  const reportSheet = acGetWeekdaySheetName(workbook, reportDate);
  const previousDate = acGetPreviousWorkday(reportDate);
  const previousSheet = acGetWeekdaySheetName(workbook, previousDate);

  const sheets = [...new Set([previousSheet, reportSheet].filter(Boolean))];

  if (!reportSheet) {
    acSetStatus(
      reportDate
        ? `No ${reportDate.toLocaleDateString("en-US", { weekday: "long" })} worksheet was found in the Already Cycle Counted file.`
        : "The report date could not be detected from Count Date values.",
      true
    );
    return [];
  }

  const sourceRows = sheets.flatMap((sheetName) =>
    acReadEveryDailyRow(workbook, sheetName)
  );

  const seen = new Set();
  const rows = sourceRows.filter((row) => {
    const key = `${row.itemNumber}|${row.initials}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });

  alreadyCountedState.selectedDay = reportSheet;
  alreadyCountedState.selectedSheets = sheets;
  alreadyCountedState.reportDate = reportDate;
  alreadyCountedState.dailySourceRowCount = sourceRows.length;
  alreadyCountedState.dailyUniqueRowCount = rows.length;
  return rows;
};

const acMatchFilesBeforeTwoDayStatus = acMatchFiles;
acMatchFiles = function matchTwoDayEntryWindow() {
  acMatchFilesBeforeTwoDayStatus();

  if (alreadyCountedState.applied && alreadyCountedState.selectedSheets?.length) {
    const locationTotal = alreadyCountedState.matchedRows.reduce(
      (sum, row) => sum + Number(row.locationCount || 0),
      0
    );
    acSetStatus(
      `${alreadyCountedState.selectedSheets.join(" + ")} entry window • ` +
      `${alreadyCountedState.dailySourceRowCount || 0} source rows read • ` +
      `${alreadyCountedState.dailyUniqueRowCount || 0} unique item/initial entries • ` +
      `${alreadyCountedState.matchedRows.length} items matched • ${locationTotal} cycle counts credited`,
      false
    );
  }
};
