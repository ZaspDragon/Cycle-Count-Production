"use strict";

/*
 * Read the weekday represented by the uploaded Cycle Count Detail report.
 * The daily sheet is scanned row-by-row across its full worksheet range so
 * filtered views, gaps, repeated items, capitalization, and rows far below the
 * visible header do not silently disappear.
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

function acGetWeekdaySheetName(workbook, reportDate) {
  if (!workbook || !reportDate) return null;
  const weekday = reportDate.toLocaleDateString("en-US", { weekday: "long" });
  return workbook.SheetNames.find(
    (sheetName) => sheetName.trim().toLowerCase() === weekday.toLowerCase()
  ) || null;
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

acFindAlreadyCountedRows = function findEveryAlreadyCountedRowForReportDay(workbook) {
  const reportDate = acGetReportCountDate(state.workbook);
  const weekdaySheet = acGetWeekdaySheetName(workbook, reportDate);

  if (!weekdaySheet) {
    acSetStatus(
      reportDate
        ? `No ${reportDate.toLocaleDateString("en-US", { weekday: "long" })} worksheet was found in the Already Cycle Counted file.`
        : "The report date could not be detected from Count Date values.",
      true
    );
    return [];
  }

  const rows = acReadEveryDailyRow(workbook, weekdaySheet);
  alreadyCountedState.selectedDay = weekdaySheet;
  alreadyCountedState.reportDate = reportDate;
  alreadyCountedState.dailySourceRowCount = rows.length;
  return rows;
};

const acMatchFilesBeforeFullDayStatus = acMatchFiles;
acMatchFiles = function matchEveryDailyRow() {
  acMatchFilesBeforeFullDayStatus();

  if (alreadyCountedState.applied && alreadyCountedState.selectedDay) {
    const locationTotal = alreadyCountedState.matchedRows.reduce(
      (sum, row) => sum + Number(row.locationCount || 0),
      0
    );
    acSetStatus(
      `${alreadyCountedState.selectedDay} • ${alreadyCountedState.dailySourceRowCount || 0} source rows read • ` +
      `${alreadyCountedState.matchedRows.length} items matched • ${locationTotal} cycle counts credited`,
      false
    );
  }
};
