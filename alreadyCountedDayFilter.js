"use strict";

/*
 * Restrict Already Cycle Counted matching to the weekday represented by the
 * uploaded Cycle Count Detail report. Example: a Tuesday report reads only
 * the Tuesday worksheet and never mixes Monday rows into Tuesday production.
 */

function acGetReportCountDate(workbook) {
  if (!workbook) return null;

  const dates = [];

  workbook.SheetNames.forEach((sheetName) => {
    const matrix = workbookMatrix(workbook, sheetName);

    matrix.forEach((row) => {
      row.forEach((cell) => {
        if (cell instanceof Date && !Number.isNaN(cell.getTime())) {
          dates.push(cell);
          return;
        }

        if (typeof cell === "string") {
          const text = cell.trim();
          if (!/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(text)) return;
          const parsed = new Date(text);
          if (!Number.isNaN(parsed.getTime())) dates.push(parsed);
        }
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

const acFindAlreadyCountedRowsAllSheets = acFindAlreadyCountedRows;

acFindAlreadyCountedRows = function findAlreadyCountedRowsForReportDay(workbook) {
  const reportDate = acGetReportCountDate(state.workbook);
  const weekdaySheet = acGetWeekdaySheetName(workbook, reportDate);

  if (!weekdaySheet) {
    acSetStatus(
      reportDate
        ? `No ${reportDate.toLocaleDateString("en-US", { weekday: "long" })} worksheet was found in the Already Cycle Counted file.`
        : "The report date could not be detected. Select a Cycle Count Detail file with Count Date values.",
      true
    );
    return [];
  }

  const singleSheetWorkbook = {
    SheetNames: [weekdaySheet],
    Sheets: { [weekdaySheet]: workbook.Sheets[weekdaySheet] },
  };

  const rows = acFindAlreadyCountedRowsAllSheets(singleSheetWorkbook);
  alreadyCountedState.selectedDay = weekdaySheet;
  alreadyCountedState.reportDate = reportDate;
  return rows;
};

const acMatchFilesBeforeDayStatus = acMatchFiles;
acMatchFiles = function matchFilesForSingleDay() {
  acMatchFilesBeforeDayStatus();

  if (alreadyCountedState.applied && alreadyCountedState.selectedDay) {
    const locationTotal = alreadyCountedState.matchedRows.reduce(
      (sum, row) => sum + Number(row.locationCount || 0),
      0
    );
    acSetStatus(
      `${alreadyCountedState.selectedDay} only • ${alreadyCountedState.matchedRows.length} items matched • ${locationTotal} unique locations credited`,
      false
    );
  }
};
