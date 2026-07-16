# Cycle Count Production

A browser-based Excel utility that calculates cycle count production from aisle assignments and writes the totals into an existing weekly tracker.

## Aisle assignments

| Employee | Aisles |
|---|---|
| Carico | A-B |
| Ernie | C-D |
| Cherish | E-F |
| Layne | G-H |
| Madison | I-J |
| Antoine | K-L |

The daily goal is **200 cycle counts = 100% production**.

## How it works

1. Open `index.html` in a browser or publish the repository with GitHub Pages.
2. Upload the Cycle Count Detail Excel report.
3. Confirm the report sheet and Bin / Location column.
4. Review the totals and any unassigned locations.
5. Upload the weekly Excel tracker.
6. Select the day sheet, header row, employee column, cycle count column, and optional production percentage column.
7. Click **Update and download tracker**.

The original tracker file is never overwritten. The app downloads a new file ending in `_UPDATED_YYYY-MM-DD.xlsx`.

## Privacy

The Excel files are processed locally in the browser. They are not uploaded to this repository or to a server.

## Notes

- The source report should contain a bin/location value beginning with an aisle letter, such as `C-10-3`.
- Rows outside aisles A-L are shown under **Unassigned rows need review**.
- The app supports `.xlsx` and `.xls` files. Macro-enabled or unusually complex workbooks should be tested on a copy first.
