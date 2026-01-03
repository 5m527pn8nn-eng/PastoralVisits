This repo is a small Google Apps Script project for a pastoral-care Google Sheet.
It aggregates counts across facility sheets, builds history reports, and can
generate a printable visit report PDF from a manual selection.

Files
- Code.js: all Apps Script logic.
- appsscript.json: Apps Script config (timezone, V8 runtime, logging).
- pastoralCareVisitsPromptForCodeGenForAutoSelecOfVisitsSchedule.txt: a prompt
  spec for generating a future script that auto-selects institutions; it is not
  executable code.

Main workflow in Code.js
- reportingObject: simple data holder for a resident's institution/name/room/
  phone/date/comments.
- _sum(): core aggregator.
  - Iterates all sheets and skips specific ones by name (e.g., "Report",
    "History", "Eucharistic Ministers", etc.).
  - For each facility sheet, reads totals from E1:I1 (residents, Eucharists,
    anointings, etc.).
  - Builds a list of residents after the "Anointing of the Sick (Date only)"
    marker row, with special column offsets for the "Homebound" sheet.
  - Writes totals into the "Report" sheet.
  - Builds two history sheets: "History" and "History by Last Seen Overall"
    via runHistoryReports(), with different sort orders.
- runHistoryReports(sheet, reportingArray): clears a target sheet, writes
  headers, filters out "no visit" residents or missing room, and writes rows.
- findLastDateSeen(sheet, row, ifHomeBoundSheet): scans multiple date columns
  and returns the latest date.
- buildAndEmailReportFromSelection(): takes the user's selected rows, builds a
  "Pastoral Care Visit Report" sheet with extra checkbox columns, formats it
  for printing, exports to PDF, and emails it.
- exportSheetToPdfBlob_(): handles the PDF export via the Sheets export URL.
- onOpen(): adds a "Sheet Tools" custom menu to run _sum() and the PDF report
  builder.
- Utility/testing helpers: testOnEdit() (simulates an edit event), getDate(),
  and sheet-sorting helpers.
