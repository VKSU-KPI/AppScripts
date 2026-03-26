function addLinkToMainSheet(title, election_type, sheet_id) {
  // Get the active spreadsheet and the main index sheet
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let main_sheet = ss.getSheetByName(MAIN_SHEET_NAME);

  // Return early if the main sheet doesn't exist
  if (!main_sheet) return;

  // Forming title for dormitory
  if (DORMITORY_REQUIRED_BY.indexOf(election_type) !== -1) {
    title = "#" + title;
  }

  // Construct the hyperlink formula using the sheet's unique GID
  let link_formula = '=HYPERLINK("#gid=' + sheet_id + '"; "' + title + '")';

  if (HISTORY_UNIQUE_CELLS[election_type]) {
    // Handle unique elections with fixed single cells
    let target_cell = main_sheet.getRange(HISTORY_UNIQUE_CELLS[election_type]);

    // Only set the formula if the cell doesn't already display this sheet name
    if (target_cell.getValue() !== title) {
      target_cell.setFormula(link_formula);
    }
  } else if (HISTORY_CELL_RANGES[election_type]) {
    // Handle faculty elections with column ranges
    let target_range = main_sheet.getRange(HISTORY_CELL_RANGES[election_type]);
    let values = target_range.getValues();

    let already_exists = false;
    let first_empty_row_index = -1;

    // Scan the range for duplicates and find the first empty slot
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === title) {
        already_exists = true;
        break;
      }
      if (values[i][0] === "" && first_empty_row_index === -1) {
        first_empty_row_index = i;
      }
    }

    // If it's a new entry and we have space, insert and sort
    if (!already_exists && first_empty_row_index !== -1) {
      let row_offset = target_range.getRow() + first_empty_row_index;
      let col_offset = target_range.getColumn();

      let cell_to_write = main_sheet.getRange(row_offset, col_offset);
      cell_to_write.setFormula(link_formula);

      // Force the spreadsheet to apply all pending changes before sorting
      SpreadsheetApp.flush();

      // Sort the specific range alphabetically 
      target_range.sort({
        column: col_offset,
        ascending: true
      });
    }
  }
}
