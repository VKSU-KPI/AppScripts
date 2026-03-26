function onEdit(e) {
    // Check if the event object and range exist
    if (!e || !e.range) return;

    let sheet = e.source.getActiveSheet();
    let range = e.range;

    if (sheet.getName() === DASHBOARD_SHEET_NAME && range.getA1Notation() === ELECTION_TYPE_CELL) {
        let election_type = range.getValue();
        let faculty_cell = sheet.getRange(FACULTY_TYPE_CELL);

        // Check if the selected election type is in the array
        if (FACULTY_REQUIRED_BY.indexOf(election_type) !== -1) {
            // Fetch the list of faculties from the range
            let faculty_list_data = sheet.getRange(FACULTY_LIST_CELL_RANGE).getValues();

            // Flatten the 2D array and filter out any empty rows
            let faculty_list = faculty_list_data.map(function(row) {
                return row[0];
            }).filter(String);

            // Create and apply the dropdown list
            let rule = SpreadsheetApp.newDataValidation()
                .requireValueInList(faculty_list, true)
                .build();
            faculty_cell.setDataValidation(rule);
        } else {
            // Remove the dropdown and clear the cell for elections without faculties
            faculty_cell.clearDataValidations();
            faculty_cell.clearContent();
        }
    }
}