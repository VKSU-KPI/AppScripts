function onEdit(e) {
    // Check if the event object and range exist
    if (!e || !e.range) return;

    var sheet = e.source.getActiveSheet();
    var range = e.range;

    const election_type_cell = "E4";
    const faculty_type_cell = "E6";
    const faculties_cells = "A2:A24";

    // Define election types that require a faculty selection
    const requires_faculty = ["ВРп", "КСУп", "КТКп", "СРг", "СРп"];

    // Verify the edited cell is E4, which represents the merged E4:E5 cell
    if (range.getA1Notation() === election_type_cell) {
        var electionType = range.getValue();
        var facultyCell = sheet.getRange(faculty_type_cell);

        // Check if the selected election type is in the array
        if (requires_faculty.indexOf(electionType) !== -1) {
            // Fetch the list of faculties from the range A2:A24
            var facultyData = sheet.getRange(faculties_cells).getValues();

            // Flatten the 2D array and filter out any empty rows
            var facultyList = facultyData.map(function(row) {
                return row[0];
            }).filter(String);

            // Create and apply the dropdown list
            var rule = SpreadsheetApp.newDataValidation()
                .requireValueInList(facultyList, true)
                .build();
            facultyCell.setDataValidation(rule);
        } else {
            // Remove the dropdown and clear the cell for elections without faculties
            facultyCell.clearDataValidations();
            facultyCell.clearContent();
        }
    }
}