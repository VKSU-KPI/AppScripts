function createElectionSheet() {
    // Get the active spreadsheet and the current dashboard sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboardSheet = ss.getActiveSheet();

    const election_type_cell = "E2";
    const faculty_type_cell = "E4";
    const faculties_cells = "A2:A24";

    // Define election types that require a faculty selection
    const requires_faculty = ["ВРп", "КСУп", "КТКп", "СРг", "СРп"];

    // Read the election type from E2 and the faculty from E4
    var electionType = dashboardSheet.getRange(election_type_cell).getValue();
    var faculty = dashboardSheet.getRange(faculty_type_cell).getValue();

    // Validate if an election type is actually selected
    if (!electionType) {
        SpreadsheetApp.getUi().alert("Спочатку оберіть тип виборів.");
        return;
    }

    var needsFaculty = requires_faculty.indexOf(electionType) !== -1;

    // Validate that a faculty is selected if the election type requires it
    if (needsFaculty && !faculty) {
        SpreadsheetApp.getUi().alert("Оберіть підрозділ, це обов'язково для цього типу виборів.");
        return;
    }

    // Construct the expected name for the new sheet
    var newSheetName = electionType;
    if (needsFaculty) {
        newSheetName = electionType + " " + faculty;
    }

    // Check if a sheet with this name already exists in the spreadsheet
    if (ss.getSheetByName(newSheetName)) {
        SpreadsheetApp.getUi().alert("Аркуш з назвою '" + newSheetName + "' вже існує!");
        return;
    }

    // Construct the expected template name based on the naming convention
    var templateName = electionType + " ШАБЛОН";
    var templateSheet = ss.getSheetByName(templateName);

    // Verify that the required template sheet actually exists
    if (!templateSheet) {
        SpreadsheetApp.getUi().alert("Шаблон '" + templateName + "' не знайдений. Спочатку створіть його.");
        return;
    }

    // Create a duplicate of the template, rename it, and bring it to the front
    var newSheet = templateSheet.copyTo(ss);
    newSheet.setName(newSheetName);
    ss.setActiveSheet(newSheet);

    // Clear the dashboard input fields and remove data validation for the next time
    dashboardSheet.getRange(election_type_cell).clearContent();
    dashboardSheet.getRange(faculty_type_cell).clearContent();
    dashboardSheet.getRange(faculty_type_cell).clearDataValidations();
}