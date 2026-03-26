function createElectionSheet() {
  // Get the active spreadsheet and the current dashboard sheet
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let dashboard_sheet = ss.getActiveSheet();

  // Verify we are running this script from the correct dashboard sheet
  if (dashboard_sheet.getName() !== DASHBOARD_SHEET_NAME) {
    SpreadsheetApp.getUi().alert("Цей скрипт можна запускати лише з головного дашборду.");
    return;
  }

  // Read input parameters
  let {
    election_type,
    faculty,
    dormitory
  } = getInputParameters(dashboard_sheet);

  // Validate if an election type is actually selected
  if (!election_type) {
    SpreadsheetApp.getUi().alert("Спочатку оберіть тип виборів.");
    return;
  }

  let is_faculty_needed = FACULTY_REQUIRED_BY.indexOf(election_type) !== -1;
  let is_dormitory_needed = DORMITORY_REQUIRED_BY.indexOf(election_type) !== -1;

  // Validate that a faculty is selected if required
  if (is_faculty_needed && !faculty) {
    SpreadsheetApp.getUi().alert("Оберіть підрозділ, це обов'язково для цього типу виборів.");
    return;
  }
  // Validate that a dormitory is selected if required
  if (is_dormitory_needed && !dormitory) {
    SpreadsheetApp.getUi().alert("Оберіть гуртожиток, це обов'язково для цього типу виборів.");
    return;
  }

  // Construct the expected name for the new sheet
  let new_sheet_name = election_type;
  if (is_faculty_needed) {
    new_sheet_name = election_type + " " + faculty;
  } else if (is_dormitory_needed) {
    new_sheet_name = election_type + " #" + dormitory;
  }

  // Check if a sheet with this name already exists in the spreadsheet
  if (ss.getSheetByName(new_sheet_name)) {
    SpreadsheetApp.getUi().alert("Аркуш з назвою '" + new_sheet_name + "' вже існує!");
    return;
  }

  // Construct the expected template name based on the naming convention
  let template_name = election_type + " ШАБЛОН";
  let template_sheet = ss.getSheetByName(template_name);

  // Verify that the required template sheet actually exists
  if (!template_sheet) {
    SpreadsheetApp.getUi().alert("Шаблон '" + template_name + "' не знайдений. Спочатку створіть його.");
    return;
  }

  // Create a duplicate of the template, rename it, and bring it to the front
  let new_sheet = template_sheet.copyTo(ss);
  new_sheet.setName(new_sheet_name);
  ss.setActiveSheet(new_sheet);

  // Call the indexer function to add a hyperlink to the main sheet
  let link_title;
  if (is_faculty_needed) {
    link_title = faculty;
  } else if (is_dormitory_needed) {
    link_title = dormitory;
  } else {
    link_title = election_type;
  }
  addLinkToMainSheet(link_title, election_type, new_sheet.getSheetId());

  // Clear the dashboard input fields and remove data validation for the next time
  clearInputCells(dashboard_sheet);
}
