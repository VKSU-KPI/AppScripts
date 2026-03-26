function getInputParameters(sheet) {
  // Verify we are running this script from the correct dashboard sheet
  if (sheet.getName() !== DASHBOARD_SHEET_NAME) {
    SpreadsheetApp.getUi().alert("Цей скрипт можна запускати лише з головного дашборду.");
    return;
  }

  let election_type = sheet.getRange(ELECTION_TYPE_CELL).getValue();
  let faculty = sheet.getRange(FACULTY_TYPE_CELL).getValue();

  return {
    election_type: election_type,
    faculty: faculty,
  }
}


function clearInputCells(sheet) {
  // Verify we are running this script from the correct dashboard sheet
  if (sheet.getName() !== DASHBOARD_SHEET_NAME) {
    SpreadsheetApp.getUi().alert("Цей скрипт можна запускати лише з головного дашборду.");
    return;
  }

  // Clear the dashboard input fields and remove data validation for the next time
  sheet.getRange(ELECTION_TYPE_CELL).clearContent();
  sheet.getRange(FACULTY_TYPE_CELL).clearContent();
  sheet.getRange(FACULTY_TYPE_CELL).clearDataValidations();
}