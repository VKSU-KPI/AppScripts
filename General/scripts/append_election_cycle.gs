function appendElectionCycle() {
  // Get the active spreadsheet and dashboard sheet
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
    faculty
  } = getInputParameters(dashboard_sheet);

  // Validate if an election type is selected
  if (!election_type) {
    SpreadsheetApp.getUi().alert("Спочатку оберіть тип виборів.");
    return;
  }

  let is_faculty_needed = FACULTY_REQUIRED_BY.indexOf(election_type) !== -1;

  // Validate that a faculty is selected if required
  if (is_faculty_needed && !faculty) {
    SpreadsheetApp.getUi().alert("Оберіть підрозділ, це обов'язково для цього типу виборів.");
    return;
  }

  // Construct the target sheet name and template name
  let target_sheet_name = election_type;
  if (is_faculty_needed) {
    target_sheet_name = election_type + " " + faculty;
  }
  let template_name = election_type + " ШАБЛОН";

  // Fetch the target and template sheets
  let target_sheet = ss.getSheetByName(target_sheet_name);
  let template_sheet = ss.getSheetByName(template_name);

  // Verify both sheets exist before proceeding
  if (!target_sheet) {
    SpreadsheetApp.getUi().alert("Аркуш '" + target_sheet_name + "' не існує. Створіть його спочатку.");
    return;
  }
  if (!template_sheet) {
    SpreadsheetApp.getUi().alert("Шаблон '" + template_name + "' не знайдений.");
    return;
  }

  // Dynamically get all content from the template sheet
  let source_range = template_sheet.getDataRange();

  // Find the last row with content in the target sheet
  let last_row = target_sheet.getLastRow();

  // Calculate the starting row for the new table with spacing
  let start_row_for_new_table = last_row > 0 ? last_row + 1 : 1;
  let destination_cell = target_sheet.getRange(start_row_for_new_table, 1);

  // Copy the template to the destination
  source_range.copyTo(destination_cell);

  // MAGIC FIX FOR CONDITIONAL FORMATTING
  // Calculate how many rows down the new table was pasted
  let row_offset = start_row_for_new_table - source_range.getRow();

  // Get all conditional formatting rules from the target sheet
  let rules = target_sheet.getConditionalFormatRules();
  let updated_rules = [];

  for (let i = 0; i < rules.length; i++) {
    let rule = rules[i];
    let boolean_condition = rule.getBooleanCondition();

    // Check if it's a custom formula and contains a fully absolute reference (like $B$3)
    let is_custom_formula = boolean_condition && boolean_condition.getCriteriaType() == SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA;
    let has_absolute_reference = false;
    let formula = "";

    if (is_custom_formula) {
      formula = boolean_condition.getCriteriaValues()[0];
      has_absolute_reference = /\$[A-Za-z]+\$\d+/.test(formula);
    }

    // If it has an absolute reference, we split the rule and shift the rows
    if (is_custom_formula && has_absolute_reference) {
      let ranges = rule.getRanges();
      let old_ranges = [];
      let new_ranges = [];

      // Separate ranges belonging to old tables from the newly pasted one
      for (let j = 0; j < ranges.length; j++) {
        if (ranges[j].getRow() >= start_row_for_new_table) {
          new_ranges.push(ranges[j]);
        } else {
          old_ranges.push(ranges[j]);
        }
      }

      if (new_ranges.length > 0) {
        // Shift absolute row references by the calculated offset
        let new_formula = formula.replace(/(\$[A-Za-z]+)\$(\d+)/g, function(match, col_ref, row_num_str) {
          let old_row_num = parseInt(row_num_str, 10);
          return col_ref + "$" + (old_row_num + row_offset);
        });

        let new_rule = rule.copy().setRanges(new_ranges).whenFormulaSatisfied(new_formula).build();
        updated_rules.push(new_rule);
      }

      if (old_ranges.length > 0) {
        let old_rule = rule.copy().setRanges(old_ranges).build();
        updated_rules.push(old_rule);
      }
    } else {
      // If there is no absolute reference (like $G3), keep the rule exactly as is
      // Google Sheets natively handles relative references correctly across merged ranges
      updated_rules.push(rule);
    }
  }

  // Apply the corrected rules back to the sheet
  target_sheet.setConditionalFormatRules(updated_rules);

  // Clear the dashboard input fields and remove data validation for the next time
  clearInputCells(dashboard_sheet);

  // Activate the target sheet
  ss.setActiveSheet(target_sheet);
}