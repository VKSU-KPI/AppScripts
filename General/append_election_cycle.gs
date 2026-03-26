function appendElectionTable() {
    // Get the active spreadsheet and dashboard sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboardSheet = ss.getActiveSheet();

    const election_type_cell = "E2";
    const faculty_type_cell = "E4";
    const faculties_cells = "A2:A24";

    // Define election types that require a faculty selection
    const requires_faculty = ["ВРп", "КСУп", "КТКп", "СРг", "СРп"];

    // Read the election type and faculty from the dashboard
    var electionType = dashboardSheet.getRange(election_type_cell).getValue();
    var faculty = dashboardSheet.getRange(faculty_type_cell).getValue();

    // Validate if an election type is selected
    if (!electionType) {
        SpreadsheetApp.getUi().alert("Спочатку оберіть тип виборів.");
        return;
    }

    var needsFaculty = requires_faculty.indexOf(electionType) !== -1;

    // Validate that a faculty is selected if required
    if (needsFaculty && !faculty) {
        SpreadsheetApp.getUi().alert("Оберіть підрозділ, це обов'язково для цього типу виборів.");
        return;
    }

    // Construct the target sheet name and template name
    var targetSheetName = electionType;
    if (needsFaculty) {
        targetSheetName = electionType + " " + faculty;
    }
    var templateName = electionType + " ШАБЛОН";

    // Fetch the target and template sheets
    var targetSheet = ss.getSheetByName(targetSheetName);
    var templateSheet = ss.getSheetByName(templateName);

    // Verify both sheets exist before proceeding
    if (!targetSheet) {
        SpreadsheetApp.getUi().alert("Аркуш '" + targetSheetName + "' не існує. Створіть його спочатку.");
        return;
    }
    if (!templateSheet) {
        SpreadsheetApp.getUi().alert("Шаблон '" + templateName + "' не знайдений.");
        return;
    }

    // Dynamically get all content from the template sheet
    var sourceRange = templateSheet.getDataRange();

    // Find the last row with content in the target sheet
    var lastRow = targetSheet.getLastRow();

    // Calculate the starting row for the new table with spacing
    var startRowForNewTable = lastRow > 0 ? lastRow + 1 : 1;
    var destinationCell = targetSheet.getRange(startRowForNewTable, 1);

    // Copy the template to the destination
    sourceRange.copyTo(destinationCell);

    // MAGIC FIX FOR CONDITIONAL FORMATTING
    // Calculate how many rows down the new table was pasted
    var rowOffset = startRowForNewTable - sourceRange.getRow();

    // Get all conditional formatting rules from the target sheet
    var rules = targetSheet.getConditionalFormatRules();
    var updatedRules = [];

    for (var i = 0; i < rules.length; i++) {
        var rule = rules[i];
        var booleanCondition = rule.getBooleanCondition();

        // Check if it's a custom formula and contains a fully absolute reference (like $B$3)
        var isCustomFormula = booleanCondition && booleanCondition.getCriteriaType() == SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA;
        var hasAbsoluteReference = false;
        var formula = "";

        if (isCustomFormula) {
            formula = booleanCondition.getCriteriaValues()[0];
            hasAbsoluteReference = /\$[A-Za-z]+\$\d+/.test(formula);
        }

        // If it has an absolute reference, we split the rule and shift the rows
        if (isCustomFormula && hasAbsoluteReference) {
            var ranges = rule.getRanges();
            var oldRanges = [];
            var newRanges = [];

            // Separate ranges belonging to old tables from the newly pasted one
            for (var j = 0; j < ranges.length; j++) {
                if (ranges[j].getRow() >= startRowForNewTable) {
                    newRanges.push(ranges[j]);
                } else {
                    oldRanges.push(ranges[j]);
                }
            }

            if (newRanges.length > 0) {
                // Shift absolute row references by the calculated offset
                var newFormula = formula.replace(/(\$[A-Za-z]+)\$(\d+)/g, function(match, colRef, rowNumStr) {
                    var oldRowNum = parseInt(rowNumStr, 10);
                    return colRef + "$" + (oldRowNum + rowOffset);
                });

                var newRule = rule.copy().setRanges(newRanges).whenFormulaSatisfied(newFormula).build();
                updatedRules.push(newRule);
            }

            if (oldRanges.length > 0) {
                var oldRule = rule.copy().setRanges(oldRanges).build();
                updatedRules.push(oldRule);
            }
        } else {
            // If there is no absolute reference (like $G3), keep the rule exactly as is
            // Google Sheets natively handles relative references correctly across merged ranges
            updatedRules.push(rule);
        }
    }

    // Apply the corrected rules back to the sheet
    targetSheet.setConditionalFormatRules(updatedRules);

    // Clear the dashboard inputs
    dashboardSheet.getRange(election_type_cell).clearContent();
    dashboardSheet.getRange(faculty_type_cell).clearContent();
    dashboardSheet.getRange(faculty_type_cell).clearDataValidations();

    // Activate the target sheet
    ss.setActiveSheet(targetSheet);
}