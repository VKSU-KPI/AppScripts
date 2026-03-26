// Dashboard sheet cells
const ELECTION_TYPE_CELL = "F2";
const FACULTY_TYPE_CELL = "F4";
const DORMITORY_TYPE_CELL = "F6";

const FACULTY_LIST_CELL_RANGE = "A2:A24";
const DORMITORY_LIST_CELL_RANGE = "C2:C22";

// Election types that require faculty
const FACULTY_REQUIRED_BY = ["ВРп", "КСУп", "КТКп", "СРп"];
// Election types that require dormitory
const DORMITORY_REQUIRED_BY = ["СРг"];

// Sheets
const MAIN_SHEET_NAME = "Головна";
const DASHBOARD_SHEET_NAME = "Дашборд";

// History
const HISTORY_UNIQUE_CELLS = {
  "ВР КПІ": "B3",
  "Президента": "B4"
};
const HISTORY_CELL_RANGES = {
  "ВРп": "C3:C100",
  "КСУп": "D3:D100",
  "КТКп": "E3:E100",
  "СРг": "F3:F100",
  "СРп": "G3:G100"
};
