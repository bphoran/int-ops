const MAIN_SHEET_ID = 76279088;
const ANALYSIS_SHEET_ID = 2060640258;

const ACCEPTANCE_SHEET_NAME = "Copy of Acceptance - May Lock";
const ANALYSIS_SHEET_NAME = "Responses Test (May Lock)";

// Response table header
const HEADER_ROW = 6;
const HEADER_COL_COUNT = 5;

// Date cells
const DATE_CELLS = "F5,K5,P5,U5,Z5,AE5,AJ5,AO5,AT5,AY5,BD5,BI5";
const DATE_COLUMNS = [6, 11, 16, 21, 26, 31, 36, 41, 46, 51, 56, 61];

// Protected ranges
const PROTECTED_RANGES =
  "A:E,AC7:AC300,AE7:AE300,AH7:AH300,AJ7:AJ300," +
  "AM7:AM300,AO7:AO300,AR7:AR300,AT7:AT300,AW7:AW300,AY7:AY300," +
  "BB7:BB300,BD7:BD300,BG7:BG300,BI7:BI300,BL7:BL300,BN7:BQ300," +
  "F7:F300,I7:I300,K7:K300,N7:N300,P7:P300,S7:S300,U7:U300,X7:X300," +
  "Z7:Z300,F6:BQ6,F1:BQ3,F5:BS5";

// Column groups
const COL_GROUP_INDICES = [8, 13, 18, 23, 28, 33, 38, 43, 48, 53, 58, 63];

// Ratio ranges
const RATIO_RANGES =
  "I7:I298,N7:N298,S7:S298,X7:X298,AC7:AC298,AH7:AH298," +
  "AM7:AM298,AR7:AR298,AW7:AW298,BB7:BB298,BG7:BG298,BL7:BL298";

// Clearable ranges
const CLEARABLE_RANGES =
  "A7:H300,J7:M300,O7:R300,T7:W300,Y7:AB300,AD7:AG300," +
  "AI7:AL300,AN7:AQ300,AS7:AV300,AX7:BA300,BC7:BF300,BH7:BK300,BM7:BQ300";

const LAYOUT = {
  headers: [
    "Platform",
    "Partner",
    "Shore",
    "Channel",
    "Staff Group",
    "Ecosystem",
    "Core/Premium",
    "Week",
    "Demand",
    "Partner Supply",
  ],

  // Demand & Partner Supply ranges
  dataRanges: [
    "F7:G",
    "K7:L",
    "P7:Q",
    "U7:V",
    "Z7:AA",
    "AE7:AF",
    "AJ7:AK",
    "AO7:AP",
    "AT7:AU",
    "AY7:AZ",
    "BD7:BE",
    "BI7:BJ",
  ],

  // Update before importing responses
  dataRowCount: 0,

  getLeftFilterRange() {
    return "A7:E" + (6 + this.dataRowCount);
  },

  getRightFilterRange() {
    return "BO7:BP" + (6 + this.dataRowCount);
  },

  // Grab column letter from dataRanges, append '5'
  getDateCell(weekNumber) {
    return this.dataRanges[weekNumber].split("7")[0] + 5;
  },

  getDataRange(weekNumber) {
    return this.dataRanges[weekNumber] + (6 + this.dataRowCount);
  },

  getResponseRange(weekNumber, segment) {
    switch (segment) {
      case "filters":
        return (
          "A" +
          (2 + weekNumber * this.dataRowCount) +
          ":G" +
          (1 + this.dataRowCount + weekNumber * this.dataRowCount)
        );
      case "week":
        return (
          "H" +
          (2 + weekNumber * this.dataRowCount) +
          ":H" +
          (1 + this.dataRowCount + weekNumber * this.dataRowCount)
        );
      case "data":
        return (
          "I" +
          (2 + weekNumber * this.dataRowCount) +
          ":J" +
          (1 + this.dataRowCount + weekNumber * this.dataRowCount)
        );
    }
  },
};

function test() {
  LAYOUT.dataRowCount = 121;

  Logger.log(LAYOUT.getResponseRange(1, "data"));
}

/**
 * Return sheet with the given gid.
 * @param {Integer} gid Sheet Grid ID.
 */
function getSheetById(gid) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();

  for (let i in sheets) {
    if (sheets[i].getSheetId() == gid) return sheets[i];
  }

  return null;
}

/**
 * Remove column groups at the COL_GROUP_INDICES.
 */
function removeGroups() {
  let acceptanceSheet = getSheetById(MAIN_SHEET_ID);

  if (acceptanceSheet == null) return;

  let colGroup;
  for (let i in COL_GROUP_INDICES) {
    try {
      colGroup = acceptanceSheet.getColumnGroup(COL_GROUP_INDICES[i], 1);
      colGroup.remove();
    } catch (err) {
      Logger.log(err);
    }
  }
}

/**
 * Restore column groups at the COL_GROUP_INDICES.
 */
function addGroups() {
  let acceptanceSheet = getSheetById(MAIN_SHEET_ID);

  if (acceptanceSheet == null) return;

  let numRows = acceptanceSheet.getLastRow();
  removeGroups();

  for (let i in COL_GROUP_INDICES) {
    acceptanceSheet
      .getRange(1, COL_GROUP_INDICES[i] + 1, numRows, 2)
      .shiftColumnGroupDepth(1);
  }
}

function createAcceptance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName("Lock Acceptance - Template");
  const dataSheet = ss.getSheetByName("lock_accept_US");

  if (templateSheet == null || dataSheet == null) {
    Logger.log("Error: Missing template sheet and/or data sheet.");
    return;
  }

  const acceptanceSheet = ss.insertSheet({ template: templateSheet });
  const numRows = dataSheet.getRange("B2:B").getValues().filter(String).length;

  // Primary filters
  let data = dataSheet.getRange(2, 1, numRows, 5).getValues();
  acceptanceSheet.getRange(7, 1, numRows, 5).setValues(data);

  // Secondary filters
  data = dataSheet.getRange(2, 6, numRows, 4).getValues();
  acceptanceSheet.getRange(7, 66, numRows, 4).setValues(data);

  // Data columns
  for (let i = 10; i < 10 + DATE_COLUMNS.length; i++) {
    let dateColumn = DATE_COLUMNS[i - 10];

    // Dates
    data = dataSheet.getRange(1, i).getValues();
    // acceptanceSheet.getRange(4, dateColumn).setValues(data);
    acceptanceSheet.getRange(5, dateColumn).setValues(data);

    // Demand data
    data = dataSheet.getRange(2, i, numRows, 1).getValues();
    acceptanceSheet.getRange(7, dateColumn, numRows, 1).setValues(data);
  }

  acceptanceSheet.getRange("A1").activate();
}

function importResponses() {
  // let ss = SpreadsheetApp.getActiveSpreadsheet();
  // // let analysisSheet = ss.getSheetByName('Lock Acceptance - Analysis');

  // let acceptanceSheet = null;
  // let analysisSheet = null;
  // let sheets = ss.getSheets();

  // for (let i in sheets) {
  //   if (sheets[i].getSheetId() == MAIN_SHEET_ID)
  //     acceptanceSheet = sheets[i];
  //   else if (sheets[i].getSheetId() == ANALYSIS_SHEET_ID)
  //     analysisSheet = sheets[i];
  // }

  // let acceptanceSheet = getSheetById(MAIN_SHEET_ID);
  // let analysisSheet = getSheetById(ANALYSIS_SHEET_ID);

  const acceptanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    ACCEPTANCE_SHEET_NAME
  );

  const responseSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ANALYSIS_SHEET_NAME);

  if (acceptanceSheet == null || responseSheet == null) return;

  // Logger.log('Acceptance Sheet: ' + acceptanceSheet.getSheetName());
  // Logger.log('Analysis Sheet: ' + analysisSheet.getSheetName());

  // let filters = acceptanceSheet.getRange('A6:I' + acceptanceSheet.getLastRow()).getValues();

  // Find number of data rows
  // let numRows = acceptanceSheet.getRange('B7:B').getValues().filter(String).length;
  // Logger.log('numRows: ' + numRows);

  // let headers = acceptanceSheet.getRange(HEADER_ROW, 1, 1, HEADER_COL_COUNT).getValues();
  // let filters = acceptanceSheet.getRange(HEADER_ROW + 1, 1,
  //   acceptanceSheet.getLastRow(), HEADER_COL_COUNT).getValues();
  // Logger.log('filters.length: ' + filters.length);
  // Logger.log('filters[0].length: ' + filters[0].length);

  // let headers = ['Platform', 'Partner', 'Shore', 'Channel', 'Staff Group',
  //   'Ecosystem', 'Core/Premium', 'Week', 'Demand', 'Partner Supply'];

  // Logger.log('headers: ' + headers);

  // let filters = acceptanceSheet.getRange(7, 1, numRows, 5).getValues();
  // let rightFilters = acceptanceSheet.getRange(7, 67, numRows, 2).getValues();

  // TEST BEGIN
  LAYOUT.dataRowCount = acceptanceSheet
    .getRange("B7:B")
    .getValues()
    .filter(String).length;
  let filters = acceptanceSheet
    .getRange(LAYOUT.getLeftFilterRange())
    .getValues();
  let rightFilters = acceptanceSheet
    .getRange(LAYOUT.getRightFilterRange())
    .getValues();
  // TEST END

  for (let i in filters) {
    filters[i] = filters[i].concat(rightFilters[i]);
    // Logger.log(filters[i]);
  }

  // let responseHeaderRange = responseSheet.getRange(1, 1, 1, LAYOUT.headers.length);
  // let responseRange = responseSheet.getRange(2, 1, filters.length, filters[0].length);
  // Logger.log(responseRange.getA1Notation());

  // responseSheet.clear();
  // responseHeaderRange.setValues([LAYOUT.headers]);
  // responseRange.setValues(filters);

  // // Week
  // responseRange = responseSheet.getRange(2, 8, LAYOUT.dataRowCount, 1)
  //   .setValue(acceptanceSheet.getRange(LAYOUT.getDateCell(1)).getValue());

  // // Demand & Partner Supply
  // responseRange = responseSheet.getRange(2, 9, LAYOUT.dataRowCount, 2)
  //   .setValues(acceptanceSheet.getRange(LAYOUT.getDataRange(1)).getValues());

  // TEST BEGIN
  responseSheet.clear();
  responseSheet
    .getRange(1, 1, 1, LAYOUT.headers.length)
    .setValues([LAYOUT.headers]);

  for (let i in LAYOUT.dataRanges) {
    responseSheet
      .getRange(LAYOUT.getResponseRange(i, "filters"))
      .setValues(filters);
    responseSheet
      .getRange(LAYOUT.getResponseRange(i, "week"))
      .setValue(acceptanceSheet.getRange(LAYOUT.getDateCell(i)).getValues());
    responseSheet
      .getRange(LAYOUT.getResponseRange(i, "data"))
      .setValues(acceptanceSheet.getRange(LAYOUT.getDataRange(i)).getValues());
  }
  // TEST END
}

/**
 * Protect the relevant ranges on the acceptance sheet.
 */
function protectRanges() {
  // const sheet = getSheetById(1607244281);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    ACCEPTANCE_SHEET_NAME
  );

  if (sheet == null) {
    Logger.log("Error: Acceptance sheet not found.");
    return;
  }

  const protectedRanges = PROTECTED_RANGES.split(",");
  const me = Session.getEffectiveUser();

  for (let i in protectedRanges) {
    let range = sheet.getRange(protectedRanges[i]);
    let protection = range.protect();

    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());

    if (protection.canDomainEdit()) protection.setDomainEdit(false);

    Logger.log(protectedRanges[i] + ": protected");
  }
}

function removeProtections() {
  // Remove all range protections in the spreadsheet that the user has permission to edit.
  const ss = SpreadsheetApp.getActive();
  // const sheet = getSheetById(1631195170);

  let protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  for (let i in protections) {
    let protection = protections[i];
    if (
      protection.getRange().getSheet().getSheetId() == 1631195170 &&
      protection.canEdit()
    ) {
      protection.remove();
    }
  }
}

function clearSheet() {
  let acceptanceSheet = getSheetById(MAIN_SHEET_ID);

  if (acceptanceSheet == null) return;

  let clearable_ranges = CLEARABLE_RANGES.split(",");

  for (let i in clearable_ranges) {
    // Logger.log(clearable_ranges[i]);
    acceptanceSheet.getRange(clearable_ranges[i]).clearContent();
  }

  // acceptanceSheet.getRange('H7').clearContent();

  acceptanceSheet.getRange("A1").activate();
}
