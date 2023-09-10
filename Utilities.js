
function findSheets(ss, sheetNamePrefixes) {
  const sheets = ss.getSheets();
  if (typeof sheetNamePrefixes === 'string') {
    sheetNamePrefixes = [sheetNamePrefixes];
  }
  return sheets.filter(sheet => sheetNamePrefixes.some(prefix => sheet.getName().toLowerCase().startsWith(prefix)));
}

function getSheetName() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return sheet.getName();
}


function getMonthNumFromSheet(sheetName) {
  const monthIndex = months.findIndex(month => sheetName.toLowerCase().startsWith(month.toLowerCase()));
  return monthIndex !== -1 ? (monthIndex + 1).toString().padStart(2, '0') : null;
}

function convertDatesToText(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (values[i][j] instanceof Date) {
        values[i][j] = Utilities.formatDate(values[i][j], Session.getScriptTimeZone(), "MM/dd/yy");
      }
    }
  }

  range.setValues(values);
}

function getHeaders(sheet, requiredHeaders) {
  const headers = sheet.getRange("A1:G1").getValues()[0];
  const result = [];

  requiredHeaders.forEach(header => {
    const index = headers.indexOf(header);
    if (index === -1) {
      throw new Error(`Header '${header}' not found`);
    }
    result.push({ name: header, index: index });
  });

  return result;
}

function getRequiredHeaders(bankName) {
  switch (bankName) {
    case "BoA":
      return ["Posted Date", "Payee", "Amount"];
    case "Capital":
      return ["Transaction Date", "Description", "Debit"];
    default:
      return ["Transaction Date", "Description", "Amount"];
  }
}

function getCategory(description, bankName) {
  description = description.toLowerCase();
  const utilitiesData = getUtilitiesData();
  
  // Check if the description exists in the Utilities sheet
  for (const [utilDescription, category] of Object.entries(utilitiesData)) {
    if (description.includes(utilDescription)) {
      return category;
    }
  }

  return ""; // Default category if none of the above conditions are met
}

function getUtilitiesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const utilitiesSheet = ss.getSheetByName("Utilities");
  if (!utilitiesSheet) {
    throw new Error('Utilities sheet not found');
  }
  const data = utilitiesSheet.getDataRange().getValues();
  let categoryColumnIndex = -1;
  let descriptionColumnIndex = -1;
  for (let i = 0; i < data[0].length; i++) {
    if (data[0][i] === "Category for Description") {
      categoryColumnIndex = i;
    }
    if (data[0][i] === "Description") {
      descriptionColumnIndex = i;
    }
  }
  if (categoryColumnIndex === -1 || descriptionColumnIndex === -1) {
    throw new Error('Category or Description column not found in Utilities sheet');
  }
  const utilitiesData = {};
  for (let i = 1; i < data.length; i++) {
    utilitiesData[data[i][descriptionColumnIndex].toLowerCase()] = data[i][categoryColumnIndex];
  }
  return utilitiesData;
}
