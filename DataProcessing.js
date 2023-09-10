function findHeadersAndLastRow(sheet, requiredHeaders) {
  const data = sheet.getDataRange().getValues();
  let activeHeaders, lastRow;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    if (requiredHeaders.every(header => row.includes(header))) {
      activeHeaders = row;
      const indices = requiredHeaders.map(header => activeHeaders.indexOf(header));
      lastRow = i;

      for (let j = i + 1; j < data.length; j++) {
        const lastRowData = data[j];
        if (indices.some(index => lastRowData[index])) {
          lastRow = j;
        }
      }

      return { activeHeaders, lastRow: lastRow + 2 };
    }
  }
  
  throw new Error("No header row found");
}

function buildExistingTransactionsDict(data, headers) {
  const existingTransactions = {};
  const dateColumn = headers.indexOf('Date');
  const amountColumn = headers.indexOf('Amount');

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let date = new Date(row[dateColumn]);
    let formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yy");
    let amount = parseFloat(row[amountColumn]);
    let key = `${formattedDate}_${amount.toFixed(2)}`;
    console.log("Existing key: ", key); // Debugging line
    existingTransactions[key] = true;
  }

  return existingTransactions;
}


function processBankData(bankData, bankHeaders, activeSheet, activeHeaders, lastRow, monthNum, bankName, bankSheetName) {
  let activeData = activeSheet.getDataRange().getValues();
  const startColumn = activeHeaders.indexOf('Date') + 1;
  const headersLength = ["Date", "Amount", "Category", "Description"].length;
  let existingTransactions = buildExistingTransactionsDict(activeData, activeHeaders);

  // Extract the part of the bank name up to the first non-alpha character
  const bankNameSuffix = bankSheetName.match(/[a-zA-Z0-9]+/)[0];

  for (let i = 1; i < bankData.length; i++) {
    const row = bankData[i];
    let amountHeader = bankName === "Capital" ? 'Debit' : 'Amount';
    let amount = parseFloat(row[bankHeaders.find(h => h.name === amountHeader).index]);

    if (isNaN(amount)) { // Skip rows with empty Debit/Amount field
      continue;
    }

    // Skip negative amounts for Capital sheets
    if (bankName !== "Capital" && amount >= 0) {
      continue;
    }

    // Make the amount negative for non-Capital sheets
    if (bankName === "Capital") {
      amount = -Math.abs(amount);
    }

    let dateColumn = bankName === "BoA" ? "Posted Date" : "Transaction Date"; // Adjust date column name for BoA
    let date = new Date(row[bankHeaders.find(h => h.name === dateColumn).index]);
    let formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yy");

    if (formattedDate.substring(0, 2) !== monthNum) {
      continue;
    }

    let descriptionColumn = bankName === "BoA" ? "Payee" : "Description"; // BoA uses 'Payee' instead of 'Description'
    let description = row[bankHeaders.find(h => h.name === descriptionColumn).index];

    // Append the extracted part of the bank sheet name to the description
    description += " - " + bankNameSuffix;

    // Check for duplicates
    const key = `${formattedDate}_${amount.toFixed(2)}`; // No description in the key
    console.log("Checking key: ", key); // Debugging line
    const duplicateExists = existingTransactions.hasOwnProperty(key);
    if (duplicateExists) {
      console.log("Duplicate found for key: ", key); // Debugging line
    } else {
      const newRow = new Array(headersLength).fill('');
      newRow[activeHeaders.indexOf('Date') - startColumn + 1] = formattedDate;
      newRow[activeHeaders.indexOf('Amount') - startColumn + 1] = amount.toFixed(2);
      newRow[activeHeaders.indexOf('Category') - startColumn + 1] = getCategory(description, bankName);
      newRow[activeHeaders.indexOf('Description') - startColumn + 1] = description;

      activeSheet.getRange(lastRow, startColumn, 1, newRow.length).setValues([newRow]);
      lastRow++;
      activeData = activeSheet.getDataRange().getValues();
      existingTransactions[key] = true;
    }
  }
}
