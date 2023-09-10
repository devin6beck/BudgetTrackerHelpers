function transferDataFromSheet(bankSheet, activeSheet, monthNum, bankName) {
  convertDatesToText(bankSheet);
  const bankHeaders = getHeaders(bankSheet, getRequiredHeaders(bankName)); // Use dynamic headers
  const bankData = bankSheet.getDataRange().getDisplayValues();
  const { activeHeaders, lastRow } = findHeadersAndLastRow(activeSheet, ["Date", "Amount", "Category", "Description"]);
  processBankData(bankData, bankHeaders, activeSheet, activeHeaders, lastRow, monthNum, bankName, bankSheet.getName()); // Pass bankName and bankSheet.getName()
}

