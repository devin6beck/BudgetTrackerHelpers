function TransferBankingDataToActiveSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const chaseSheets = findSheets(ss, 'chase');
    const verizonSheets = findSheets(ss, 'verizon');
    const boaSheets = findSheets(ss, ['boa', 'bankofamerica']); // New BoA sheets
    const capitalSheets = findSheets(ss, 'capital'); // New Capital sheets
    const activeSheet = ss.getActiveSheet();
    const monthNum = getMonthNumFromSheet(activeSheet.getName());

    if (!monthNum) {
      throw new Error('The active sheet name must be the name of a Month (January - December) example: "June"');
    }

    if (
      chaseSheets.length === 0 && 
      verizonSheets.length === 0 &&
      boaSheets.length === 0 &&
      capitalSheets.length === 0
    ) {
      throw new Error("No banking sheet was found with the name 'Chase', 'Verizon', 'BoA', 'BankOfAmerica' or 'Capital'");
    }

    chaseSheets.forEach(chaseSheet => transferDataFromSheet(chaseSheet, activeSheet, monthNum, "Chase"));
    verizonSheets.forEach(verizonSheet => transferDataFromSheet(verizonSheet, activeSheet, monthNum, "Verizon"));
    boaSheets.forEach(boaSheet => transferDataFromSheet(boaSheet, activeSheet, monthNum, "BoA"));
    capitalSheets.forEach(capitalSheet => transferDataFromSheet(capitalSheet, activeSheet, monthNum, "Capital"));
    
  } catch (error) {
    console.error("An error occurred: ", error);
    SpreadsheetApp.getUi().alert(error.message);
  }
}