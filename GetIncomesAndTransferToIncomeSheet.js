function GetIncomesAndTransferToIncomeSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const incomeSheet = ss.getSheetByName("Income - Year");

    // If the "Income - Year" sheet does not exist, throw an error.
    if (!incomeSheet) {
        throw new Error("Income - Year sheet not found");
    }

    const lastRow = incomeSheet.getLastRow();
    if (lastRow >= 31) {
        // Clear the existing content of the "Income - Year" sheet starting from row 31.
        incomeSheet.getRange(31, 1, lastRow - 30, incomeSheet.getLastColumn()).clearContent();
    }

    // Append header row to the sheet at row 31.
    incomeSheet.getRange(31, 1, 1, 5).setValues([['Month', 'D&H Beck', 'Hannah', 'Devin', 'Total']]); // Added 'Total'

    let monthIncomeData = {};

    // Initialize monthIncomeData with months and starting incomes of 0
    months.forEach((month) => {
        monthIncomeData[month] = { "Devin": 0, "Hannah": 0, "D&H Beck": 0 };
    });

    months.forEach((month) => {
        const monthSheet = ss.getSheetByName(month);
        if (!monthSheet) return;

        const data = monthSheet.getDataRange().getValues();

        for (let i = 0; i < data.length; i++) {
            for (let j = 0; j < data[i].length - 1; j++) {
                let cellContent = data[i][j].toString().trim().replace(/[^\w\s]/gi, ''); // Remove symbols and whitespace
                
                if (cellContent.toLowerCase() === "devin" && typeof data[i][j + 1] === 'number' && data[i][j + 1] > 0) {
                    monthIncomeData[month]["Devin"] += data[i][j + 1];
                }

                if (cellContent.toLowerCase() === "hannah" && typeof data[i][j + 1] === 'number' && data[i][j + 1] > 0) {
                    monthIncomeData[month]["Hannah"] += data[i][j + 1];
                }

                if (cellContent.toLowerCase() === "dh beck" && typeof data[i][j + 1] === 'number' && data[i][j + 1] > 0) {
                    monthIncomeData[month]["D&H Beck"] += data[i][j + 1];
                }
            }
        }
    });

    // Convert data objects to the format needed for sheet input
    let allData = [];
    for (let month in monthIncomeData) {
        const total = monthIncomeData[month]["D&H Beck"] + monthIncomeData[month]["Hannah"] + monthIncomeData[month]["Devin"]; // Calculate total
        if (monthIncomeData[month]["Hannah"] > 0 || monthIncomeData[month]["Devin"] > 0 || monthIncomeData[month]["D&H Beck"] > 0) {
            allData.push([month, monthIncomeData[month]["D&H Beck"], monthIncomeData[month]["Hannah"], monthIncomeData[month]["Devin"], total]);
        }
    }

    // If there's any data collected, set it to the "Income - Year" sheet starting from row 32.
    if (allData.length) {
        incomeSheet.getRange(32, 1, allData.length, 5).setValues(allData); // Adjusted to 5 columns
    }
}
