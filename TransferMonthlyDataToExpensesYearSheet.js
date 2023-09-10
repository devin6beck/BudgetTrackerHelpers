function TransferMonthlyDataToExpensesYearSheet() {
    // Get the active spreadsheet instance.
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get the sheet named "Expenses - Year".
    const expensesSheet = ss.getSheetByName("Expenses - Year");

    // If the "Expenses - Year" sheet does not exist, throw an error.
    if (!expensesSheet) {
        throw new Error("Expenses - Year sheet not found");
    }

    const lastRow = expensesSheet.getLastRow();
    if (lastRow >= 31) {
        expensesSheet.getRange(31, 1, lastRow - 30, expensesSheet.getLastColumn()).clearContent();
    }


    // Append header row to the sheet at row 31.
    expensesSheet.getRange(31, 1, 1, 3).setValues([['Date', 'Amount', 'Category']]);

    // Initialize an empty array to collect data from all months.
    let allData = [];

    // Loop through each month to fetch the relevant data.
    months.forEach((month, index) => {
        const monthSheet = ss.getSheetByName(month);

        // If the month sheet doesn't exist, skip the current iteration.
        if (!monthSheet) return;

        const { activeHeaders, lastRow } = findHeadersAndLastRow(monthSheet, ['Date', 'Amount', 'Category']);
        if (!activeHeaders) return;

        // Find the header row which contains the "Date" column.
        const headerRow = monthSheet.getDataRange().getValues().findIndex(row => row.includes("Date"));

        // Get the indices of 'Amount' and 'Category' columns.
        const amountIndex = activeHeaders.indexOf('Amount');
        const categoryIndex = activeHeaders.indexOf('Category');

        // Fetch data below the header row.
        const data = monthSheet.getRange(headerRow + 2, 1, lastRow - headerRow, activeHeaders.length).getValues();

        // Loop through the fetched data to filter and transform.
        data.forEach(row => {
            const category = row[categoryIndex];
            let amount = row[amountIndex];

            // If the category includes "income", skip the current iteration.
            if (!category || !amount || category.toLowerCase().includes("income") || category === "Total" || category === "") return;

            // If the month index is greater than or equal to 6, negate the amount.
            if (index >= 6) {
                amount = -amount;
            }

            // Push the transformed data to the allData array.
            allData.push([month, amount, category]);
        });

        // Special handling for the first six months.
        if (months.indexOf(month) <= 5) {
            const monthlyBillsRow = monthSheet.getDataRange().getValues().findIndex(row => row.includes("Monthy Bills"));
            let currentRow = monthlyBillsRow + 2;

            // Loop to fetch data for monthly bills.
            while (true) {
                const category = monthSheet.getRange(currentRow, 6).getValue();
                const amount = monthSheet.getRange(currentRow, 8).getValue();

                if (!category || !amount) return;

                // Check various conditions before proceeding. If any of them is met, exit the loop.
                if (category === "Total") {
                    break;
                }

                // Push the data to the allData array.
                allData.push([month, amount, category]);
                currentRow++;
            }
        }
    });

    // If there's any data collected, set it to the "Expenses - Year" sheet starting from row 32.
    if (allData.length) {
        expensesSheet.getRange(32, 1, allData.length, 3).setValues(allData);
    }
}
