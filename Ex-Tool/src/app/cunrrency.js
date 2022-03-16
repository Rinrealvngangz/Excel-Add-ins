/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
async function currency(event){
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
    
        // Create the headers and format them to stand out.
        let headers = [
          ["Product", "Quantity", "Unit Price", "Totals"]
        ];
        let headerRange = sheet.getRange("B2:E2");
        headerRange.values = headers;
        headerRange.format.fill.color = "#4472C4";
        headerRange.format.font.color = "white";
    
        // Create the product data rows.
        let productData = [
          ["Almonds", 6, 7.5],
          ["Coffee", 20, 34.5],
          ["Chocolate", 10, 9.56],
        ];
        let dataRange = sheet.getRange("B3:D5");
        dataRange.values = productData;
    
        // Create the formulas to total the amounts sold.
        let totalFormulas = [
          ["=C3 * D3"],
          ["=C4 * D4"],
          ["=C5 * D5"],
          ["=SUM(E3:E5)"]
        ];
        let totalRange = sheet.getRange("E3:E6");
        totalRange.formulas = totalFormulas;
        totalRange.format.font.bold = true;
    
        // Display the totals as US dollar amounts.
        totalRange.numberFormat = [["$0.00"]];
    
        await context.sync();
    });
    event.completed();
}
