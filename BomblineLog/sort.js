function sortMultipleItemBlocksPreserveFormulas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const columnPairs = [1, 4, 7, 10, 13, 16, ]; // A/B, D/E, G/H, J/K
  const startRow = 3;

  columnPairs.forEach(colStart => {
    const colEnd = colStart + 1;
    const lastRow = sheet.getLastRow();
    const numRows = lastRow - startRow + 1;

    const blockRange = sheet.getRange(startRow, colStart, numRows, 2);
    const blockValues = blockRange.getValues();
    const formulaColumnRange = sheet.getRange(startRow, colEnd, numRows, 1);
    const formulas = formulaColumnRange.getFormulas();

    // Attach formulas to values for later reattachment
    const rowsWithMetadata = blockValues.map((row, i) => {
      const countMatch = row[1].toString().match(/\d+/);
      const count = countMatch ? parseInt(countMatch[0]) : 0;
      return {
        item: row[0],
        text: row[1],
        count: count,
        formula: formulas[i][0],
      };
    });

    // Sort rows descending by count
    rowsWithMetadata.sort((a, b) => b.count - a.count);

    // Write back values and formulas
    rowsWithMetadata.forEach((row, i) => {
      sheet.getRange(startRow + i, colStart).setValue(row.item);
      const targetCell = sheet.getRange(startRow + i, colEnd);
      if (row.formula) {
        targetCell.setFormula(row.formula);
      } else {
        targetCell.setValue(row.text);
      }
    });
  });
}