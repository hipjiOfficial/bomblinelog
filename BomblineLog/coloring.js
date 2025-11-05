function colorCellsBasedOnReferenceList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const logSheet = ss.getSheetByName('Log');
  const referenceSheet = ss.getSheetByName('Statistics');

  const logRange = logSheet.getRange('B1:E4000');
  const logValues = logRange.getValues();

  const referenceRange = referenceSheet.getRange('A3:Q4000');
  const referenceValues = referenceRange.getValues();
  const referenceColors = referenceRange.getBackgrounds();

  // creates map
  const colorMap = {};
  for (let row = 0; row < referenceValues.length; row++) {
    for (let col = 0; col < referenceValues[0].length; col++) {
      const key = referenceValues[row][col].toString().trim().toLowerCase();
      if (key && !(key in colorMap)) {
        colorMap[key] = referenceColors[row][col];
      }
    }
  }

  const newColors = logValues.map(row =>
    row.map(cell => {
      const lookup = cell.toString().trim().toLowerCase();
      return colorMap[lookup] || '#d9d9d9'; 
    })
  );

  logRange.setBackgrounds(newColors);
}

// there are 85 player cards as of 7/20/25
