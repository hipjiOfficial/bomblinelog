function daysLost() { 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const checkRange = sheet.getRange("A3:A4000");
  const targetCell = sheet.getRange("I31");
  const targetColor = "#ea4335";
  const backgrounds = checkRange.getBackgrounds();
  
  let redCount = 0;
  for (let row of backgrounds) {
    for (let color of row) {
      if (color && color.toLowerCase() === targetColor.toLowerCase()) {
        redCount++;
      }
    }
  }

  targetCell.setValue(redCount);
}
