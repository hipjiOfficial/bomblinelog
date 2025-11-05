function updateCardAndSkinProbabilities() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Log");
  const statsSheet = ss.getSheetByName("Statistics");

  const logData = logSheet.getRange(3, 2, logSheet.getLastRow() - 2, 4).getValues(); // B–E (3 cards + 1 skin)

  const cardCounts = {};
  const skinCounts = {};
  let totalCardSlots = 0;
  let totalSkinSlots = 0;

  logData.forEach(row => {
    // first 3 columns = cards
    for (let i = 0; i < 3; i++) {
      const card = String(row[i]).trim().toLowerCase();
      if (card) {
        totalCardSlots++;
        cardCounts[card] = (cardCounts[card] || 0) + 1;
      }
    }

    // last column = skin
    const skin = String(row[3]).trim().toLowerCase();
    if (skin) {
      totalSkinSlots++;
      skinCounts[skin] = (skinCounts[skin] || 0) + 1;
    }
  });

  const lastRow = statsSheet.getLastRow();
  const cardCols = getPopulatedThirdColumns(statsSheet, 1);   // a, d, d, ...
  const skinCols = getPopulatedThirdColumns(statsSheet, 13);  // m, p, s, ...

  //update cards
  cardCols.forEach(col => {
    const cardRange = statsSheet.getRange(3, col, lastRow - 2, 1);
    const cardNames = cardRange.getValues();
    

    const cardOutput = cardNames.map(([cardName]) => {
      const name = String(cardName).trim().toLowerCase();
      const appearances = cardCounts[name] || 0;
      const probability = totalCardSlots ? appearances / totalCardSlots : 0;
      return [probability];
    });

    statsSheet.getRange(3, col + 2, cardOutput.length, 1).setValues(cardOutput); // probability goes in col + 2
    statsSheet.getRange(3, col + 2, cardOutput.length, 1).setNumberFormat("0.00%");
  });

  // update skins
  skinCols.forEach(col => {
    const skinRange = statsSheet.getRange(3, col, lastRow - 2, 1);
    const skinNames = skinRange.getValues();

    const skinOutput = skinNames.map(([skinName]) => {
      const name = String(skinName).trim().toLowerCase();
      const appearances = skinCounts[name] || 0;
      const probability = totalSkinSlots ? appearances / totalSkinSlots : 0;
      return [probability];
    });

    statsSheet.getRange(3, col + 2, skinOutput.length, 1).setValues(skinOutput); // probability goes in col + 2
    statsSheet.getRange(3, col + 2, skinOutput.length, 1).setNumberFormat("0.00%");
  });

  SpreadsheetApp.flush();
}

function getPopulatedThirdColumns(sheet, startCol) {
  const lastRow = sheet.getLastRow();
  const populatedCols = [];

  for (let col = startCol; col <= sheet.getLastColumn(); col += 3) {
    const range = sheet.getRange(3, col, lastRow - 2, 1).getValues();
    const hasData = range.some(row => String(row[0]).trim() !== '');
    if (hasData) {
      populatedCols.push(col);
    } else {
      break; // stop when the first empty column is found
    }
  }

  return populatedCols;
}
