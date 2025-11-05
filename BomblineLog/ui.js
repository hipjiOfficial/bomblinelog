function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Tools')
    .addItem('Colorize Log', 'colorCellsBasedOnReferenceList')
    .addItem('Sort', 'sortMultipleItemBlocksPreserveFormulas')
    .addItem('Count Lost Media', 'daysLost')
    .addToUi();
}
/*
This makes a little button at the top that gives you a menu where you can run each function. 
