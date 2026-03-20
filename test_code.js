function onOpen() {
    SpreadsheetApp.getUi()
    .createMenu('Custom')
    .addItem("filler", "columnfiller")
    .addSeparator()
    .addItem("column 1 filler" , "column1filler")
    .addToUi();
};

const columnfiller = () => {
     const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const startRow = 2;
  const numRows = 50;
  const column = 3; // column C

  const values = Array.from({ length: numRows }, () => [
    Math.floor(Math.random() * 100) + 1
  ]);

  sheet.getRange(startRow, column, numRows, 1).setValues(values);
}
const column1filler = () =>{
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = 1;
  const numRows = 50;
  const column = 1; // column C

  const values = Array.from({ length: numRows }, () => [
    Math.floor(Math.random() * 100) + 1
  ]);

  sheet.getRange(startRow, column, numRows, 1).setValues(values);
}