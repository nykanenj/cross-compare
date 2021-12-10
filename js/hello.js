const hello = () => {
  SpreadsheetApp.getActive()
    .getActiveSheet()
    .getRange("A1")
    .setValue("Hello World!");
};
