function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Custom-Menu')
  .addItem('Search Client', 'searchClientVisitDataBlockFormat')
  .addItem('Generate Consumption', 'generateProductConsumptionReport')
  .addToUi();
}
