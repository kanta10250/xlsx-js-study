const XlsxPopulate = require('xlsx-populate');

XlsxPopulate.fromFileAsync("./Book1.xlsx").then(Workbook => {
  //Sheet1のA1セル取得
  const value = Workbook.sheet('Sheet1').cell('A1').value()

  console.log(value)
})