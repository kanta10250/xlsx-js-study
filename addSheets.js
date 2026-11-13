const XlsxPopulate = require('xlsx-populate');
const Workbook = require('xlsx-populate/lib/Workbook');

// 空のワークブックCreate
XlsxPopulate.fromBlankAsync().then(Workbook => {
  // Sheet5を最後
  const sheet5 = Workbook.addSheet('Sheet5');

  // Sheet2を一番目に
  const sheet2 = Workbook.addSheet('Sheet2', 1);

  // Sheet3をSheet5のまえに
  const sheet3 = Workbook.addSheet('Sheet3', 'Sheet5')

  // Sheet4をSheet5の前に
  const sheet4 = Workbook.addSheet('Sheet4', sheet5);

  return Workbook.toFileAsync("./sheetTest.xlsx");
});
