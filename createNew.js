const XlsxPopulate = require('xlsx-populate');
const Workbook = require('xlsx-populate/lib/Workbook');

// 空のワークブック
XlsxPopulate.fromBlankAsync().then(Workbook => {
  // Sheet1のA1セルに文章入れる
  Workbook.sheet('Sheet1').cell('A1').value('新しく作った Excel');
  Workbook.sheet('Sheet1').column('A').width(25)
  // ファイル書き出し
  return Workbook.toFileAsync("./Book1.xlsx");
});