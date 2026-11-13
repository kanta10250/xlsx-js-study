const XlsxPopulate = require('xlsx-populate');

// 既存の Book1.xlsx ワークブックを読み込む
XlsxPopulate.fromFileAsync("./Book1.xlsx").then(workbook => {
  workbook.sheet(0).cell('A1').value('特点表')

  workbook.sheet(0).range('B1:D1').value(
    [
      ['英語', '国語', '数学']
    ]
  );

  workbook.sheet(0).cell('A2').value(
    [
        ['a くん'],
        ['b くん'],
        ['c くん'],
        ['d くん'],
        ['e くん'],
    ]
  );
  const range = workbook.sheet(0).range("B2:D6");
  range.value((cell, ri, ci, range) => Math.floor(Math.random() * 101));

  return workbook.toFileAsync("./points.xlsx");
});