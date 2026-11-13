const XlsxPopulate = require('xlsx-populate');

// 既存の 得点表の points.xlsx ワークブックを読み込む
XlsxPopulate.fromFileAsync("./points.xlsx").then(workbook => {

  workbook.find("特典表", '点数票');

  return workbook.toFileAsync("./points2.xlsx");
});