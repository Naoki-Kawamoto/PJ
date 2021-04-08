//Excelへ書き込むコード
const XLSX = require("xlsx");
const Utils = XLSX.utils; // XLSX.utilsのalias
// Workbookの読み込み
const book = XLSX.readFile("患者用　健康診断結果用紙.xlsx");
// Sheet1読み込み
const sheet1 = book.Sheets["結果用紙"];

// セル更新
sheet1["K14"] = {"これはフロントエンドからgetElementsで持ってくる" };
// 範囲の更新を忘れずに (30分ハマった)
sheet1["!ref"] = ":";
book.Sheets["Sheet2"] = sheet2;
// ファイルを書き込み
XLSX.writeFile(book, "患者用　健康診断結果用紙.xlsx");




function onButtonclick1(){
    //ここからExcelへ書き込んでいく
    var your-name = document.getElementsByName("your-name");


}