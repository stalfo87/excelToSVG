const ExcelJS = require('exceljs');
const fs = require('fs')
const CircularJSON = require('circular-json')

// var workbook = XLSX.readFile('./assets/aa.xlsx',{cellStyles:true});

// console.log(workbook.Sheets[workbook.SheetNames[0]])
// fs.writeFileSync('data.json', JSON.stringify(workbook.Sheets[workbook.SheetNames[0]]))
// const html = XLSX.utils.sheet_to_html(workbook.Sheets[workbook.SheetNames[0]])
// fs.writeFileSync('data.html', html)
  
const workbook = new ExcelJS.Workbook()
workbook.xlsx.readFile('./assets/aa.xlsx').then(res => {
    // res.eachSheet(function(worksheet, sheetId) {
    //     console.log(sheetId)
    //   });
     /* sheetIds for this file: 51, 71, 60, 50, 66, 67 */

    // console.log(res.getWorksheet(51))

    let sheet1 = res.getWorksheet(51)
    delete sheet1._workbook
    // console.log( Object.keys(sheet1))
    fs.writeFileSync('data.json', CircularJSON.stringify(sheet1._rows))

 })
// fs.writeFileSync('data.json', JSON.stringify(Workbook))
