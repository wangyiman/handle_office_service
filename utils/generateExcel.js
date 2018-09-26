const Excel = require('exceljs');
const fs = require('fs');
const handleHeader = require('./handleheader');
const handlBody = require('./handlebody');

var  generateExcel = async (data) => {
  var filename='excelTemplate/excelTemplate2.xlsx';
  var workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filename)
  return insertWorkBook(workbook, data);
}

async function insertWorkBook(workbook, postdata) {

    var ws1 = workbook.getWorksheet("橱柜报价清单");
    let data = postdata.bomList;
    let userInfo = postdata.userInfo;

    //userInfo
    let userInfoRow1 = ws1.getRow(3);
    userInfoRow1.getCell('C').value = userInfo.designId;
    userInfoRow1.getCell('P').value = userInfo.designUser;   
    let userInfoRow2 = ws1.getRow(4);
    userInfoRow2.getCell('P').value = userInfo.customerAddress;
    //data
    let index;
    let nowRow;
    //first type
    //橱柜部分
    //header
    nowRow = 6;
    index = 0;
    let headerResult = handleHeader(ws1, nowRow, index);
    ws1 = headerResult.ws1;
    nowRow = headerResult.nowRow;
    //body
    let bodyResult = handlBody(data, ws1, nowRow, index);
    ws1 = bodyResult.ws1;
    nowRow = bodyResult.nowRow;

    //
    index = 1;  
    headerResult = handleHeader(ws1, nowRow, index);
    ws1 = headerResult.ws1;
    nowRow = headerResult.nowRow;
    //body
    bodyResult = handlBody(data, ws1, nowRow, index);
    ws1 = bodyResult.ws1;
    nowRow = bodyResult.nowRow;

    //
    index = 2;  
    headerResult = handleHeader(ws1, nowRow, index);
    ws1 = headerResult.ws1;
    nowRow = headerResult.nowRow;
    //body
    bodyResult = handlBody(data, ws1, nowRow, index);
    ws1 = bodyResult.ws1;
    nowRow = bodyResult.nowRow;

    //
    index = 3;  
    headerResult = handleHeader(ws1, nowRow, index);
    ws1 = headerResult.ws1;
    nowRow = headerResult.nowRow;
    //body
    bodyResult = handlBody(data, ws1, nowRow, index);
    ws1 = bodyResult.ws1;
    nowRow = bodyResult.nowRow;

    ws1.eachRow(function(row, rowNumber) {
      row.eachCell(function(cell, colNumber) {
        cell.border = {
          top: {style:'thin', color: {argb:'00000000'}},
          left: {style:'thin', color: {argb:'00000000'}},
          bottom: {style:'thin', color: {argb:'00000000'}},
          right: {style:'thin', color: {argb:'00000000'}}
        };
      });
    });

    var date = new Date();
    var formatDate = `报价清单${date.getFullYear()}${date.getMonth() + 1}${date.getDate()}${date.getHours()}${date.getMinutes()}${date.getSeconds()}`;

    var filename=`/${formatDate}.xlsx`;//生成的文件名
    await workbook.xlsx.writeFile(`./tempExcel${filename}`);
    return filename;
}

 module.exports = generateExcel;
