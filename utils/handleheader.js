const Excel = require('exceljs');
const fs = require('fs');

//handle header
var handleHeader = (ws, nowRow, index) => {
  var title = ['橱柜部分','板件部分','水电器五金配件表','装饰线条表'];
  var letter = ['A', 'B', 'C', 'D', 'E','F','G','H','I','J','K','L','M','N','O','P','Q'];
  var headerName = ['序号','名称','品牌','编号',{
    '标准尺寸': ['宽','深','高']
  },
  {
    '材料': ['柜身','柜门']
  },'单位','数量','件数','单价','定制费','折扣','折后价','说明'];

  //insert rows.
  // title
  var nowTitle = title[index];
  var titleRow = nowRow;
  ws.spliceRows(titleRow, 0, [nowTitle]);;
  ws.mergeCells(`A${titleRow}: Q${titleRow}`);
  ws.getCell(`A${titleRow}`).value = nowTitle;
  ws.getCell(`A${titleRow}`).fill = {type: "pattern", pattern:"solid", fgColor: {argb:"FF00B0F0"}};
  nowRow++;

  // headerName iterater.
  for(var i = 0; i < 14;i++) {
    //0,1,2,3
    let nowI = i;
    if(i !== 4 && i !== 5) {
      if(i > 5) {
        nowI = i + 3;
      }
      ws.mergeCells(`${letter[nowI]}${nowRow}: ${letter[nowI]}${nowRow + 1}`);
      ws.getCell(`${letter[nowI]}${nowRow}`).value = headerName[i];
      ws.getCell(`${letter[nowI]}${nowRow}`).alignment = { vertical: 'middle', horizontal: 'center' };  
      ws.getCell(`${letter[nowI]}${nowRow}`).fill = {type: "pattern", pattern:"solid", fgColor: {argb:"FFD9D9D9"}};
    } else if(i === 4) {
      let row4 = nowRow;
      let nowI4 = i;
      ws.mergeCells(`${letter[nowI4]}${row4}: ${letter[nowI4+2]}${row4}`);
      ws.getCell(`${letter[nowI4]}${row4}`).value = '标准尺寸';
      ws.getCell(`${letter[nowI4]}${row4}`).alignment = { vertical: 'middle', horizontal: 'center' };  
      ws.getCell(`${letter[nowI4]}${row4}`).fill = {type: "pattern", pattern:"solid", fgColor: {argb:"FFD9D9D9"}};

      ws.getCell(`${letter[nowI4]}${++row4}`).value = '宽';
      ws.getCell(`${letter[nowI4]}${row4}`).alignment = { vertical: 'middle', horizontal: 'center' };  
      ws.getCell(`${letter[nowI4]}${row4}`).fill = {type: "pattern", pattern:"solid", fgColor: {argb:"FFD9D9D9"}};
      
      ws.getCell(`${letter[++nowI4]}${row4}`).value = '深';
      ws.getCell(`${letter[nowI4]}${row4}`).alignment = { vertical: 'middle', horizontal: 'center' };  
      ws.getCell(`${letter[nowI4]}${row4}`).fill = {type: "pattern", pattern:"solid", fgColor: {argb:"FFD9D9D9"}};
      ws.getCell(`${letter[++nowI4]}${row4}`).value = '高';   
      ws.getCell(`${letter[nowI4]}${row4}`).alignment = { vertical: 'middle', horizontal: 'center' };  
      ws.getCell(`${letter[nowI4]}${row4}`).fill = {type: "pattern", pattern:"solid", fgColor: {argb:"FFD9D9D9"}}; 
    } else if(i === 5 && index === 0) {
      let row7 = nowRow;
      let nowI7 = i + 2;        
      ws.mergeCells(`${letter[nowI7]}${row7}: ${letter[nowI7+1]}${row7}`);
      ws.getCell(`${letter[nowI7]}${row7}`).value = '材料';
      ws.getCell(`${letter[nowI7]}${row7}`).alignment = { vertical: 'middle', horizontal: 'center' };  
      ws.getCell(`${letter[nowI7]}${row7}`).fill = {type: "pattern", pattern:"solid", fgColor: {argb:"FFD9D9D9"}};
      
      ws.getCell(`${letter[nowI7]}${++row7}`).value = '柜身';
      ws.getCell(`${letter[nowI7]}${row7}`).alignment = { vertical: 'middle', horizontal: 'center' };  
      ws.getCell(`${letter[nowI7]}${row7}`).fill = {type: "pattern", pattern:"solid", fgColor: {argb:"FFD9D9D9"}};
      ws.getCell(`${letter[++nowI7]}${row7}`).value = '柜门';
      ws.getCell(`${letter[nowI7]}${row7}`).alignment = { vertical: 'middle', horizontal: 'center' };  
      ws.getCell(`${letter[nowI7]}${row7}`).fill = {type: "pattern", pattern:"solid", fgColor: {argb:"FFD9D9D9"}}; 
    } else if(i === 5) {
      let row5 = nowRow;
      let nowI5 = i + 2;  
      ws.mergeCells(`${letter[nowI5]}${row5}: ${letter[nowI5+1]}${row5 + 1}`);
      ws.getCell(`${letter[nowI5]}${nowRow}`).value = '材料';
      ws.getCell(`${letter[nowI5]}${nowRow}`).alignment = { vertical: 'middle', horizontal: 'center' };  
      ws.getCell(`${letter[nowI5]}${nowRow}`).fill = {type: "pattern", pattern:"solid", fgColor: {argb:"FFD9D9D9"}};
    }
  }
  nowRow = nowRow + 2;
  
  return {
    ws1: ws,
    nowRow: nowRow
  }
};

module.exports = handleHeader;
