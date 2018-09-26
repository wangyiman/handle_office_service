let handleBody = (data, ws1, nowRow, index) => {
  let computedType = ['customComputed', 'decoComputed', 'eleComputed', 'plateComputed']
  let dataLength = data[index][computedType[index]].length;
  let newRowValue;
  let row;  
  if(dataLength < 2) {
    for(var i = 0; i < 2; i++) {
      if(i >= dataLength) {
        newRowValue = [];
      } else {
        newRowValue = [
          data[index][computedType[index]][i].order, 
          data[index][computedType[index]][i].name,
          data[index][computedType[index]][i].brand, 
          data[index][computedType[index]][i].sku,
          data[index][computedType[index]][i].standlizeSize.x,
          data[index][computedType[index]][i].standlizeSize.y,
          data[index][computedType[index]][i].standlizeSize.z, 
          data[index][computedType[index]][i].material.body,
          data[index][computedType[index]][i].material.door, 
          data[index][computedType[index]][i].unit,
          data[index][computedType[index]][i].quantity,
          data[index][computedType[index]][i].number,
          data[index][computedType[index]][i].unitPrice,
          data[index][computedType[index]][i].customPrice,
          data[index][computedType[index]][i].discount,
          data[index][computedType[index]][i].discountPrice,
          data[index][computedType[index]][i].instructions
        ];
      }
      ws1.spliceRows(nowRow + i, 0, newRowValue);
    }
    dataLength = 2;
  } else if(dataLength >= 2) {
    for(var i = 0; i < dataLength; i++) {
      newRowValue = [
        data[index][computedType[index]][i].order, 
        data[index][computedType[index]][i].name,
        data[index][computedType[index]][i].brand, 
        data[index][computedType[index]][i].sku,
        data[index][computedType[index]][i].standlizeSize.x,
        data[index][computedType[index]][i].standlizeSize.y,
        data[index][computedType[index]][i].standlizeSize.z, 
        data[index][computedType[index]][i].material.body,
        data[index][computedType[index]][i].material.door, 
        data[index][computedType[index]][i].unit,
        data[index][computedType[index]][i].quantity,
        data[index][computedType[index]][i].number,
        data[index][computedType[index]][i].unitPrice,
        data[index][computedType[index]][i].customPrice,
        data[index][computedType[index]][i].discount,
        data[index][computedType[index]][i].discountPrice,
        data[index][computedType[index]][i].instructions
      ];
      ws1.spliceRows(nowRow + i, 0, newRowValue);
    }
  }
  nowRow = nowRow + dataLength;
  ws1.mergeCells(`A${nowRow}:Q${nowRow}`);
  ws1.getCell(`A${nowRow}`).value = '小计';
  ws1.getCell(`A${nowRow}`).fill = {type: "pattern", pattern:"solid", fgColor: {argb:"FFD9D9D9"}};   
  nowRow++;
  
  return {
    ws1: ws1,
    nowRow: nowRow
  }
}

module.exports = handleBody;