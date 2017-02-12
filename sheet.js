const constants = require('./constants');    
const helpers = require('./helpers');

function buildRow (cols, data, rowIndex) { 
  const rows = [`<x:row r="${rowIndex}" spans="1:${cols.length}">`];

  for (var i = 0; i < data.length; i++) {
    const cellIndex = i + 1;

    const isObject = typeof data[i] === 'object';
    const cellData = isObject ? data[i].value : data[i];
    const styleIndex = isObject ? data[i].style : null;
    const cellType = cols[i].type;

    switch (cellType) {
    case 'number':
    case 'date':
      rows.push(helpers.addCell({ 
        cellIndex: cellIndex + rowIndex,
        value: cellData,
        styleIndex,
        type: 'n' }));
      break;
    case 'bool':
      rows.push(helpers.addCell({ 
        cellIndex: cellIndex + rowIndex,
        cellData: helpers.convertToBit(cellData),
        value: styleIndex,
        type: 'b'
      }));
      break;
    default:
      rows.push(helpers.addCell({ 
        cellIndex: cellIndex + rowIndex,
        value: cellData,
        styleIndex
      }));
    }
  }

  rows.push('</x:row>');
  return rows.join('');
}

function Sheet (config) {
  const cols = config.cols;
  const data = config.rows;

  // build column elements
  const colDefs = cols.map((col, i) => {
    const index = i + 1;
    return `<col customWidth="1" width="${col.width}" max="${index}" min="${index}"/>`;
  }).join('');
  
  // build xml rows
  const rows = data.map((rowData, i) => buildRow(cols, rowData, i + 1)).join('');

  return `${constants.sheetFront}<cols>${colDefs}</cols><x:sheetData>${rows}</x:sheetData>${constants.sheetBack}`;
}

module.exports = Sheet;

