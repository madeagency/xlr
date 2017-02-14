const { sheetFront, sheetBack, columnDefinition } = require('./constants')
const { addCell, convertToBit } = require('./helpers')

function buildRow (columns, data, rowIndex) { 
  const rows = [`<x:row r="${rowIndex}" spans="1:${columns.length}">`]

  for (var i = 0; i < data.length; i++) {
    const cellIndex = i + 1

    const isObject = typeof data[i] === 'object'
    const cellData = isObject ? data[i].value : data[i]
    const styleIndex = isObject ? data[i].style : null
    const cellType = columns[i].type

    const options = { 
      cellIndex: cellIndex + rowIndex,
      value: cellData,
      styleIndex
    }

    switch (cellType) {
    case 'number':
    case 'date':
      rows.push(addCell(Object.assign({ type: 'n' }, options)))
      break
    case 'bool':
      rows.push(addCell(Object.assign({ value: convertToBit(cellData), type: 'b' }, options)))
      break
    default:
      rows.push(addCell(options))
    }
  }

  rows.push('</x:row>')
  return rows.join('')
}

function Sheet (config) {
  const { columns, rows } = config

  // build column elements
  const columnDefinitions = columns.map((column, i) => columnDefinition(column.width, i + 1)).join('')  
  // build xml rows
  const xmlRows = rows.map((rowData, i) => buildRow(columns, rowData, i + 1)).join('')
  return `${sheetFront}<cols>${columnDefinitions}</cols><x:sheetData>${xmlRows}</x:sheetData>${sheetBack}`
}

module.exports = Sheet

