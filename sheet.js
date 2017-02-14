const { sheetFront, sheetBack, columnDefinition } = require('./constants')
const { addMergeCell, addCell, convertToBit } = require('./helpers')

function buildRow (columns, data, rowIndex) { 
  const rows = [`<x:row r="${rowIndex}" spans="1:${columns.length}">`]

  for (var i = 0; i < data.length; i++) {
    const cellIndex = i + 1

    const isObject = typeof data[i] === 'object'
    const cellData = isObject ? data[i].value : data[i]
    const styleIndex = isObject ? data[i].style : null
    const cellType = columns[i].type

    const options = { 
      cellIndex,
      value: cellData,
      styleIndex,
      rowIndex
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

function buildMergeCells (merge) {
  if (merge) {
    const mergeCells = merge.map((mergeCell) => addMergeCell(mergeCell)).join('')
    return `<mergeCells count="${merge.length}">${mergeCells}</mergeCells>`
  } else {
    return ''
  }
}

function Sheet ({ columns, rows, merge }) {
  if (!columns) throw 'No columns provided'
  if (!rows) throw 'No row data provided'

  // build column elements
  const columnDefinitions = columns.map((column, i) => columnDefinition(column.width, i + 1)).join('')  
  // build xml rows
  const xmlRows = rows.map((rowData, i) => buildRow(columns, rowData, i + 1)).join('')
  return `${sheetFront}<cols>${columnDefinitions}</cols><x:sheetData>${xmlRows}</x:sheetData>${buildMergeCells(merge)}${sheetBack}`
}

module.exports = Sheet

