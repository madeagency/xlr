function getColumnLetter (column) {
  if (!column || column <= 0) {
    throw 'column must be more than 0'
  }

  const array = []
  while (column > 0) {
    var remainder = column % 26
    column /= 26
    column = Math.floor(column)
    if (remainder === 0) {
      remainder = 26
      column--
    }
    array.push(64 + remainder)
  }

  return String.fromCharCode.apply(null, array.reverse())
}

exports.convertToBit = function (value) {
  if (typeof value === 'boolean') {
    return value ? 1 : 0
  } else if (typeof value === 'string') {
    return value.toLowerCase() === 'true' ? 1 : 0
  } else {
    return 0
  }
}

exports.addCell = function (options) {
  const { rowIndex, type } = options
  const cellReference = `${getColumnLetter(options.cellIndex)}${rowIndex}`
  const styleIndex = options.styleIndex || 0
  const value = options.value || ''

  if (type) {
    return `<x:c r="${cellReference}" s="${styleIndex}" t="${type}"><x:v>${value}</x:v></x:c>`
  } else {
    const encodedValue = value.replace(/&/g, '&amp;').replace(/'/g, '&apos;').replace(/>/g, '&gt;').replace(/</g, '&lt;')
    return `<x:c r="${cellReference}" s="${styleIndex}" t="inlineStr"><x:is><x:t>${encodedValue}</x:t></x:is></x:c>`
  }
}

exports.addMergeCell = function ({ row, fromColumn, toColumn }) {
  const fromCellReference = `${getColumnLetter(fromColumn)}${row}`
  const toCellReference = `${getColumnLetter(toColumn)}${row}`
  return `<mergeCell ref="${fromCellReference}:${toCellReference}" />`
}

