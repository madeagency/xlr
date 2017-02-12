function getColumnLetter (col) {
  if (!col || col <= 0) {
    throw 'col must be more than 0';
  }

  const array = [];
  while (col > 0) {
    var remainder = col % 26;
    col /= 26;
    col = Math.floor(col);
    if (remainder === 0) {
      remainder = 26;
      col--;
    }
    array.push(64 + remainder);
  }

  return String.fromCharCode.apply(null, array.reverse());
}

exports.convertToBit = function (value) {
  if (typeof value === 'boolean') {
    return value ? '1' : '0';
  } else if (typeof value === 'string') {
    return value.toLowerCase() === 'true' ? '1' : '0';
  } else {
    return '0';
  }
};

exports.addCell = function (options) {
  const cellRef = getColumnLetter(options.cellIndex);
  const styleIndex = options.styleIndex || 0;
  const value = options.value || '';
  const type = options.type;

  if (type) {
    return `<x:c r="${cellRef}" s="${styleIndex}" t="${type}"><x:v>${value}</x:v></x:c>`;
  } else {
    const encodedValue = value.replace(/&/g, '&amp;').replace(/'/g, '&apos;').replace(/>/g, '&gt;').replace(/</g, '&lt;');
    return `<x:c r="${cellRef}" s="${styleIndex}" t="inlineStr"><x:is><x:t>${encodedValue}</x:t></x:is></x:c>`;
  }
};

exports.startTag = function (obj, tagName, closed) {
  var result = '<' + tagName;
  for (var p in obj) {
    result += ' ' + p + '=' + obj[p];
  }

  if (!closed) {
    result += '>';
  } else {
    result += '/>';
  }

  return result;
};

exports.endTag = function (tagName) {
  return `</${tagName}>`;
};
