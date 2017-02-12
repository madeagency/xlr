# XLR

A simple node.js module for exporting a data set to an Excel xlsx file.

## Using XLR

If generating multiple sheets, the configs object should be an array of multiple config objects.
If generating a single sheet, the config object should be a single object.

Example of a single worksheet
```javascript
{
  name: 'name.xlsx',
  cols: [
    {
      type: 'string',
      width: 10,
    }
  ],
  rows: [].
  stylesXmlFile: 'styles.xml'
}
```

- name: Specify worksheet name
- cols: `Array` of column definitions
  - type: string / date / bool / number
  - width: (optional) total characters in cell (blame excel)
- rows: `Array` of data to be exported. Data needs to be a 2D-`Array`. Each row should be the same length as cols.
- stylesXmlFile: Absolute path to Excel `styles.xml` file. An easy way to get a `styles.xml` file is to unzip an existing xlsx file which has the desired styles and copy the `styles.xml` file.

## Example usage with Express

```javascript
const express = require('express');
const xlr = require('xlr');
const app = express();

app.get('/', function (req, res) {
  const conf = {
    stylesXmlFile: __dirname + '/styles.xml',
    name: 'report-is-a-thing',
    cols: [
      {
        caption: '',
        type: 'string',
        width: 15
      },
      {
        caption: '',
        type: 'string',
        width: 15
      },
      {
        caption: '',
        type: 'string',
        width: 30
      }
    ],
    rows: [
      [{ value: 'First Name', style: 1 }, { value: 'Last Name', style: 1 }, { value: 'Email', style: 1 }],
      ['Bruce', 'Wayne', 'info@bat.man'],
      ['Clark', 'Kent', 'info@super.man'],
      ['Test', '', 'test@email.com'],
      ['Peter', 'Parker', 'info@spider.man']
    ]
  };

  const result = xlr(conf);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats');
  res.setHeader('Content-Disposition', 'attachment; filename=Report.xlsx');
  res.send(new Buffer(result, 'binary'));
});

app.listen(3001, function () {
  console.log('listening on port 3001');
});
```
