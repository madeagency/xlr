# XLR

A simple node module for exporting a data set to an Excel xlsx file.

## Using XLR

Example of a config object to generate a Excel worksheet
```javascript
{
  name: 'name.xlsx',
  columns: [
    {
      type: 'string',
      width: 10,
    }
  ],
  rows: [],
  merge: []
  stylesXmlFile: 'styles.xml'
}
```

- name: Specify worksheet name
- columns: `Array` of column definitions
  - type: string / date / bool / number
  - width: (optional) total characters in cell
- rows: `Array` of data to be exported. Data needs to be a 2D-`Array` and should be the same length as the columns array.
- merge: `Array` of merge cell objects.
  - row: row index (starting at 1)
  - fromColumn: start the merge from column index (starting at 1)
  - toColumn: end the merge at column index (starting at 1)
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
        width: 25
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
      ['This cell is going to be merged', '', 'test@email.com'],
      ['Peter', 'Parker', 'info@spider.man']
    ],
     merge: [
      { row: 4, fromColumn: 1, toColumn: 2 } 
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

License
-------

XLR is © 2016 MADE Code PTY Ltd.
It is free software, and may be redistributed under the terms specified in the [LICENSE] file.

[LICENSE]: LICENSE

Maintained by
----------------

[![madeagency](https://www.made.co.za/logo.png)](https://www.made.co.za?utm_source=github)

XLR was created and is maintained MADE Agency PTY Ltd.
The names and logos for MADE Code are trademarks of MADE Code PTY Ltd.

We love open source software. See our [Github Profile](https://github.com/madeagency) for more.

We're always looking for talented people who love programming. [Get in touch] with us.

[Get in touch]: https://www.made.co.za?utm_source=github

