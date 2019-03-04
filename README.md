# XLR
[![Code Climate](https://codeclimate.com/repos/58a2efd0c075372cbb00271c/badges/51daec96fa87c65918c5/gpa.svg)](https://codeclimate.com/repos/58a2efd0c075372cbb00271c/feed)
[![FOSSA Status](https://app.fossa.io/api/projects/git%2Bhttps%3A%2F%2Fgithub.com%2Fmadeagency%2Fxlr.svg?type=shield)](https://app.fossa.io/projects/git%2Bhttps%3A%2F%2Fgithub.com%2Fmadeagency%2Fxlr?ref=badge_shield)

[![https://nodei.co/npm/xlr.png?downloads=true&downloadRank=true&stars=true](https://nodei.co/npm/xlr.png?downloads=true&downloadRank=true&stars=true)](https://www.npmjs.com/package/xlr)

A simple node module for exporting a data set to an Excel xlsx file.

Heavily inspired by [Node-Excel-Export](https://github.com/functionscope/Node-Excel-Export)

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
      { type: 'string', width: 15 },
      { type: 'string', width: 25 },
      { type: 'string', width: 30 },
      { type: 'number', width: 15 }
    ],
    rows: [
      [{ value: 'First Name', style: 1 }, 
       { value: 'Last Name', style: 1 }, 
       { value: 'Email', style: 1 }, 
       { value: 'age', style: 1, type: 'string' }],
      ['Bruce', 'Wayne', 'info@bat.man', 25],
      ['Clark', 'Kent', 'info@super.man', 35],
      ['This cell is going to be merged', '', 'test@email.com', 45],
      ['Peter', 'Parker', 'info@spider.man', 55]
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

XLR is © 2017 MADE Code PTY Ltd.
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



## License
[![FOSSA Status](https://app.fossa.io/api/projects/git%2Bhttps%3A%2F%2Fgithub.com%2Fmadeagency%2Fxlr.svg?type=large)](https://app.fossa.io/projects/git%2Bhttps%3A%2F%2Fgithub.com%2Fmadeagency%2Fxlr?ref=badge_large)