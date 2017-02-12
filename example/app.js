const express = require('express');
const xlr = require('../index');
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