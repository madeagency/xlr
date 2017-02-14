module.exports = {
  stylesXmlFile: __dirname + '/styles.xml',
  name: 'report-is-a-thing',
  columns: [
    {
      type: 'string',
      width: 15
    },
    {
      type: 'string',
      width: 25
    },
    {
      type: 'string',
      width: 30
    }
  ],
  rows: [
    [{ value: 'First Name', style: 1 }, { value: 'Last Name', style: 1 }, { value: 'Email', style: 1 }],
    ['Bruce', 'Wayne', 'info@bat.man'],
    ['Clark', 'Kent', 'info@super.man'],
    ['Testing a merge cell because reasons', '', 'test@email.com'],
    ['Peter', 'Parker', 'info@spider.man']
  ],
  merge: [
    { row: 4, fromColumn: 1, toColumn: 2 } 
  ]
}