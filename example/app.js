const express = require('express')
const xlr = require('../index')
const config = require('./config.js')
const app = express()

app.get('/', function (req, res) {
  const result = xlr(config)
  res.setHeader('Content-Type', 'application/vnd.openxmlformats')
  res.setHeader('Content-Disposition', 'attachment; filename=Report.xlsx')
  res.send(new Buffer(result, 'binary'))
})

app.listen(3001, function () {
  console.log('listening on port 3001')
})