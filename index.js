const nodeZip = require('node-zip')
const fs = require('fs')
const Sheet = require('./sheet')
const { 
  sheetsFront,
  sheetsBack,
  contentTypeFront,
  contentTypeBack,
  relationshipFront,
  relationshipBack,
  relationships,
  xlsTemplate,
  sheet,
  contentTypeOverride,
  sheetRelationship
} = require('./constants')

// generate workbook
function generateWorkbook (configs) {
  const workbook = configs.map((config, i) => sheet(config.name, i + 1)).join('')
  return `${sheetsFront}${workbook}${sheetsBack}`
}

// generate workbook content types
function generateContentTypes(configs) {
  const contentTypes = configs.map((config) => contentTypeOverride(config.fileName)).join('')
  return `${contentTypeFront}${contentTypes}${contentTypeBack}`
}

// generate workbook relationships
function generateRelationships(configs) {
  const relationships = configs.map((config, i) => sheetRelationship(config.name, i + 1)).join('')
  return `${relationshipFront}${relationships}${relationshipBack}`
}


function xlr (options) {
  const xlsx = nodeZip(xlsTemplate, { base64: true, checkCRC32: false })
  const configs = typeof options === 'object' ? [options] : options

  // only write the stylesheet once, so use first conf object
  if (configs[0].stylesXmlFile) {
    const styles = fs.readFileSync(configs[0].stylesXmlFile, 'utf8')
    if (styles) {
      xlsx.file('xl/styles.xml', styles)
    }
  }

  // write constants and static files
  xlsx.file('xl/workbook.xml', generateWorkbook(configs))
  xlsx.file('xl/_rels/workbook.xml.rels', generateRelationships(configs))
  xlsx.file('_rels/.rels', relationships)
  xlsx.file('[Content_Types].xml', generateContentTypes(configs))

  // generate worksheets
  configs.forEach((config) => {
    // replace all characters which would cause Excel to break
    const fileName = config.name ? config.name.replace(/[*?\]\[\/\/][ ]/g, '-') : 'sheet'
    xlsx.file(`xl/worksheets/${fileName}.xml`, Sheet(config))
  })

  return xlsx.generate({ base64: false, compression: 'DEFLATE' })
}

module.exports = xlr
