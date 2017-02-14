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
const generateWorkbook = ({name}) => `${sheetsFront}${sheet(name)}${sheetsBack}`
// generate workbook content types
const generateContentTypes = ({name}) => `${contentTypeFront}${contentTypeOverride(name)}${contentTypeBack}`
// generate workbook relationships
const generateRelationships = ({name}) => `${relationshipFront}${sheetRelationship(name)}${relationshipBack}`

function xlr (config) {
  const xlsx = nodeZip(xlsTemplate, { base64: true, checkCRC32: false })

  if (config.stylesXmlFile) {
    const styles = fs.readFileSync(config.stylesXmlFile, 'utf8')
    if (styles) {
      xlsx.file('xl/styles.xml', styles)
    }
  }

  // write constants and static files
  xlsx.file('xl/workbook.xml', generateWorkbook(config))
  xlsx.file('xl/_rels/workbook.xml.rels', generateRelationships(config))
  xlsx.file('_rels/.rels', relationships)
  xlsx.file('[Content_Types].xml', generateContentTypes(config))

  // replace all characters which would cause Excel to break
  const fileName = config.name ? config.name.replace(/[*?\]\[\/\/][ ]/g, '-') : 'sheet'
  // generate worksheet
  xlsx.file(`xl/worksheets/${fileName}.xml`, Sheet(config))
  return xlsx.generate({ base64: false, compression: 'DEFLATE' })
}

module.exports = xlr
