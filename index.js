const nodeZip = require('node-zip');
const fs = require('fs');
const Sheet = require('./sheet');
const constants = require('./constants');

// generate workbook
function generateWorkbook (configs) {
  const workbook = configs.map(function(config, i) {
    const index = i + 1;
    return `<sheet name="${config.name}" sheetId="${index}" r:id="rId${index}"/>`;
  }).join('');

  return `${constants.sheetsFront}${workbook}${constants.sheetsBack}`;
}

// generate workbook content types
function generateContentType(configs) {
  const contentTypes = configs.map(function (config) {
    return `<Override PartName="/${config.fileName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />`;
  }).join('');

  return `${constants.contentTypeFront}${contentTypes}${constants.contentTypeBack}`;
}

// generate workbook rels
function generateRel(configs) {
  const rels = configs.map(function (config, i) {
    const index = i + 1;
    return `<Relationship Id="rId${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${config.name}.xml"/>`;
  }).join('');

  return `${constants.relFront}${rels}${constants.relBack}`;
}


function xlr (config) {
  const xlsx = nodeZip(constants.xlsTemplate, { base64: true, checkCRC32: false });
  const conf = typeof config === 'object' ? [config] : config;

  // only write the stylesheet once, so use first conf object
  if (conf[0].stylesXmlFile) {
    const styles = fs.readFileSync(config.stylesXmlFile, 'utf8');
    if (styles) {
      xlsx.file('xl/styles.xml', styles);
    }
  }

  // write constants and static files
  xlsx.file('xl/workbook.xml', generateWorkbook(conf));
  xlsx.file('xl/_rels/workbook.xml.rels', generateRel(conf));
  xlsx.file('_rels/.rels', constants.rels);
  xlsx.file('[Content_Types].xml', generateContentType(conf));

  // generate worksheets
  conf.forEach(function (cfg) {
    // replace all characters which would cause Excel to break
    const fileName = cfg.name ? cfg.name.replace(/[*?\]\[\/\/][ ]/g, '') : 'sheet';
    xlsx.file(`xl/worksheets/${fileName}.xml`, Sheet(cfg));
  });

  return xlsx.generate({ base64: false, compression: 'DEFLATE' });
}

module.exports = xlr;