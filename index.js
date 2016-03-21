'use strict';

module.exports.process = function(source) {
  // require libs
  var fse       = require('node-fs-extra');
  var Excel     = require('exceljs');
  var path      = require('path');
  var PrnParser = require(__dirname + '/prn_parser');
  var ContentUpdater = require(__dirname + '/content_updater');

  // copy from sample
  var sampleFilePath = __dirname + '/sample.xlsx';
  var newFilePath    = __dirname + '/' + path.basename(source, '.PRN') + '.xlsx';
  fse.copySync(sampleFilePath, newFilePath);

  // exceljs
  var workbook  = new Excel.Workbook();
  workbook.xlsx.readFile(newFilePath)
    .then(function() {
      var parser = PrnParser(source);
      var contentUpdater = ContentUpdater(workbook);

      contentUpdater.update(parser.attributes);
    })
    .then(function() {
      return workbook.xlsx.writeFile(newFilePath);
    })
    .then(function() {
      console.log('File: `' + newFilePath + '` is Done!');
    });
};