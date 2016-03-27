'use strict';

module.exports.convert = function(prnFile, destinationDirectory, callback) {
  var destinationDirectory = destinationDirectory || __dirname;

  // require libs
  var fs        = require('fs');
  var fse       = require('node-fs-extra');
  var Excel     = require('exceljs');
  var path      = require('path');
  var PrnParser    = require(__dirname + '/prn_parser');
  var ExcelUpdater = require(__dirname + '/excel_updater');

  // copy from sample
  var blankFilePath  = __dirname + '/blank.xlsx';
  var baseName       = path.basename(prnFile.toLowerCase(), '.prn');
  var newFilePath    = destinationDirectory + '/' + baseName + '.xlsx';
  fse.copySync(blankFilePath, newFilePath);

  // exceljs
  var workbook  = new Excel.Workbook();
  workbook.xlsx.readFile(newFilePath)
    .then(function() {
      var prnParser = PrnParser(prnFile);
      var invoices  = prnParser.invoices;
      var excelUpdater = ExcelUpdater(workbook);

      excelUpdater.update(invoices);
    })
    .then(function() {
      return workbook.xlsx.writeFile(newFilePath);
    })
    .then(function() {
      callback(null, newFilePath);
    });
};