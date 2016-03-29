'use strict';
var fs        = require('fs');
var fse       = require('node-fs-extra');
var Excel     = require('exceljs');
var path      = require('path');
var _         = require('lodash');
var PrnParser    = require(__dirname + '/prn_parser');
var ExcelUpdater = require(__dirname + '/excel_updater');

var findCustomer = function(invoice, customers) {
  return _.find(customers, function(customer) { return customer.en_name == invoice.customer.en_name });
};

module.exports.convert = function(prnFile, destinationDirectory, customers, callback) {
  var destinationDirectory = destinationDirectory || __dirname;
  var prnParser = PrnParser(prnFile);
  var invoices  = prnParser.invoices;
  var missings  = [];

  _.forEach(invoices, function(invoice) {
    var customer = findCustomer(invoice, customers);
    if (customer) {
      invoice.customer.kh_name     = customer.kh_name;
      invoice.customer.kh_address1 = customer.kh_address1;
      invoice.customer.kh_address2 = customer.kh_address2;
    } else {
      missings.push(invoice.customer.en_name);
    }
  });
  missings = _.uniq(missings);

  if (customers && missings.length > 0) {
    callback(missings);
  }

  // copy from sample
  var blankFilePath  = __dirname + '/blank.xlsx';
  var baseName       = path.basename(prnFile.toLowerCase(), '.prn');
  var newFilePath    = destinationDirectory + '/' + baseName + '.xlsx';
  fse.copySync(blankFilePath, newFilePath);

  // exceljs
  var workbook  = new Excel.Workbook();
  workbook.xlsx.readFile(newFilePath)
    .then(function() {
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