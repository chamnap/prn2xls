'use strict';
var fs        = require('fs');
var fse       = require('node-fs-extra');
var Excel     = require('exceljs');
var path      = require('path');
var _         = require('lodash');
var unoconv   = require('unoconv2');

//  
var dialog    = require('electron').dialog;
var exec      = require('child_process').exec;
var util      = require('util');

var PrnParser    = require('./prn_parser');
var ExcelUpdater = require('./excel_updater');

var findCustomer = function(invoice, customers) {
  return _.find(customers, function(customer) { return customer.en_name == invoice.customer.en_name });
};

var convertXls = function(prnFile, destinationDirectory, options, callback) {
  if (_.isFunction(options)) {
    callback = options;
    options = {};
  }

  var destinationDirectory = destinationDirectory || __dirname;
  var prnParser = PrnParser(prnFile);
  var invoices  = prnParser.invoices;
  var missings  = [];
  var customers = options.customers;

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

// options has two keys: customers and unoconvPath
var convertPdf = function(prnFile, destinationDirectory, options, callback) {
  // if (_.isFunction(options)) {
  //   callback = options;
  //   options = {};
  // }

  // convertXls(prnFile, destinationDirectory, options, function(error, excelPath) {
  //   if (error) {
  //     return;
  //   }

  //   var pdfPath = excelPath.replace('.xlsx', '.pdf');
  //   var unoconvOptions = { bin: options.unoconvPath };
  //   unoconv.convert(excelPath, 'pdf', unoconvOptions, function (err, result) {
  //     fs.writeFileSync(pdfPath, result);
  //     fse.removeSync(excelPath);
  //     callback(null, pdfPath);
  //   });
  // });

  var prnParser = PrnParser(prnFile);
  var invoices  = prnParser.invoices;
  var invoice   = invoices[0];

  var htmlFileName = 'pdf.html', 
    pdfFileName    = 'page.pdf';

  var child = exec('wkhtmltopdf -s A4 -L 15.05mm -R 19.05mm -T 19.05mm -B 19.05mm ' + htmlFileName + ' page.pdf', function(err, stdout, stderr) {
    if(err) { throw err; }
    util.log(stderr);
  });


  // dialog.showMessageBox({ message: invoiceProp, buttons:[] })
};

module.exports = {
  xls: convertXls,
  pdf: convertPdf
};