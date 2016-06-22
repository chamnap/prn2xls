'use strict';
var fs        = require('fs');
var fse       = require('node-fs-extra');
var Excel     = require('exceljs');
var path      = require('path');
var _         = require('lodash');

var wkhtmltopdf = require('wkhtmltopdf');
var jade        = require('jade');

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

var convertPdf = function(prnFile, destinationDirectory, options, callback) {
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

  var baseName     = path.basename(prnFile.toLowerCase(), '.prn');
  var pdfPath      = destinationDirectory + '/' + baseName + '.pdf';
  var jadeTemplate = fs.readFileSync(__dirname + '/pdf.pug', 'utf8');
  var fn = jade.compile(jadeTemplate);

  var htmlOutput = '';

  for(var i = 0; i < invoices.length; i++) {
    htmlOutput += fn({ invoice: invoices[i] });
  }

  wkhtmltopdf(htmlOutput,
    {
      pageSize: 'A4',
      marginLeft: '15.05mm',
      marginRight: '19.05mm',
      marginTop: '19.05mm',
      marginBottom: '19.05mm'
    }, function(err){
      if(err) {
        callback(err, null);
      } else {
        callback(null, pdfPath);
      }
    })
    .pipe(fs.createWriteStream(pdfPath));
};

module.exports = {
  xls: convertXls,
  pdf: convertPdf
};