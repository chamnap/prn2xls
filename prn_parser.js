var fs = require('fs');
var _ = require('lodash');
var InvoiceParser  = require(__dirname + '/invoice_parser');

var PrnParser = function(prnFile) {

  // contents
  var content   = fs.readFileSync(prnFile).toString();
  var contents  = content.split('VAT INVOICE');
  contents      = contents.slice(1, contents.length);

  var invoices = [];
  contents.forEach(function(content, index) {
    var parser    = InvoiceParser(content);
    var invoice   = parser.attributes;

    invoices.push(invoice);
  });

  this.invoices = invoices;
};

module.exports = function(prnFile) {
  return new PrnParser(prnFile);
}