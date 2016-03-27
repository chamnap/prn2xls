var fs = require('fs');
var _ = require('lodash');
var InvoiceParser  = require(__dirname + '/invoice_parser');

var PrnParser = function(prnFile) {

  // contents
  var content   = fs.readFileSync(prnFile).toString();
  var contents  = content.split('VAT INVOICE');
  contents      = contents.slice(1, contents.length);

  var invoices = [];
  var previousInvoice = null;
  contents.forEach(function(content, index) {
    var parser    = InvoiceParser(content);
    var invoice   = parser.attributes;
    var lineItems = [];

    if (invoice.pageNumber == 1) {
      invoices.push(invoice);
    } else if (invoice.pageNumber > 1) {
      lineItems = _.concat(invoice.lineItems, previousInvoice.lineItems);
      invoices[index-1].lineItems = _.flatten(lineItems);
    }

    previousInvoice = invoice;
  });

  this.invoices = invoices;
};

module.exports = function(prnFile) {
  return new PrnParser(prnFile);
}