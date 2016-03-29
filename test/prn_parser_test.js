'use strict';

var fs        = require('fs');
var expect    = require('chai').expect;
var PrnParser = require('../prn_parser');
var InvoiceParser = require('../invoice_parser');

describe('PrnParser', function() {
  var prnFile  = __dirname + '/support/several_pages.prn';

  it('has 5 invoices', function() {
    var parser = PrnParser(prnFile);

    expect(parser.invoices.length).to.equal(5);
  });

  describe('lineItems', function() {
    it('1st invoice has 20 items', function() {
      var parser  = PrnParser(prnFile);
      var invoice = parser.invoices[0];

      expect(invoice.invoice.number).to.equal('MI-A080804');
      expect(invoice.lineItems.length).to.equal(20);
    });

    it('2nd invoice has 11 items', function() {
      var parser  = PrnParser(prnFile);
      var invoice = parser.invoices[1];

      expect(invoice.invoice.number).to.equal('MI-A080804');
      expect(invoice.lineItems.length).to.equal(11);
    });

    it('3rd invoice has 7 items', function() {
      var parser  = PrnParser(prnFile);
      var invoice = parser.invoices[2];

      expect(invoice.invoice.number).to.equal('MI-A080805');
      expect(invoice.lineItems.length).to.equal(7);
    });

    it('4th invoice has 7 items', function() {
      var parser  = PrnParser(prnFile);
      var invoice = parser.invoices[3];

      expect(invoice.invoice.number).to.equal('MI-A080806');
      expect(invoice.lineItems.length).to.equal(7);
    });

    it('5th invoice has 3 items', function() {
      var parser  = PrnParser(prnFile);
      var invoice = parser.invoices[4];

      expect(invoice.invoice.number).to.equal('MI-A080807');
      expect(invoice.lineItems.length).to.equal(3);
    });
  });
});