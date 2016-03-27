'use strict';

var fs        = require('fs');
var expect    = require('chai').expect;
var InvoiceParser = require('../invoice_parser');

describe('#InvoiceParser, multiple pages', function() {
  var sourceFile  = __dirname + '/support/several_pages.prn';
  var content     = fs.readFileSync(sourceFile).toString();
  var contents    = content.split('VAT INVOICE');
  contents        = contents.slice(1, contents.length);

  describe('#lineItems', function() {
    it('1st content', function() {
      var parser = InvoiceParser(contents[0]).attributes;

      expect(parser.lineItems.length).to.equal(20);

      expect(parser.lineItems[0]).to.deep.equal({ lineItem: null, codeItem: '-', description: 'Ganzberg Beer Ads 02 16', spot: null, amount: null });
      expect(parser.lineItems[1]).to.deep.equal({ lineItem: '002', codeItem: 'I10A4511', description: 'SP Package-45s 1100-1200', spot: '13', amount: '919.19' });
      expect(parser.lineItems[2]).to.deep.equal({ lineItem: '003', codeItem: 'I10A2511', description: 'SP Package-25s 1100-1200', spot: '23', amount: '1,626.26' });
      expect(parser.lineItems[3]).to.deep.equal({ lineItem: '004', codeItem: 'I10A4512', description: 'SP Package-45s 1200-1300', spot: '8', amount: '754.21' });
      expect(parser.lineItems[4]).to.deep.equal({ lineItem: '005', codeItem: 'I10A2512', description: 'SP Package-25s 1200-1300', spot: '19', amount: '1,791.25' });
      expect(parser.lineItems[5]).to.deep.equal({ lineItem: '006', codeItem: 'I10A2512', description: 'SP Package-25s 1200-1300', spot: '17', amount: '1,664.34' });
      expect(parser.lineItems[6]).to.deep.equal({ lineItem: '007', codeItem: 'I10A4512', description: 'SP Package-45s 1200-1300', spot: '9', amount: '881.12' });
    });
  });
});