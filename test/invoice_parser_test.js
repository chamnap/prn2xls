'use strict';

var fs        = require('fs');
var _         = require('lodash');
var expect    = require('chai').expect;
var InvoiceParser = require('../invoice_parser');

describe('#InvoiceParser, multiple pages', function() {
  var sourceFile  = __dirname + '/support/several_pages.prn';
  var content     = fs.readFileSync(sourceFile).toString();
  var contents    = content.split('VAT INVOICE');
  contents        = contents.slice(1, contents.length);

  describe('#lineItems', function() {
    describe('1st content', function() {
      it('#line', function() {
        var invoice = InvoiceParser(contents[0]).attributes;
        var lines   = _.map(invoice.lineItems, function(lineItem) { return lineItem.line; });

        expect(lines).to.deep.equal([null, '002', '003', '004', '005', '006', '007', '008', '009', '010', '011', '012', '013', '014', '015', '016', '017', null, '019', '020']);
      });

      it('#spot', function() {
        var invoice = InvoiceParser(contents[0]).attributes;
        var spots   = _.map(invoice.lineItems, function(lineItem) { return lineItem.spot; });

        expect(spots).to.deep.equal([null, '13', '23', '8', '19', '17', '9', '14', '11', '13', '40', '15', '15', '8', '8', '83', '50', null, '20', '20']);
      });

      it('#itemCode', function() {
        var invoice   = InvoiceParser(contents[0]).attributes;
        var codeItems = _.map(invoice.lineItems, function(lineItem) { return lineItem.codeItem; });

        expect(codeItems).to.deep.equal(['-', 'I10A4511', 'I10A2511', 'I10A4512', 'I10A2512', 'I10A2512', 'I10A4512', 'I10A4520', 'I10A2520', 'I0004514', 'I10A2506', 'I10A4520', 'I10A2520', 'I0002514', 'I0004514', 'I0002514', 'I0002514', '-', 'I10A4512', 'I10A4517']);
      });

      it('subTotal is correct with amount', function() {
        var invoice = InvoiceParser(contents[0]).attributes;

        var subTotal = _.sumBy(invoice.lineItems, function(lineItem) { return (lineItem.amount == null) ? 0.0 : parseFloat(lineItem.amount.replace(',', '')); });
        expect(invoice.lineItems.length).to.equal(20);
        expect(invoice.subTotal.replace(',', '')).to.equal(subTotal.toFixed(2));
      });
    });
  });
});

describe('InvoiceParser, single page', function() {
  var sourceFile  = __dirname + '/support/one_page.PRN';
  var content     = fs.readFileSync(sourceFile).toString();

  it('should have #tin', function() {
    var parser = InvoiceParser(content).attributes;

    expect(parser.tin).to.equal('100046029');
  });

  it('should have #pageNumber', function() {
    var parser = InvoiceParser(content).attributes;

    expect(parser.pageNumber).to.equal('1');
  });

  it('should have #cCode', function() {
    var parser = InvoiceParser(content).attributes;

    expect(parser.cCode).to.equal('CAMS0002');
  });

  it('should have #vat', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.vat).to.equal('901501217');
  });

  it('should have #total', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.total).to.equal('727.27');
  });

  it('should have #vat10', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.vat10).to.equal('72.73');
  });

  it('should have #grandTotal', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.grandTotal).to.equal('800.00');
  });

  it('should have #soNumber', function() {
    var parser = InvoiceParser(content).attributes;

    expect(parser.soNumber).to.equal('MI-A083857');
  });

  it('should have #contact', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.contact).to.equal('ADS');
  });

  it('should have #customer=>en_name', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.customer.en_name).to.equal('Ads Marketing Solution Co., Ltd');
  });

  it('should have #customer=>en_address1', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.customer.en_address1).to.equal('#90Eo, St.02 A, Sangkat Phnom Penh');
  });

  it('should have #customer=>en_address2', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.customer.en_address2).to.equal('Thmey, Khan Sen Sok, Phnom Penh');
  });

  it('should have #invoice=>number', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.invoice.number).to.equal('MI-A080964');
  });

  it('should have #invoice=>date', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.invoice.date).to.equal('29/02/2016');
  });

  it('should have #contract=>number', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.contract.number).to.equal('YBR1602');
  });

  it('should have #contract=>date', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.contract.date).to.equal('1-Feb-16');
  });

  it('should have #contract=>number', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.contract.number).to.equal('YBR1602');
  });

  it('should have #contract=>terms1', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.contract.terms1).to.equal('7 days after the');
  });

  it('should have #contract=>terms2', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.contract.terms2).to.equal('date of receiving invoice');
  });

  it('should have #account=>name', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.account.name).to.equal('Cambodian Broadcasting Service');
  });

  it('should have #account=>number', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.account.number).to.equal('USD 116345');
  });

  it('should have #account=>bank_name', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.account.bank_name).to.equal('ANZ Royal Bank (Cambodia) Ltd');
  });

  it('should have #account=>swift_code', function() {
    var parser = InvoiceParser(content).attributes;
    expect(parser.account.swift_code).to.equal('ANZBKHPP');
  });

  describe('#lineItems', function() {
    it('have 6 rows', function() {
      var parser = InvoiceParser(content).attributes;

      expect(parser.lineItems.length).to.equal(8);
    });

    it('#first', function() {
      var parser = InvoiceParser(content).attributes;
      var lineItem = parser.lineItems[0];

      expect(lineItem.line).to.equal(null);
      expect(lineItem.codeItem).to.equal('-');
      expect(lineItem.description).to.equal('Bacchus Ads Feb 16');
      expect(lineItem.spot).to.equal(null);
      expect(lineItem.amount).to.equal(null);
    });

    it('#second', function() {
      var parser = InvoiceParser(content).attributes;
      var lineItem = parser.lineItems[1];

      expect(lineItem.line).to.equal(null);
      expect(lineItem.codeItem).to.equal('I99003');
      expect(lineItem.description).to.equal('Package as per contract');
      expect(lineItem.spot).to.equal(null);
      expect(lineItem.amount).to.equal('727.27');
    });

    it('#third', function() {
      var parser = InvoiceParser(content).attributes;
      var lineItem = parser.lineItems[2];

      expect(lineItem.line).to.equal('006');
      expect(lineItem.codeItem).to.equal('I10A2512');
      expect(lineItem.description).to.equal('SP Package-25s 1200-1300');
      expect(lineItem.spot).to.equal('17');
      expect(lineItem.amount).to.equal('1,000.00');
    });

    it('#fifth', function() {
      var parser = InvoiceParser(content).attributes;
      var lineItem = parser.lineItems[4];

      expect(lineItem.line).to.equal('008');
      expect(lineItem.codeItem).to.equal('I10A4520');
      expect(lineItem.description).to.equal('SP Package-45s 2000-2100');
      expect(lineItem.spot).to.equal('14');
      expect(lineItem.amount).to.equal(null);
    });

    it('#seventh', function() {
      var parser = InvoiceParser(content).attributes;
      var lineItem = parser.lineItems[6];

      expect(lineItem.line).to.equal('010');
      expect(lineItem.codeItem).to.equal('I10A6520');
      expect(lineItem.description).to.equal('SP Package-65s 2000-2100');
      expect(lineItem.spot).to.equal('7');
      expect(lineItem.amount).to.equal('2,000.00');
    });

    it('#eight', function() {
      var parser = InvoiceParser(content).attributes;
      var lineItem = parser.lineItems[7];

      expect(lineItem.line).to.equal('011');
      expect(lineItem.codeItem).to.equal('I10A8520');
      expect(lineItem.description).to.equal('SP Package-85s 2000-2100');
      expect(lineItem.spot).to.equal('7');
      expect(lineItem.amount).to.equal(null);
    });
  });
});