'use strict';

var expect = require('chai').expect;
var Parser = require('../index');

describe('#Parser', function() {
  var path = __dirname + '/support/source.prn';

  it('should have #tin', function() {
    var parser = Parser(path);
    expect(parser.tin).to.equal('100046029');
  });

  it('should have #cCode', function() {
    var parser = Parser(path);
    expect(parser.cCode).to.equal('CAMS0002');
  });

  it('should have #vat', function() {
    var parser = Parser(path);
    expect(parser.vat).to.equal('901501217');
  });

  it('should have #total', function() {
    var parser = Parser(path);
    expect(parser.total).to.equal('727.27');
  });

  it('should have #vat10', function() {
    var parser = Parser(path);
    expect(parser.vat10).to.equal('72.73');
  });

  it('should have #grantTotal', function() {
    var parser = Parser(path);
    expect(parser.grantTotal).to.equal('800.00');
  });

  it('should have #company=>name', function() {
    var parser = Parser(path);
    expect(parser.company.name).to.equal('Ads Marketing Solution Co., Ltd');
  });

  it('should have #company=>address', function() {
    var parser = Parser(path);
    expect(parser.company.address).to.equal('#90Eo, St.02 A, Sangkat Phnom Penh Thmey, Khan Sen Sok, Phnom Penh');
  });

  it('should have #invoice=>number', function() {
    var parser = Parser(path);
    expect(parser.invoice.number).to.equal('MI-A080964');
  });

  it('should have #invoice=>date', function() {
    var parser = Parser(path);
    expect(parser.invoice.date).to.equal('29/02/2016');
  });

  it('should have #contract=>number', function() {
    var parser = Parser(path);
    expect(parser.contract.number).to.equal('YBR1602');
  });

  it('should have #contract=>date', function() {
    var parser = Parser(path);
    expect(parser.contract.date).to.equal('1-Feb-16');
  });

  it('should have #contract=>number', function() {
    var parser = Parser(path);
    expect(parser.contract.number).to.equal('YBR1602');
  });

  it('should have #contract=>terms1', function() {
    var parser = Parser(path);
    expect(parser.contract.terms1).to.equal('7 days after the');
  });

  it('should have #contract=>terms2', function() {
    var parser = Parser(path);
    expect(parser.contract.terms2).to.equal('date of receiving invoice');
  });

  it('should have #contract=>name', function() {
    var parser = Parser(path);
    expect(parser.contract.name).to.equal('ADS');
  });

  it('should have #account=>name', function() {
    var parser = Parser(path);
    expect(parser.account.name).to.equal('Cambodian Broadcasting Service');
  });

  it('should have #account=>number', function() {
    var parser = Parser(path);
    expect(parser.account.number).to.equal('USD 116345');
  });

  it('should have #account=>bank_name', function() {
    var parser = Parser(path);
    expect(parser.account.bank_name).to.equal('ANZ Royal Bank (Cambodia) Ltd');
  });

  it('should have #account=>swift_code', function() {
    var parser = Parser(path);
    expect(parser.account.swift_code).to.equal('ANZBKHPP');
  });

  describe('#lineItems', function() {
    it('have 6 rows', function() {
      var parser = Parser(path);

      expect(parser.lineItems.length).to.equal(7);
    });

    it('#first', function() {
      var parser = Parser(path);
      var lineItem = parser.lineItems[0];

      expect(lineItem.lineItem).to.equal(null);
      expect(lineItem.codeItem).to.equal('-');
      expect(lineItem.description).to.equal('Bacchus Ads Feb 16');
      expect(lineItem.spot).to.equal(null);
      expect(lineItem.amount).to.equal(null);
    });

    it('#second', function() {
      var parser = Parser(path);
      var lineItem = parser.lineItems[1];
      console.log(lineItem);

      expect(lineItem.lineItem).to.equal(null);
      expect(lineItem.codeItem).to.equal('I99003');
      expect(lineItem.description).to.equal('Package as per contract');
      expect(lineItem.spot).to.equal(null);
      expect(lineItem.amount).to.equal('727.27');
    });

    it('#third', function() {
      var parser = Parser(path);
      var lineItem = parser.lineItems[2];

      expect(lineItem.lineItem).to.equal('006');
      expect(lineItem.codeItem).to.equal('I10A2512');
      expect(lineItem.description).to.equal('SP Package-25s 1200-1300');
      expect(lineItem.spot).to.equal('17');
      expect(lineItem.amount).to.equal('1,000.00');
    });

    it('#fifth', function() {
      var parser = Parser(path);
      var lineItem = parser.lineItems[4];

      expect(lineItem.lineItem).to.equal('008');
      expect(lineItem.codeItem).to.equal('I10A4520');
      expect(lineItem.description).to.equal('SP Package-45s 2000-2100');
      expect(lineItem.spot).to.equal('14');
      expect(lineItem.amount).to.equal(null);
    });

    it('#seventh', function() {
      var parser = Parser(path);
      var lineItem = parser.lineItems[6];

      expect(lineItem.lineItem).to.equal('010');
      expect(lineItem.codeItem).to.equal('I10A6520');
      expect(lineItem.description).to.equal('SP Package-65s 2000-2100');
      expect(lineItem.spot).to.equal('7');
      expect(lineItem.amount).to.equal('2,000.00');
    });
  });
});