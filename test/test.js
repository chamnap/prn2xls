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
});