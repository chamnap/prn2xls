var fs = require('fs');
var _ = require('lodash');

var InvoiceParser = function(content) {

  // content
  this.content        = content;

  // company
  var companyName     = this.matchContent('Bill To:', 'Invoice  No.:');
  var companyAddress1 = this.matchBeforeContent('Invoice  Dt.:');
  var companyAddress2 = this.matchBeforeContent('S.O.No.');

  // invoice
  var invoiceNumber   = this.matchAfterContent('Invoice  No.:');
  var invoiceDate     = this.matchAfterContent('Invoice  Dt.:');

  // contract
  var contractNumber  = this.matchAfterContent('Contract No.:');
  var contractDate    = this.matchAfterContent('Contract Dt.:');
  var contractTerms1  = this.matchAfterContent('Terms:');
  var contractTerms2  = this.matchAfterContent(contractTerms1);

  // total
  var tin             = this.matchContent('VAT TIN:', 'Page No.:');
  var pageNumber      = this.matchAfterContent('Page No.:');
  var cCode           = this.matchContent('C.Code :', 'Contract No.:');
  var vat             = this.matchContent('VAT    :', 'Contract Dt.:');
  var total           = this.matchAfterContent('Total  :');
  var vat10           = this.matchAfterContent('VAT 10%:');
  var grantTotal      = this.matchAfterContent('GRAND TOTAL:');
  var contact         = this.matchContent('Contact:', 'Terms:');
  var soNumber        = this.matchAfterContent('S.O.No.     :');
  var subTotal        = this.matchAfterContent('Sub Total Page No.  1:');

  // account
  var accountName     = this.matchContent('Account Name:', '...................');
  var accountNumber   = this.matchAfterContent('Account No  :');
  var bankName        = this.matchAfterContent('Bank Name   :');
  var swiftCode       = this.matchAfterContent('SWIFT CODE:')

  this.attributes = {
    tin: tin,
    pageNumber: pageNumber,
    cCode: cCode,
    vat: vat,
    total: total,
    vat10: vat10,
    grantTotal: grantTotal,
    contact: contact,
    soNumber: soNumber,
    subTotal: subTotal,
    company: {
      name: companyName,
      address1: companyAddress1,
      address2: companyAddress2
    },
    invoice: {
      number: invoiceNumber,
      date: invoiceDate,
    },
    contract: {
      number: contractNumber,
      date: contractDate,
      terms1: contractTerms1,
      terms2: contractTerms2
    },
    account: {
      name: accountName,
      number: accountNumber,
      bank_name: bankName,
      swift_code: swiftCode
    },
    lineItems: this.getLineItems()
  };
};

InvoiceParser.prototype = {
  matchContent: function(before, after) {
    before = _.escapeRegExp(before);
    after  = _.escapeRegExp(after);

    var reString;
    if (before && !after) {
      reString = before + '(.+)';
    }

    if (after && !before) {
      reString = '(.+?)' + after;
    }

    if (before && after) {
      reString = before + '(.+?)' + after;
    }

    var regx     = new RegExp(reString);
    var match    = this.content.match(regx);

    if (match && match.length > 1) {
      return match[1].trim();
    } else {
      return null;
    }
  },

  matchBeforeContent: function(before) {
    return this.matchContent(null, before);
  },

  matchAfterContent: function(after) {
    return this.matchContent(after, null);
  },

  getLineItems: function() {
    var lines = this.content.split("\n");

    var start = null;
    for(var i=0; i<lines.length; i++) {
      var line = lines[i].trim();
      if (line == '_______________________________________________________________________') {
        start = i + 4;
        break;
      }
    }

    var lineItems = [];
    for(var i=start; i<lines.length; i++) {
      var line = lines[i].trim();
      if (line.indexOf('Total  :') != -1 || line.indexOf('Sub Total Page No.') != -1) {
        break;
      }

      var lineItem = this.parseLineItem(line);
      if (lineItem && !this.isBlankLineItem(lineItem)) {
        lineItems.push(lineItem);
      }
    }

    return lineItems;
  },

  isBlankLineItem: function(lineItem) {
    var unique = _.uniq(_.values(lineItem));

    return unique.length == 1 && unique[0] == null;
  },

  parseLineItem: function(line) {
    var attributes = { line: null, codeItem: null, description: null, spot: null, amount: null };

    // replace the overflow fields
    line       = line.replace(/ {30,}\d+(?:\.\d+)?$/, '');
    var values = line.split(/\s{2,}/);

    if (values.length == 6) {
      attributes.line         = values[0];
      attributes.codeItem     = values[1];
      attributes.description  = values[2];
      attributes.spot         = values[3];
      attributes.amount       = values[4];
    } else if(values.length == 5) {
      attributes.line         = values[0];
      attributes.codeItem     = values[1];
      attributes.description  = values[2];
      attributes.spot         = values[3];

      if (values[4].match(/\.|,/)) {
        attributes.amount = values[4];
      }
    } else if (values.length == 4) {
      if (!isNaN(parseInt(values[0]))) {
        attributes.line = values[0];
      }

      if (values[1].indexOf(' ') == -1) {
        attributes.codeItem = values[1];
      }

      if (values[2].indexOf(' ') != -1) {
        attributes.description = values[2];
      }

      if (!isNaN(parseInt(values[3]))) {
        attributes.spot = values[3];
      }
    } else if (values.length == 3) {
      if (values[0].indexOf(' ') == -1) {
        attributes.codeItem = values[0];
      }

      if (values[1].indexOf(' ') != -1) {
        attributes.description = values[1];
      }

      if (values[2].match(/\.|,/)) {
        attributes.amount = values[2];
      }
    } else if (values.length == 2) {
      if (values[0] == '-') {
        attributes.codeItem = values[0];
      }

      if (values[1].indexOf(' ') != -1) {
        attributes.description = values[1];
      }
    } else if (values.length <= 1) {
      return null;
    }

    return attributes;
  }
};

module.exports = function(path) {
  return new InvoiceParser(path);
}