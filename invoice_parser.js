var fs = require('fs');
var _ = require('lodash');

var InvoiceParser = function(content) {

  // content
  this.content        = content;

  // company
  var customerName     = this.matchContent('Bill To:', 'Invoice  No.:');
  var customerAddress1 = this.matchBeforeContent('Invoice  Dt.:');
  var customerAddress2 = this.matchBeforeContent('S.O.No.');

  // invoice
  var invoiceNumber   = this.matchAfterContent('Invoice  No.:');
  var invoiceDate     = this.matchAfterContent('Invoice  Dt.:');

  // contract
  var contractNumber  = this.matchAfterContent('Contract No.:');
  var contractDate    = this.matchAfterContent('Contract Dt.:');
  var contractTerms1  = this.matchAfterContent('Terms:');
  if(contractTerms1 == '7 days after the') {
    var contractTerms2  = 'date of receiving invoice';
  }

  // total
  var tin             = this.matchContent('VAT TIN:', 'Page No.:');
  var pageNumber      = this.matchAfterContent('Page No.:');
  var khmerPageNumber = this.khmerNumber(pageNumber);
  var cCode           = this.matchContent('C.Code :', 'Contract No.:');
  var vat             = this.matchContent('VAT    :', 'Contract Dt.:');
  var total           = this.matchAfterContent('Total  :');
  var vat10           = this.matchAfterContent('VAT 10%:');
  var grandTotal      = this.matchAfterContent('GRAND TOTAL:');
  var contact         = this.matchContent('Contact:', 'Terms:');
  var soNumber        = this.matchAfterContent('S.O.No.     :');
  var subTotalString  = this.matchAfterContent('Sub Total Page No.');
  var subTotal;

  if (subTotalString) {
    var subTotalArray = subTotalString.split(':');
    subTotal = subTotalArray[subTotalArray.length-1].trim();
  }

  // account
  var accountName     = this.matchContent('Account Name:', '...................');
  var accountNumber   = this.matchAfterContent('Account No  :');
  var bankName        = this.matchAfterContent('Bank Name   :');
  var swiftCode       = this.matchAfterContent('SWIFT CODE:')

  this.attributes = {
    tin: tin,
    pageNumber: pageNumber,
    khmerPageNumber: khmerPageNumber,
    cCode: cCode,
    vat: vat,
    total: total,
    vat10: vat10,
    grandTotal: grandTotal,
    contact: contact,
    soNumber: soNumber,
    subTotal: subTotal,
    customer: {
      en_name: customerName,
      en_address1: customerAddress1,
      en_address2: customerAddress2
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
        // var keys = Object.keys(lineItem);

        // console.log('------line-------');
        // for (var j = 0; j < keys.length; j++) {
        //     var val = lineItem[keys[j]];
        //     console.log(val);
        // }

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

    var LINE        = /^\d{1,3}$/;
    var CODE_ITEM   = /^I[A-Z0-9]+$|^-/;
    var DESCRIPTION = /(^[a-zA-Z0-9-\s]*)$/g;
    var SPOT        = /^\d{1,3}$/;
    var AMOUNT      = /^[0-9,\.]+$/;

    for(var i =0; i < values.length; i++) {
      var value = values[i];

      if(LINE.test(value) && attributes.line == null) {
        attributes.line = value.match(LINE)[0];

      } else if (CODE_ITEM.test(value) && attributes.codeItem == null ) {
        attributes.codeItem = value.match(CODE_ITEM)[0];

      } else if (DESCRIPTION.test(value) && !SPOT.test(value)) {
        var description = value.match(DESCRIPTION)[0];

        if(attributes.description == null) {
          attributes.description = description;
        } else {
          attributes.description = attributes.description + ' ' + description;
        }

      } else if (SPOT.test(value) && attributes.spot == null) {
        attributes.spot = value.match(SPOT)[0];

      } else if (AMOUNT.test(value) && attributes.amount == null) {
        attributes.amount = value.match(AMOUNT)[0];
      }
    }

    return attributes;
  },

  khmerNumber: function(number) {
    number = number + '';

    number = number.replace(/1/gm, '១');
    number = number.replace(/2/gm, '២');
    number = number.replace(/3/gm, '៣');
    number = number.replace(/4/gm, '៤');
    number = number.replace(/5/gm, '៥');
    number = number.replace(/6/gm, '៦');
    number = number.replace(/7/gm, '៧');
    number = number.replace(/8/gm, '៨');
    number = number.replace(/9/gm, '៩');
    number = number.replace(/0/gm, '០');

    return number;
  }
};

module.exports = function(path) {
  return new InvoiceParser(path);
}