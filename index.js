'use strict';

/**
 * Parser
 * @param {string} path
 * @return {JSON}
 */
module.exports = function(path) {
  var fs = require('fs');
  var _ = require('lodash');

  // http://stackoverflow.com/questions/7124778/how-to-match-anything-up-until-this-sequence-of-characters-in-a-regular-expres
  // content.match(/Bill To: (.+?(?=Invoice  No))/)[1].trim();
  // content.match(/(.+?)Invoice  Dt/)[1].trim();
  // content.match(/(.+?)S\.O\.No\./)[1].trim();
  var matchContent = function(content, before, after) {
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
    var match    = content.match(regx);

    if (match && match.length > 1) {
      return match[1].trim();
    } else {
      return null;
    }
  };

  var getLineItems = function(content) {
    var lines = content.split("\n");

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
      if (line.indexOf('Total  :') != -1) {
        break;
      }

      var lineItem = parseLineItem(line);
      if (lineItem) {
        lineItems.push(lineItem);
      }
    }

    return lineItems;
  }

  var parseLineItem = function(line) {
    var attributes = { lineItem: null, codeItem: null, description: null, spot: null, amount: null };
    var values = line.split(/\s{3,}/);

    if (values.length == 6) {
      attributes.lineItem     = values[0];
      attributes.codeItem     = values[1];
      attributes.description  = values[2];
      attributes.spot         = values[3];
      attributes.amount       = values[4];
    } else if(values.length == 5) {
      attributes.lineItem     = values[0];
      attributes.codeItem     = values[1];
      attributes.description  = values[2];
      attributes.spot         = values[3];

      if (values[4].match(/\.|,/).length > 0) {
        attributes.amount = values[4];
      }
    } else if (values.length == 4) {
      if (!isNaN(parseInt(values[0]))) {
        attributes.lineItem = values[0];
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

      if (values[2].match(/\.|,/).length > 0) {
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
  };

  var matchBeforeContent = function(content, before) {
    return matchContent(content, null, before);
  };

  var matchAfterContent = function(content, after) {
    return matchContent(content, after, null);
  };

  var attributes = {};

  var content         = fs.readFileSync(path).toString();

  // company
  var companyName     = matchContent(content, 'Bill To:', 'Invoice  No.:');
  var companyAddress1 = matchBeforeContent(content, 'Invoice  Dt.:');
  var companyAddress2 = matchBeforeContent(content, 'S.O.No.');

  // invoice
  var invoiceNumber   = matchAfterContent(content, 'Invoice  No.:');
  var invoiceDate     = matchAfterContent(content, 'Invoice  Dt.:');

  //
  var tin             = matchContent(content, 'VAT TIN:', 'Page No.:');
  var cCode           = matchContent(content, 'C.Code :', 'Contract No.:');
  var vat             = matchContent(content, 'VAT    :', 'Contract Dt.:');

  // contract
  var contractNumber  = matchAfterContent(content, 'Contract No.:');
  var contractDate    = matchAfterContent(content, 'Contract Dt.:');
  var contractTerms1  = matchAfterContent(content, 'Terms:');
  var contractTerms2  = matchAfterContent(content, contractTerms1);
  var contractName    = matchContent(content, 'Contact:', 'Terms:');

  // total
  var total           = matchAfterContent(content, 'Total  :');
  var vat10           = matchAfterContent(content, 'VAT 10%:');
  var grantTotal      = matchAfterContent(content, 'GRAND TOTAL:');

  // account
  var accountName     = matchContent(content, 'Account Name:', '...................');
  var accountNumber   = matchAfterContent(content, 'Account No  :');
  var bankName        = matchAfterContent(content, 'Bank Name   :');
  var swiftCode       = matchAfterContent(content, 'SWIFT CODE:')

  attributes = {
    tin: tin,
    cCode: cCode,
    vat: vat,
    total: total,
    vat10: vat10,
    grantTotal: grantTotal,
    company: {
      name: companyName,
      address: companyAddress1 + ' ' + companyAddress2
    },
    invoice: {
      number: invoiceNumber,
      date: invoiceDate,
    },
    contract: {
      number: contractNumber,
      date: contractDate,
      terms1: contractTerms1,
      terms2: contractTerms2,
      name: contractName
    },
    account: {
      name: accountName,
      number: accountNumber,
      bank_name: bankName,
      swift_code: swiftCode
    },
    lineItems: getLineItems(content)
  };

  return attributes;
};