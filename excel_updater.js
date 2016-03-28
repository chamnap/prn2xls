var _ = require('lodash');

var ExcelUpdater = function(workbook) {
  this.workbook = workbook;
  this.worksheet = this.workbook.getWorksheet(1);
};

ExcelUpdater.prototype = {
  update: function(invoices) {
    var startingRow = 0;
    var lastRow;
    var self = this;

    _.forEach(invoices, function(invoice) {
      self.updateHeader(invoice, startingRow);

      self.updateLineItems(invoice, startingRow);

      lastRow = self.updateFooter(invoice, startingRow);

      startingRow = lastRow + 12;
    });
  },

  updateHeader: function(invoice, startingRow) {
    this.setCell('B' + (1 + startingRow), 'វិក្កយប័ត្រអាករ');
    this.setCell('B' + (2 + startingRow), 'VAT INVOICE');
    this.setCell('B' + (3 + startingRow), 'លេខអត្តសញ្ញាណកម្ម អតប');
    this.setCell('B' + (4 + startingRow), 'VAT TIN: ' + invoice.tin);
    this.setCell('B' + (6 + startingRow), 'ឈ្មោះអតិថិជន');
    this.setCell('C' + (6 + startingRow), invoice.customer.kh_name);
    this.setCell('B' + (7 + startingRow), 'Bill To:           ');
    this.setCell('C' + (7 + startingRow), invoice.customer.en_name);
    this.setCell('C' + (8 + startingRow), invoice.customer.kh_address1);
    this.setCell('C' + (9 + startingRow), invoice.customer.en_address1);
    this.setCell('C' + (10 + startingRow), invoice.customer.kh_address2);
    this.setCell('C' + (11 + startingRow), invoice.customer.en_address2);

    this.setCell('G' + (3 + startingRow), 'ទំព័រ');
    this.setCell('G' + (4 + startingRow), 'Page No.:        ' + invoice.pageNumber);
    this.setCell('G' + (6 + startingRow), 'លេខរៀងវិក្កយប័ត្រ');
    this.setCell('G' + (7 + startingRow), 'Invoice  No.');
    this.setCell('H' + (7 + startingRow), ':');
    this.setCell('I' + (7 + startingRow), invoice.invoice.number);

    this.setCell('G' + (8 + startingRow), 'កាលបរិចេ្ឆទ');
    this.setCell('G' + (9 + startingRow), 'Invoice  Dt.');
    this.setCell('H' + (9 + startingRow), ':');
    this.setCell('I' + (9 + startingRow), invoice.invoice.date);

    this.setCell('G' + (10 + startingRow), 'លេខរៀងការបញ្ជាទិញ');
    this.setCell('G' + (11 + startingRow), 'S.O.No.     ');
    this.setCell('H' + (11 + startingRow), ':');
    this.setCell('I' + (11 + startingRow), invoice.soNumber);

    this.setCell('B' + (13 + startingRow), 'លេខកូតអតិថិជន');
    this.setCell('B' + (14 + startingRow), 'C.Code :');
    this.setCell('C' + (14 + startingRow), invoice.cCode);
    this.setCell('B' + (15 + startingRow), 'លេខអត្តសញ្ញាណកម្ម អតប');
    this.setCell('B' + (15 + startingRow), ' VAT    :');

    this.setCell('G' + (13 + startingRow), 'កិច្ចសន្យាលេខ');
    this.setCell('G' + (14 + startingRow), 'Contract No.');
    this.setCell('H' + (14 + startingRow), ':');
    this.setCell('I' + (14 + startingRow), invoice.contract.number);
    this.setCell('G' + (15 + startingRow), 'កាលបរិចេ្ឆទនៃកិច្ចសន្យា');
    this.setCell('G' + (16 + startingRow), 'Contract Dt.');
    this.setCell('H' + (16 + startingRow), ':');
    this.setCell('I' + (16 + startingRow), invoice.contract.date);

    this.setCell('B' + (17 + startingRow), 'ទំនាក់ទំនង');
    this.setCell('B' + (18 + startingRow), 'Contact:');
    this.setCell('C' + (18 + startingRow), invoice.contact);

    this.setCell('G' + (17 + startingRow), 'រយ:ពេលទូទាត់');
    this.setCell('I' + (17 + startingRow), '៧ថ្ងៃបន្ទាប់ពីទទួលបានវិក្កយប័ត្រ');
    this.setCell('G' + (18 + startingRow), 'Terms:   7 days after the');
    this.setCell('G' + (19 + startingRow), 'date of receiving invoice');

    this.setCell('B' + (22 + startingRow), 'លេខរៀង');
    this.setCell('C' + (22 + startingRow), 'កូតទំនិញេ/សេវា');
    this.setCell('D' + (22 + startingRow), 'បរិយាយមុខទំនិញ/សេវា');
    this.setCell('G' + (22 + startingRow), 'បរិមាណ');
    this.setCell('I' + (22 + startingRow), 'ចំនួនទឹកប្រាក់');
    this.setCell('B' + (23 + startingRow), 'Line');
    this.setCell('C' + (23 + startingRow), 'Item Code');
    this.setCell('D' + (23 + startingRow), 'Item Description');
    this.setCell('G' + (23 + startingRow), 'No.of Spot');
    this.setCell('I' + (23 + startingRow), 'US$ Amount');
  },

  updateLineItemsHeader: function(startingRow) {
    var rowNumber     = startingRow + 21;
    var borderOptions = { border: { top: { style:'thin' } } };
    this.setCell('B' + rowNumber, '', borderOptions);
    this.setCell('C' + rowNumber, '', borderOptions);
    this.setCell('D' + rowNumber, '', borderOptions);
    this.setCell('E' + rowNumber, '', borderOptions);
    this.setCell('F' + rowNumber, '', borderOptions);
    this.setCell('G' + rowNumber, '', borderOptions);
    this.setCell('H' + rowNumber, '', borderOptions);
    this.setCell('I' + rowNumber, '', borderOptions);

    var rowNumber     = startingRow + 25;
    this.setCell('B' + rowNumber, '', borderOptions);
    this.setCell('C' + rowNumber, '', borderOptions);
    this.setCell('D' + rowNumber, '', borderOptions);
    this.setCell('E' + rowNumber, '', borderOptions);
    this.setCell('F' + rowNumber, '', borderOptions);
    this.setCell('G' + rowNumber, '', borderOptions);
    this.setCell('H' + rowNumber, '', borderOptions);
    this.setCell('I' + rowNumber, '', borderOptions);
  },

  updateLineItemsFooter: function(rowNumber) {
    var borderOptions = { border: { top: { style:'thin' } } };
    this.setCell('B' + rowNumber, '', borderOptions);
    this.setCell('C' + rowNumber, '', borderOptions);
    this.setCell('D' + rowNumber, '', borderOptions);
    this.setCell('E' + rowNumber, '', borderOptions);
    this.setCell('F' + rowNumber, '', borderOptions);
    this.setCell('G' + rowNumber, '', borderOptions);
    this.setCell('H' + rowNumber, '', borderOptions);
    this.setCell('I' + rowNumber, '', borderOptions);

    var rowNumber     = rowNumber + 3;
    this.setCell('B' + rowNumber, '', borderOptions);
    this.setCell('C' + rowNumber, '', borderOptions);
    this.setCell('D' + rowNumber, '', borderOptions);
    this.setCell('E' + rowNumber, '', borderOptions);
    this.setCell('F' + rowNumber, '', borderOptions);
    this.setCell('G' + rowNumber, '', borderOptions);
    this.setCell('H' + rowNumber, '', borderOptions);
    this.setCell('I' + rowNumber, '', borderOptions);
  },

  updateFooter: function(invoice, startingRow) {
    if (invoice.subTotal) {
      var rowNumber = startingRow + 26 + 20 + 1;
      this.updateLineItemsFooter(rowNumber);
      this.setCell('E' + (rowNumber + 1), 'សរុបទំព័រទី១​');
      this.setCell('G' + (rowNumber + 1), 'Sub Total Page No.  1:  ');
      this.setCell('I' + (rowNumber + 1), invoice.subTotal, this.rightAlignmentOptions());

      return rowNumber + 3;
    } else {
      var rowNumber = startingRow + 26 + 20;
      this.setCell('F' + (rowNumber + 1), 'សរុប', this.rightAlignmentOptions());
      this.setCell('G' + (rowNumber + 1), 'Total  :', this.rightAlignmentOptions());
      this.setCell('I' + (rowNumber + 1), invoice.total, this.rightAlignmentOptions());
      this.setCell('F' + (rowNumber + 2), 'អាករលើតម្លែបន្ថែម ១០%', this.rightAlignmentOptions());
      this.setCell('G' + (rowNumber + 2), 'VAT 10%:', this.rightAlignmentOptions());
      this.setCell('I' + (rowNumber + 2), invoice.vat, this.rightAlignmentOptions());

      this.updateLineItemsFooter(rowNumber + 4);
      this.setCell('F' + (rowNumber + 5), 'សរុបរួម', this.rightAlignmentOptions());
      this.setCell('G' + (rowNumber + 5), 'GRAND TOTAL:', this.rightAlignmentOptions());
      this.setCell('I' + (rowNumber + 5), invoice.grandTotal, this.rightAlignmentOptions());

      // terms
      this.setCell('B' + (rowNumber + 8), 'ការទូទាត់អាចធ្វើឡើងតាមមូលឡបទានប័ត្រ');
      this.setCell('B' + (rowNumber + 9), 'ឬ តាមរយ:ការទូទាត់ទៅធនាគារ');
      this.setCell('B' + (rowNumber + 10), 'Payment can be made by Cheque');
      this.setCell('B' + (rowNumber + 11), 'or Bank Transfer payable to :');
      this.setCell('B' + (rowNumber + 13), 'Account Name: ' + invoice.account.name);
      this.setCell('B' + (rowNumber + 14), 'Account No  : ' + invoice.account.number);
      this.setCell('B' + (rowNumber + 15), 'Bank Name   : ' + invoice.account.bank_name);
      this.setCell('B' + (rowNumber + 16), 'SWIFT CODE: ' + invoice.account.swift_code);

      this.setCell('I' + (rowNumber + 8), 'ហត្ថលេខា');
      this.setCell('I' + (rowNumber + 9), 'Authorized Signature');
      this.setCell('I' + (rowNumber + 13), '-------------------------------');

      return rowNumber + 17;
    }
  },

  // lineItems, max is 20
  updateLineItems: function(invoice, startingRow) {
    var self      = this;

    // header
    this.updateLineItemsHeader(startingRow);

    // body
    var rowNumber = startingRow + 26;
    invoice.lineItems.forEach(function(lineItem, index) {
      var hash = { line: 'B', code: 'C', description: 'D', spot: 'G', amount: 'I' };

      var lineCell        = 'B' + rowNumber;
      var codeItemCell    = 'C' + rowNumber;
      var descriptionCell = 'D' + rowNumber;
      var spotCell        = 'G' + rowNumber;
      var amountCell      = 'I' + rowNumber;

      self.setCell(lineCell, lineItem.line);
      self.setCell(codeItemCell, lineItem.codeItem);
      self.setCell(descriptionCell, lineItem.description);
      self.setCell(spotCell, lineItem.spot, self.rightAlignmentOptions());
      self.setCell(amountCell, lineItem.amount, self.rightAlignmentOptions());

      rowNumber++;
    });
  },

  setCell: function(cell, value, options) {
    this.worksheet.getCell(cell).font  = { size: 9 };
    this.worksheet.getCell(cell).value = value;

    if (options && options.alignment) {
      this.worksheet.getCell(cell).alignment = options.alignment;
    }

    if (options && options.border) {
      this.worksheet.getCell(cell).border = options.border;
    }
  },

  rightAlignmentOptions: function() {
    return { alignment: { horizontal: 'right' } };
  }
};

module.exports = function(workbook) {
  return new ExcelUpdater(workbook);
}