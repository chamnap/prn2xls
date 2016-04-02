var _ = require('lodash');

var ExcelUpdater = function(workbook) {
  this.workbook = workbook;

  this.rightAlignmentOptions = { alignment: { horizontal: 'right' } };
  this.khmerFontOptions = { font: { name: 'Khmer OS', size: 9 } };

  this.leftTopAlignmentKhmerFont  = { alignment: { vertical: 'top', horizontal: 'left', wrapText: true }, font: { name: 'Khmer OS', size: 9 } };
  this.rightAlignmentKhmerFont    = _.extend({}, this.rightAlignmentOptions, this.khmerFontOptions);
};

ExcelUpdater.prototype = {
  update: function(invoices) {
    var self = this;

    _.forEach(invoices, function(invoice, index) {
      if (index == 0) {
        var worksheet = self.workbook.worksheets[0];
      } else {
        var worksheet = self.workbook.addWorksheet('Sheet' + (index+1));
      }

      worksheet.columns = [
        { width: 3.6 },
        { width: 10.6 },
        { width: 14.2 },
        { width: 9.2 },
        { width: 9.3 },
        { width: 13.5 },
        { width: 12.8 },
        { width: 2.4 },
        { width: 16 }
      ];

      self.updateHeader(worksheet, invoice, 0);

      self.updateLineItems(worksheet, invoice, 0);

      self.updateFooter(worksheet, invoice, 0);
    });
  },

  updateHeader: function(worksheet, invoice, startingRow) {
    this.setCell(worksheet, 'B' + (1 + startingRow), 'វិក្កយប័ត្រអាករ', this.khmerFontOptions);
    this.setCell(worksheet, 'B' + (2 + startingRow), 'VAT INVOICE');
    this.setCell(worksheet, 'B' + (3 + startingRow), 'លេខអត្តសញ្ញាណកម្ម អតប', this.khmerFontOptions);
    this.setCell(worksheet, 'B' + (4 + startingRow), 'VAT TIN: ' + invoice.tin);
    this.setCell(worksheet, 'B' + (6 + startingRow), 'ឈ្មោះអតិថិជន:', this.khmerFontOptions);
    this.setCell(worksheet, 'C' + (6 + startingRow), invoice.customer.kh_name);
    this.setCell(worksheet, 'B' + (7 + startingRow), 'Bill To:');
    this.setCell(worksheet, 'C' + (7 + startingRow), invoice.customer.en_name);
    this.setCell(worksheet, 'B' + (8 + startingRow), 'អាស័យដ្ឋាន:', this.khmerFontOptions);
    this.setCell(worksheet, 'C' + (8 + startingRow), invoice.customer.kh_address1);
    this.setCell(worksheet, 'C' + (9 + startingRow), invoice.customer.kh_address2);
    this.setCell(worksheet, 'B' + (10 + startingRow), 'Address:');
    this.setCell(worksheet, 'C' + (10 + startingRow), invoice.customer.en_address1);
    this.setCell(worksheet, 'C' + (11 + startingRow), invoice.customer.en_address2);

    this.setCell(worksheet, 'G' + (3 + startingRow), 'ទំព័រ', this.khmerFontOptions);
    this.setCell(worksheet, 'G' + (4 + startingRow), 'Page No.:        ' + invoice.pageNumber);
    this.setCell(worksheet, 'G' + (6 + startingRow), 'លេខរៀងវិក្កយប័ត្រ', this.khmerFontOptions);
    this.setCell(worksheet, 'G' + (7 + startingRow), 'Invoice  No.');
    this.setCell(worksheet, 'H' + (7 + startingRow), ':');
    this.setCell(worksheet, 'I' + (7 + startingRow), invoice.invoice.number);

    this.setCell(worksheet, 'G' + (8 + startingRow), 'កាលបរិចេ្ឆទ', this.khmerFontOptions);
    this.setCell(worksheet, 'G' + (9 + startingRow), 'Invoice  Dt.');
    this.setCell(worksheet, 'H' + (9 + startingRow), ':');
    this.setCell(worksheet, 'I' + (9 + startingRow), invoice.invoice.date);

    this.setCell(worksheet, 'G' + (10 + startingRow), 'លេខរៀងការបញ្ជាទិញ', this.khmerFontOptions);
    this.setCell(worksheet, 'G' + (11 + startingRow), 'S.O.No.     ');
    this.setCell(worksheet, 'H' + (11 + startingRow), ':');
    this.setCell(worksheet, 'I' + (11 + startingRow), invoice.soNumber);

    this.setCell(worksheet, 'B' + (13 + startingRow), 'លេខកូតអតិថិជន', this.khmerFontOptions);
    this.setCell(worksheet, 'B' + (14 + startingRow), 'C.Code :');
    this.setCell(worksheet, 'C' + (14 + startingRow), invoice.cCode);
    this.setCell(worksheet, 'B' + (15 + startingRow), 'លេខអត្តសញ្ញាណកម្ម អតប', this.khmerFontOptions);
    this.setCell(worksheet, 'B' + (15 + startingRow), ' VAT    :');

    this.setCell(worksheet, 'G' + (13 + startingRow), 'កិច្ចសន្យាលេខ', this.khmerFontOptions);
    this.setCell(worksheet, 'G' + (14 + startingRow), 'Contract No.');
    this.setCell(worksheet, 'H' + (14 + startingRow), ':');
    this.setCell(worksheet, 'I' + (14 + startingRow), invoice.contract.number);
    this.setCell(worksheet, 'G' + (15 + startingRow), 'កាលបរិចេ្ឆទនៃកិច្ចសន្យា', this.khmerFontOptions);
    this.setCell(worksheet, 'G' + (16 + startingRow), 'Contract Dt.');
    this.setCell(worksheet, 'H' + (16 + startingRow), ':');
    this.setCell(worksheet, 'I' + (16 + startingRow), invoice.contract.date);

    this.setCell(worksheet, 'B' + (17 + startingRow), 'ទំនាក់ទំនង', this.khmerFontOptions);
    this.setCell(worksheet, 'B' + (18 + startingRow), 'Contact:');
    this.setCell(worksheet, 'C' + (18 + startingRow), invoice.contact);

    this.setCell(worksheet, 'G' + (17 + startingRow), 'រយ:ពេលទូទាត់', this.leftTopAlignmentKhmerFont);
    this.setCell(worksheet, 'I' + (17 + startingRow), '៧ថ្ងៃបន្ទាប់ពីទទួលបានវិក្កយប័ត្រ', this.leftTopAlignmentKhmerFont);
    this.setCell(worksheet, 'G' + (18 + startingRow), 'Terms:');
    this.setCell(worksheet, 'I' + (18 + startingRow), '7 days after the');
    this.setCell(worksheet, 'I' + (19 + startingRow), 'date of receiving invoice');

    this.setCell(worksheet, 'B' + (22 + startingRow), 'លេខរៀង', this.khmerFontOptions);
    this.setCell(worksheet, 'C' + (22 + startingRow), 'កូតទំនិញេ/សេវា', this.khmerFontOptions);
    this.setCell(worksheet, 'D' + (22 + startingRow), 'បរិយាយមុខទំនិញ/សេវា', this.khmerFontOptions);
    this.setCell(worksheet, 'G' + (22 + startingRow), 'បរិមាណ', this.khmerFontOptions);
    this.setCell(worksheet, 'I' + (22 + startingRow), 'ចំនួនទឹកប្រាក់', this.khmerFontOptions);
    this.setCell(worksheet, 'B' + (23 + startingRow), 'Line');
    this.setCell(worksheet, 'C' + (23 + startingRow), 'Item Code');
    this.setCell(worksheet, 'D' + (23 + startingRow), 'Item Description');
    this.setCell(worksheet, 'G' + (23 + startingRow), 'No.of Spot');
    this.setCell(worksheet, 'I' + (23 + startingRow), 'US$ Amount');
  },

  updateLineItemsHeader: function(worksheet, startingRow) {
    var rowNumber     = startingRow + 21;
    var borderOptions = { border: { top: { style:'thin' } } };
    this.setCell(worksheet, 'B' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'C' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'D' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'E' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'F' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'G' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'H' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'I' + rowNumber, '', borderOptions);

    var rowNumber     = startingRow + 25;
    this.setCell(worksheet, 'B' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'C' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'D' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'E' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'F' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'G' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'H' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'I' + rowNumber, '', borderOptions);
  },

  updateLineItemsFooter: function(worksheet, rowNumber) {
    var borderOptions = { border: { top: { style:'thin' } } };
    this.setCell(worksheet, 'B' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'C' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'D' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'E' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'F' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'G' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'H' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'I' + rowNumber, '', borderOptions);

    var rowNumber     = rowNumber + 3;
    this.setCell(worksheet, 'B' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'C' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'D' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'E' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'F' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'G' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'H' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'I' + rowNumber, '', borderOptions);
  },

  updateFooter: function(worksheet, invoice, startingRow) {
    if (invoice.subTotal) {
      var rowNumber = startingRow + 26 + 20 + 1;
      this.updateLineItemsFooter(worksheet, rowNumber);
      this.setCell(worksheet, 'E' + (rowNumber + 1), 'សរុបទំព័រទី១​', this.khmerFontOptions);
      this.setCell(worksheet, 'G' + (rowNumber + 1), 'Sub Total Page No.  1:  ');
      this.setCell(worksheet, 'I' + (rowNumber + 1), invoice.subTotal, this.rightAlignmentOptions);

      return rowNumber + 3;
    } else {
      var rowNumber = startingRow + 26 + 20;
      this.setCell(worksheet, 'F' + (rowNumber + 1), 'សរុប', this.rightAlignmentKhmerFont);
      this.setCell(worksheet, 'G' + (rowNumber + 1), 'Total  :', this.rightAlignmentOptions);
      this.setCell(worksheet, 'I' + (rowNumber + 1), invoice.total, this.rightAlignmentOptions);
      this.setCell(worksheet, 'F' + (rowNumber + 2), 'អាករលើតម្លែបន្ថែម ១០%', this.rightAlignmentKhmerFont);
      this.setCell(worksheet, 'G' + (rowNumber + 2), 'VAT 10%:', this.rightAlignmentOptions);
      this.setCell(worksheet, 'I' + (rowNumber + 2), invoice.vat10, this.rightAlignmentOptions);

      this.updateLineItemsFooter(worksheet, rowNumber + 4);
      this.setCell(worksheet, 'F' + (rowNumber + 5), 'សរុបរួម', this.rightAlignmentKhmerFont);
      this.setCell(worksheet, 'G' + (rowNumber + 5), 'GRAND TOTAL:', this.rightAlignmentOptions);
      this.setCell(worksheet, 'I' + (rowNumber + 5), invoice.grandTotal, this.rightAlignmentOptions);

      // terms
      this.setCell(worksheet, 'B' + (rowNumber + 8), 'ការទូទាត់អាចធ្វើឡើងតាមមូលឡបទានប័ត្រ', this.khmerFontOptions);
      this.setCell(worksheet, 'B' + (rowNumber + 9), 'ឬ តាមរយ:ការទូទាត់ទៅធនាគារ', this.khmerFontOptions);
      this.setCell(worksheet, 'B' + (rowNumber + 10), 'Payment can be made by Cheque');
      this.setCell(worksheet, 'B' + (rowNumber + 11), 'or Bank Transfer payable to :');
      this.setCell(worksheet, 'B' + (rowNumber + 13), 'Account Name: ' + invoice.account.name);
      this.setCell(worksheet, 'B' + (rowNumber + 14), 'Account No  : ' + invoice.account.number);
      this.setCell(worksheet, 'B' + (rowNumber + 15), 'Bank Name   : ' + invoice.account.bank_name);
      this.setCell(worksheet, 'B' + (rowNumber + 16), 'SWIFT CODE: ' + invoice.account.swift_code);

      this.setCell(worksheet, 'I' + (rowNumber + 8), 'ហត្ថលេខា', this.khmerFontOptions);
      this.setCell(worksheet, 'I' + (rowNumber + 9), 'Authorized Signature');
      this.setCell(worksheet, 'I' + (rowNumber + 13), '-------------------------------');

      return rowNumber + 17;
    }
  },

  // lineItems, max is 20
  updateLineItems: function(worksheet, invoice, startingRow) {
    var self      = this;

    // header
    this.updateLineItemsHeader(worksheet, startingRow);

    // body
    var rowNumber = startingRow + 26;
    invoice.lineItems.forEach(function(lineItem, index) {
      var hash = { line: 'B', code: 'C', description: 'D', spot: 'G', amount: 'I' };

      var lineCell        = 'B' + rowNumber;
      var codeItemCell    = 'C' + rowNumber;
      var descriptionCell = 'D' + rowNumber;
      var spotCell        = 'G' + rowNumber;
      var amountCell      = 'I' + rowNumber;

      self.setCell(worksheet, lineCell, lineItem.line);
      self.setCell(worksheet, codeItemCell, lineItem.codeItem);
      self.setCell(worksheet, descriptionCell, lineItem.description);
      self.setCell(worksheet, spotCell, lineItem.spot, self.rightAlignmentOptions);
      self.setCell(worksheet, amountCell, lineItem.amount, self.rightAlignmentOptions);

      rowNumber++;
    });
  },

  setCell: function(worksheet, cell, value, options) {
    worksheet.getCell(cell).font  = { size: 9 };
    worksheet.getCell(cell).value = value;
    worksheet.getCell(cell).height = 50;

    if (options && options.alignment) {
      worksheet.getCell(cell).alignment = options.alignment;
    }

    if (options && options.border) {
      worksheet.getCell(cell).border = options.border;
    }

    if (options && options.font) {
      worksheet.getCell(cell).font = options.font;
    }
  }
};

module.exports = function(workbook) {
  return new ExcelUpdater(workbook);
}