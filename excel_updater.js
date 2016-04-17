var _ = require('lodash');

var ExcelUpdater = function(workbook) {
  this.workbook = workbook;

  this.rightAlignmentOptions = { alignment: { horizontal: 'right' } };
  this.khmerFontOptions = { font: { name: 'Khmer OS', size: 10 } };
  this.khmerBigFontOptions = { font: { name: 'Khmer OS', size: 12 } };

  this.leftTopAlignmentKhmerFont  = { alignment: { vertical: 'top', horizontal: 'left', wrapText: true }, font: { name: 'Khmer OS', size: 10 } };
  this.rightTopAlignmentKhmerFont = { alignment: { vertical: 'top', horizontal: 'right', wrapText: true }, font: { name: 'Khmer OS', size: 10 } };
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
        { width: 3.6 },   // A
        { width: 10.6 },  // B
        { width: 14.2 },  // C
        { width: 9.2 },   // D
        { width: 9.2 },   // E
        { width: 9.2 },   // F
        { width: 9.2 },   // G
        { width: 14.2 },  // H
        { width: 18 },    // I
        { width: 2.4 },   // J
        { width: 18 }     // K
      ];

      self.updateHeader(worksheet, invoice, 0);

      self.updateLineItems(worksheet, invoice, 0);

      self.updateFooter(worksheet, invoice, 0);
    });
  },

  updateHeader: function(worksheet, invoice, startingRow) {

    // left
    this.setCell(worksheet, 'B' + (1 + startingRow), 'វិក្កយប័ត្រអាករ', this.khmerBigFontOptions);
    this.setCell(worksheet, 'B' + (2 + startingRow), 'VAT INVOICE');
    this.setCell(worksheet, 'B' + (3 + startingRow), 'លេខអត្តសញ្ញាណកម្ម អតប/VAT TIN: ' + invoice.tin, this.khmerFontOptions);
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

    this.setCell(worksheet, 'B' + (12 + startingRow), 'C.Code :');
    this.setCell(worksheet, 'C' + (12 + startingRow), invoice.cCode);

    this.setCell(worksheet, 'B' + (13 + startingRow), 'VAT    :');

    this.setCell(worksheet, 'B' + (14 + startingRow), 'Contact:');
    this.setCell(worksheet, 'C' + (14 + startingRow), invoice.contact);

    // right
    this.setCell(worksheet, 'I' + (6 + startingRow), 'ទំព័រ/Page No.', this.rightAlignmentKhmerFont);
    this.setCell(worksheet, 'J' + (6 + startingRow), ':');
    this.setCell(worksheet, 'K' + (6 + startingRow), invoice.pageNumber);

    this.setCell(worksheet, 'I' + (7 + startingRow), 'លេខរៀងវិក្កយប័ត្រ/Invoice  No.', this.rightAlignmentKhmerFont);
    this.setCell(worksheet, 'J' + (7 + startingRow), ':');
    this.setCell(worksheet, 'K' + (7 + startingRow), invoice.invoice.number);

    this.setCell(worksheet, 'I' + (8 + startingRow), 'កាលបរិចេ្ឆទ/Invoice  Dt.', this.rightAlignmentKhmerFont);
    this.setCell(worksheet, 'J' + (8 + startingRow), ':');
    this.setCell(worksheet, 'K' + (8 + startingRow), invoice.invoice.date);

    this.setCell(worksheet, 'I' + (9 + startingRow), 'លេខរៀងការបញ្ជាទិញ/S.O.No.', this.rightAlignmentKhmerFont);
    this.setCell(worksheet, 'J' + (9 + startingRow), ':');
    this.setCell(worksheet, 'K' + (9 + startingRow), invoice.soNumber);

    this.setCell(worksheet, 'I' + (10 + startingRow), 'កិច្ចសន្យាលេខ/Contract No.', this.rightAlignmentKhmerFont);
    this.setCell(worksheet, 'J' + (10 + startingRow), ':');
    this.setCell(worksheet, 'K' + (10 + startingRow), invoice.contract.number);

    this.setCell(worksheet, 'I' + (11 + startingRow), 'កាលបរិចេ្ឆទនៃកិច្ចសន្យា/Contract Dt.', this.rightAlignmentKhmerFont);
    this.setCell(worksheet, 'J' + (11 + startingRow), ':');
    this.setCell(worksheet, 'K' + (11 + startingRow), invoice.contract.date);

    this.setCell(worksheet, 'I' + (12 + startingRow), 'រយ:ពេលទូទាត់', this.rightTopAlignmentKhmerFont);
    this.setCell(worksheet, 'J' + (12 + startingRow), ':', this.leftTopAlignmentKhmerFont);
    this.setCell(worksheet, 'K' + (12 + startingRow), '៧ថ្ងៃបន្ទាប់ពីទទួលបានវិក្កយប័ត្រ', this.leftTopAlignmentKhmerFont);

    this.setCell(worksheet, 'I' + (13 + startingRow), 'Terms', this.rightAlignmentKhmerFont);
    this.setCell(worksheet, 'J' + (13 + startingRow), ':');
    this.setCell(worksheet, 'K' + (13 + startingRow), '7 days after the');
    this.setCell(worksheet, 'K' + (14 + startingRow), 'date of receiving invoice');

    this.setCell(worksheet, 'B' + (17 + startingRow), 'លេខរៀង', this.khmerFontOptions);
    this.setCell(worksheet, 'C' + (17 + startingRow), 'កូតទំនិញេ/សេវា', this.khmerFontOptions);
    this.setCell(worksheet, 'D' + (17 + startingRow), 'បរិយាយមុខទំនិញ/សេវា', this.khmerFontOptions);
    this.setCell(worksheet, 'I' + (17 + startingRow), 'បរិមាណ', this.khmerFontOptions);
    this.setCell(worksheet, 'K' + (17 + startingRow), 'ចំនួនទឹកប្រាក់', this.khmerFontOptions);

    this.setCell(worksheet, 'B' + (18 + startingRow), 'Line');
    this.setCell(worksheet, 'C' + (18 + startingRow), 'Item Code');
    this.setCell(worksheet, 'D' + (18 + startingRow), 'Item Description');
    this.setCell(worksheet, 'I' + (18 + startingRow), 'No.of Spot');
    this.setCell(worksheet, 'K' + (18 + startingRow), 'US$ Amount');
  },

  updateLineItemsHeader: function(worksheet, startingRow) {
    var rowNumber     = startingRow + 16;
    var borderOptions = { border: { top: { style:'thin' } } };
    this.setCell(worksheet, 'B' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'C' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'D' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'E' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'F' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'G' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'H' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'I' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'J' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'K' + rowNumber, '', borderOptions);

    var rowNumber     = startingRow + 20;
    this.setCell(worksheet, 'B' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'C' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'D' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'E' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'F' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'G' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'H' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'I' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'J' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'K' + rowNumber, '', borderOptions);
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
    this.setCell(worksheet, 'J' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'K' + rowNumber, '', borderOptions);

    var rowNumber     = rowNumber + 3;
    this.setCell(worksheet, 'B' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'C' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'D' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'E' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'F' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'G' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'H' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'I' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'J' + rowNumber, '', borderOptions);
    this.setCell(worksheet, 'K' + rowNumber, '', borderOptions);
  },

  updateFooter: function(worksheet, invoice, startingRow) {
    if (invoice.subTotal) {
      var rowNumber = startingRow + 21 + 20 + 1;
      this.updateLineItemsFooter(worksheet, rowNumber);
      this.setCell(worksheet, 'H' + (rowNumber + 1), 'សរុបទំព័រទី' + invoice.khmerPageNumber, this.rightAlignmentKhmerFont);
      this.setCell(worksheet, 'I' + (rowNumber + 1), 'Sub Total Page No.  ' + invoice.pageNumber + ':  ', this.rightAlignmentOptions);
      this.setCell(worksheet, 'K' + (rowNumber + 1), invoice.subTotal, this.rightAlignmentOptions);

      return rowNumber + 3;
    } else {
      var rowNumber = startingRow + 21 + 20;
      this.setCell(worksheet, 'H' + (rowNumber + 1), 'សរុប', this.rightAlignmentKhmerFont);
      this.setCell(worksheet, 'I' + (rowNumber + 1), 'Total  :', this.rightAlignmentOptions);
      this.setCell(worksheet, 'K' + (rowNumber + 1), invoice.total, this.rightAlignmentOptions);

      this.setCell(worksheet, 'H' + (rowNumber + 2), 'អាករលើតម្លែបន្ថែម ១០%', this.rightAlignmentKhmerFont);
      this.setCell(worksheet, 'I' + (rowNumber + 2), 'VAT 10%:', this.rightAlignmentOptions);
      this.setCell(worksheet, 'K' + (rowNumber + 2), invoice.vat10, this.rightAlignmentOptions);

      this.updateLineItemsFooter(worksheet, rowNumber + 4);
      this.setCell(worksheet, 'H' + (rowNumber + 5), 'សរុបរួម', this.rightAlignmentKhmerFont);
      this.setCell(worksheet, 'I' + (rowNumber + 5), 'GRAND TOTAL:', this.rightAlignmentOptions);
      this.setCell(worksheet, 'K' + (rowNumber + 5), invoice.grandTotal, this.rightAlignmentOptions);

      // terms
      this.setCell(worksheet, 'B' + (rowNumber + 8), 'ការទូទាត់អាចធ្វើឡើងតាមមូលឡបទានប័ត្រ', this.khmerFontOptions);
      this.setCell(worksheet, 'B' + (rowNumber + 9), 'ឬ តាមរយ:ការទូទាត់ទៅធនាគារ', this.khmerFontOptions);
      this.setCell(worksheet, 'B' + (rowNumber + 10), 'Payment can be made by Cheque');
      this.setCell(worksheet, 'B' + (rowNumber + 11), 'or Bank Transfer payable to :');
      this.setCell(worksheet, 'B' + (rowNumber + 13), 'Account Name: ' + invoice.account.name);
      this.setCell(worksheet, 'B' + (rowNumber + 14), 'Account No  : ' + invoice.account.number);
      this.setCell(worksheet, 'B' + (rowNumber + 15), 'Bank Name   : ' + invoice.account.bank_name);
      this.setCell(worksheet, 'B' + (rowNumber + 16), 'SWIFT CODE: ' + invoice.account.swift_code);

      this.setCell(worksheet, 'K' + (rowNumber + 8), 'ហត្ថលេខា', this.khmerFontOptions);
      this.setCell(worksheet, 'K' + (rowNumber + 9), 'Authorized Signature');
      this.setCell(worksheet, 'K' + (rowNumber + 13), '-------------------------------');

      return rowNumber + 17;
    }
  },

  // lineItems, max is 20
  updateLineItems: function(worksheet, invoice, startingRow) {
    var self      = this;

    // header
    this.updateLineItemsHeader(worksheet, startingRow);

    // body
    var rowNumber = startingRow + 21;
    invoice.lineItems.forEach(function(lineItem, index) {
      var hash = { line: 'B', code: 'C', description: 'D', spot: 'G', amount: 'I' };

      var lineCell        = 'B' + rowNumber;
      var codeItemCell    = 'C' + rowNumber;
      var descriptionCell = 'D' + rowNumber;
      var spotCell        = 'I' + rowNumber;
      var amountCell      = 'K' + rowNumber;

      self.setCell(worksheet, lineCell, lineItem.line);
      self.setCell(worksheet, codeItemCell, lineItem.codeItem);
      self.setCell(worksheet, descriptionCell, lineItem.description);
      self.setCell(worksheet, spotCell, lineItem.spot, self.rightAlignmentOptions);
      self.setCell(worksheet, amountCell, lineItem.amount, self.rightAlignmentOptions);

      rowNumber++;
    });
  },

  setCell: function(worksheet, cell, value, options) {
    worksheet.getCell(cell).font  = { size: 10 };
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