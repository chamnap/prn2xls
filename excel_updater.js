var ExcelUpdater = function(workbook) {
  this.workbook = workbook;
  this.worksheet = this.workbook.getWorksheet(1);
};

ExcelUpdater.prototype = {
  update: function(invoices) {
    var invoice = invoices[1];

    this.updateHeader(invoice, 0);
    this.updateLineItems(invoice.lineItems, 0);
  },

  updateHeader: function(invoice, startIndex) {
    this.setCell('B' + (1 + startIndex), 'វិក្កយប័ត្រអាករ');
    this.setCell('B' + (2 + startIndex), 'VAT INVOICE');
    this.setCell('B' + (3 + startIndex), 'លេខអត្តសញ្ញាណកម្ម អតប');
    this.setCell('B' + (4 + startIndex), 'VAT TIN: ' + invoice.tin);
    this.setCell('B' + (6 + startIndex), 'ឈ្មោះអតិថិជន');
    // this.setCell('C' + (6 + startIndex), invoice.customer.kh_name);
    this.setCell('B' + (7 + startIndex), 'Bill To:           ');
    // this.setCell('C' + (7 + startIndex), invoice.customer.en_name);
    // this.setCell('C' + (8 + startIndex), invoice.customer.kh_address1);
    // this.setCell('C' + (9 + startIndex), invoice.customer.en_address1);
    // this.setCell('C' + (10 + startIndex), invoice.customer.kh_address2);
    // this.setCell('C' + (11 + startIndex), invoice.customer.en_address2);

    this.setCell('G' + (3 + startIndex), 'ទំព័រ');
    this.setCell('G' + (4 + startIndex), 'Page No.:        ' + invoice.pageNumber);
    this.setCell('G' + (6 + startIndex), 'លេខរៀងវិក្កយប័ត្រ');
    this.setCell('G' + (7 + startIndex), 'Invoice  No.');
    this.setCell('H' + (7 + startIndex), ':');
    this.setCell('I' + (7 + startIndex), invoice.invoice.number);

    this.setCell('G' + (8 + startIndex), 'កាលបរិចេ្ឆទ');
    this.setCell('G' + (9 + startIndex), 'Invoice  Dt.');
    this.setCell('H' + (9 + startIndex), ':');
    this.setCell('I' + (9 + startIndex), invoice.invoice.date);

    this.setCell('G' + (10 + startIndex), 'លេខរៀងការបញ្ជាទិញ');
    this.setCell('G' + (11 + startIndex), 'S.O.No.     ');
    this.setCell('H' + (11 + startIndex), ':');
    this.setCell('I' + (11 + startIndex), invoice.soNumber);

    this.setCell('B' + (13 + startIndex), 'លេខកូតអតិថិជន');
    this.setCell('B' + (14 + startIndex), 'C.Code :');
    this.setCell('C' + (14 + startIndex), invoice.cCode);
    this.setCell('B' + (15 + startIndex), 'លេខអត្តសញ្ញាណកម្ម អតប');
    this.setCell('C' + (15 + startIndex), ' VAT    :');

    this.setCell('G' + (13 + startIndex), 'កិច្ចសន្យាលេខ');
    this.setCell('G' + (14 + startIndex), 'Contract No.');
    this.setCell('H' + (14 + startIndex), ':');
    this.setCell('I' + (14 + startIndex), invoice.contract.number);
    this.setCell('G' + (15 + startIndex), 'កាលបរិចេ្ឆទនៃកិច្ចសន្យា');
    this.setCell('G' + (16 + startIndex), 'Contract Dt.');
    this.setCell('H' + (16 + startIndex), ':');
    this.setCell('I' + (16 + startIndex), invoice.contract.date);

    this.setCell('B' + (17 + startIndex), 'ទំនាក់ទំនង');
    this.setCell('B' + (18 + startIndex), 'Contact:');
    this.setCell('C' + (18 + startIndex), invoice.contact);

    this.setCell('G' + (17 + startIndex), 'រយ:ពេលទូទាត់');
    this.setCell('I' + (17 + startIndex), '៧ថ្ងៃបន្ទាប់ពីទទួលបានវិក្កយប័ត្រ');
    this.setCell('G' + (18 + startIndex), 'Terms:   7 days after the');
    this.setCell('G' + (19 + startIndex), 'date of receiving invoice');

    this.setCell('B' + (22 + startIndex), 'លេខរៀង');
    this.setCell('C' + (22 + startIndex), 'កូតទំនិញេ/សេវា');
    this.setCell('D' + (22 + startIndex), 'បរិយាយមុខទំនិញ/សេវា');
    this.setCell('G' + (22 + startIndex), 'បរិមាណ');
    this.setCell('I' + (22 + startIndex), 'ចំនួនទឹកប្រាក់');
    this.setCell('B' + (23 + startIndex), 'Line');
    this.setCell('C' + (23 + startIndex), 'Item Code');
    this.setCell('D' + (23 + startIndex), 'Item Description');
    this.setCell('G' + (23 + startIndex), 'No.of Spot');
    this.setCell('I' + (23 + startIndex), 'US$ Amount');
  },

  updateLineItemsHeader: function(startIndex) {
    var rowNumber     = startIndex + 21;
    var borderOptions = { border: { top: { style:'thin' } } };
    this.setCell('B' + rowNumber, '', borderOptions);
    this.setCell('C' + rowNumber, '', borderOptions);
    this.setCell('D' + rowNumber, '', borderOptions);
    this.setCell('E' + rowNumber, '', borderOptions);
    this.setCell('F' + rowNumber, '', borderOptions);
    this.setCell('G' + rowNumber, '', borderOptions);
    this.setCell('H' + rowNumber, '', borderOptions);
    this.setCell('I' + rowNumber, '', borderOptions);

    var rowNumber     = startIndex + 25;
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

  // lineItems, max is 20
  updateLineItems: function(lineItems, startIndex) {
    var self      = this;

    // header
    this.updateLineItemsHeader(startIndex);

    // body
    var rowNumber = startIndex + 26;
    lineItems.forEach(function(lineItem) {
      var hash = { line: 'B', code: 'C', description: 'D', spot: 'G', amount: 'I' };

      var lineCell        = 'B' + rowNumber;
      var codeItemCell    = 'C' + rowNumber;
      var descriptionCell = 'D' + rowNumber;
      var spotCell        = 'G' + rowNumber;
      var amountCell      = 'I' + rowNumber;
      var options         = { alignment: { horizontal: 'right' } };

      self.setCell(lineCell, lineItem.line);
      self.setCell(codeItemCell, lineItem.codeItem);
      self.setCell(descriptionCell, lineItem.description);
      self.setCell(spotCell, lineItem.spot, options);
      self.setCell(amountCell, lineItem.amount, options);

      rowNumber++;
    });

    // footer
    this.updateLineItemsFooter(rowNumber + 1);
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
  }
};

module.exports = function(workbook) {
  return new ExcelUpdater(workbook);
}