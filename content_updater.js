var ContentUpdater = function(workbook) {
  this.workbook = workbook;
  this.worksheet = this.workbook.getWorksheet(1);
};

ContentUpdater.prototype = {
  update: function(parser) {
    var self = this;

    // headers
    this.updateCell('B4', 'VAT TIN: ' + parser.tin);
    this.updateCell('C7', parser.company.name);
    this.updateCell('C9', parser.company.address1);
    this.updateCell('C11', parser.company.address2);
    this.updateCell('B14', 'C.Code : ' + parser.cCode);
    this.updateCell('B16', 'VAT    : ' + parser.vat);
    this.updateCell('C18', parser.contact);
    this.updateCell('G4', 'Page No.:        ' + parser.pageNumber);
    this.updateCell('I7', parser.invoice.number);
    this.updateCell('I9', parser.invoice.date);
    this.updateCell('I11', parser.soNumber);
    this.updateCell('I14', parser.contract.number);
    this.updateCell('I16', parser.contract.date);

    // line items, max 20
    var rowNumber = 26;
    parser.lineItems.forEach(function(lineItem) {
      var hash = { line: 'B', code: 'C', description: 'D', spot: 'G', amount: 'I' };

      var lineItemCell    = 'B' + rowNumber;
      var codeItemCell    = 'C' + rowNumber;
      var descriptionCell = 'D' + rowNumber;
      var spotCell        = 'G' + rowNumber;
      var amountCell      = 'I' + rowNumber;
      var options         = { alignment: { horizontal: 'right' } };

      self.updateCell(lineItemCell, lineItem.lineItem);
      self.updateCell(codeItemCell, lineItem.codeItem);
      self.updateCell(descriptionCell, lineItem.description);
      self.updateCell(spotCell, lineItem.spot, options);
      self.updateCell(amountCell, lineItem.amount, options);

      rowNumber++;
    });

    // footer
    this.updateCell('H48', parser.total);
    this.updateCell('H50', parser.vat10);
    this.updateCell('H54', parser.grantTotal);
    this.updateCell('B62', 'Account Name: ' + parser.account.name);
    this.updateCell('B63', 'Account No  : ' + parser.account.number);
    this.updateCell('B64', 'Bank Name   : ' + parser.account.bank_name);
    this.updateCell('B65', 'SWIFT CODE: ' + parser.account.swift_code);
  },

  updateCell: function(cell, value, options) {
    this.worksheet.getCell(cell).value = value;

    if (options && options.alignment) {
      this.worksheet.getCell(cell).alignment = options.alignment;
    }
  }
};

module.exports = function(workbook) {
  return new ContentUpdater(workbook);
}