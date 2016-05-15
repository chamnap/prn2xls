#!/usr/bin/env node

var program = require('commander');
var path    = require('path');
var prn2xls = require('./index.js');

program
  .version('0.1.0')
  .option('-p, --pdf',                      'pdf format')
  .option('-d, --directory [value]',        'destination directory')
  .option('-v, --verbose [value]',          'verbose mode')
  .parse(process.argv);

var prnFile = program.args[0];

var extName = path.extname(prnFile);
if (extName !== '.PRN' && extName !== '.prn') {
  throw new Error("The source file must be .PRN file.");
}

var method = program.pdf ? 'pdf' : 'xls';
prn2xls[method](prnFile, program.directory, function(error, newFilePath) {
  if (!error) {
    console.log('File: `' + newFilePath + '` is Done!');
  }
});