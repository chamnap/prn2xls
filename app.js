#!/usr/bin/env node

var program = require('commander');
var path    = require('path');
var prn2xls = require('./index.js');

program
  .version('0.1.0')
  .option('-p, --path [value]',             'full path of prn file, eg: PI1105.PRN')
  .option('-v, --verbose [value]',          'verbose mode')
  .parse(process.argv);

var extName = path.extname(program.path);
if (extName != '.PRN') {
  throw new Error("The source file must be .PRN file.");
}

prn2xls.process(program.path);