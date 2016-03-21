#!/usr/bin/env node

var program = require('commander');
var prn2xls = require('./index.js');

program
  .version('0.1.0')
  .option('-p, --path [value]',             'full path of prn file')
  .option('-v, --verbose [value]',          'verbose mode')
  .parse(process.argv);

prn2xls.process(program.path);