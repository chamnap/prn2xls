## PRN2XLS

An npm module to convert PRN file into XLS file. You can use it stand alone with the command `prn2xls`.

    $ prn2xls -p PI1105.PRN -d /home/chamnapchhorn

This module has tests with Mocha. Run `npm test` and make sure you have a solid connection.

```
var prn2xls = require('prn2xls');
prn2xls.convert(program.path, program.directory, function(error, newFilePath) {
  if (!error) {
    console.log('File: `' + newFilePath + '` is Done!');
  }
});
```