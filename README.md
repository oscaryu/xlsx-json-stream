# xlsx-json-stream

Minimize memory usage by processing each row as soon as possible, instead of trying to convert everything into an array of JSON objects.

This does not have write functionality.

## Quick start

npm install xlsx-json-stream

````
const ExcelStreamer = require('xlsx-json-stream');


main = async () => {
    await ExcelStreamer.readToJson('Spreadsheet.xlsx', (err, data) => {
        console.log(data.row, data.rowCtr, data.dataLength, data.headerLength, data.sheetName);
        // ToDo: add your transform and load steps for each Excel row here
    });
    console.log("** Done **");
}

main()
````

## You can also override some parameters:

- MAX_EMPTY_SEQUENTIAL_ROWS = 2;
- DEFAULT_VALUE = null;
- HEADER_ROW = 0;
- DATA_START_ROW = 1;
- INCLUDE_EMPTY = false;
- SHEETS = [];
- ASYNC_BATCH_SIZE = 1;
- READ_OPTIONS = { dateFormat: 'm/d/yy;@' }; // Default date format is mm/dd/yyyy



````
const ExcelStreamer = require('xlsx-json-stream');

let options = {
  MAX_EMPTY_SEQUENTIAL_ROWS : 2,
  DEFAULT_VALUE:'',
  HEADER_ROW:0,
  DATA_START_ROW:1,
  INCLUDE_EMPTY:false,
  SHEETS:['USA','MEX'],
  ASYNC_BATCH_SIZE:1,
  READ_OPTIONS : {dateFormat: 'm/d/yy;@' }; // default format is mm/dd/yyyy
}

main = async () => {
    await ExcelStreamer.readToJson(filename, options, (err, data) => {
        console.log(data.row, data.rowCtr, data.sheetName, data.dataLength, data.headerLength);
        // ToDo: add your transform and load steps for each Excel row here
    });
    console.log("** Done **");
}

main()
````

## BREAKING CHANGE
We are now returning the standard `(error, data)`, instead of returning multiple parameters: `obj, rowCtr, sheetName, headerLength, dataLength`.  And access them this way:
 - data.row
 - data.rowCtr
 - data.sheetName
 - data.headerLength 
 - data.dataLength


## Exceptions

UnhandledPromiseRejectionWarning: Error: Sheet "MEX" not found in the *.xlsx file. Available sheets: "USA" (#1), "CAN" (#2).