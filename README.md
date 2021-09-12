# xlsx-json-stream

Minimize memory usage by processing each row as soon as possible, instead of trying to convert everything into an array of JSON objects.

This does not have write functionality.

## Quick start

npm install xlsx-json-stream

````
const ExcelStreamer = require('xlsx-json-stream');


main = async () => {
    await ExcelStreamer.readToJson('Spreadsheet.xlsx', (obj, rowCtr, sheetName, headerLength, dataLength) => {
        console.log(obj, rowCtr, dataLength, headerLength, obj, sheetName);
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


````
const ExcelStreamer = require('xlsx-json-stream');

let options = {
  MAX_EMPTY_SEQUENTIAL_ROWS : 2,
  DEFAULT_VALUE:'',
  HEADER_ROW:0,
  DATA_START_ROW:1,
  INCLUDE_EMPTY:false,
  SHEETS:['USA','MEX'],
  ASYNC_BATCH_SIZE:1
}

main = async () => {
    await ExcelStreamer.readToJson(filename, options, (obj, rowCtr, sheetName, headerLength, dataLength) => {
        console.log(obj, rowCtr, sheetName, dataLength, headerLength);
        // ToDo: add your transform and load steps for each Excel row here
    });
    console.log("** Done **");
}

main()
````

## Exceptions

UnhandledPromiseRejectionWarning: Error: Sheet "MEX" not found in the *.xlsx file. Available sheets: "USA" (#1), "CAN" (#2).