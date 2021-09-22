const ExcelStreamer = require('xlsx-json-stream');

const filename = '/home/test/Spreadsheet_20210910.xlsx';

main = async () => {
    options = {SHEETS:['US', 'MEX'], dateFormat: 'm/d/yy;@' }; // // mm/dd/yyyy
    await ExcelStreamer.readToJson(filename, options, (err, data) => {
        console.log(data.row, data.rowCtr, data.sheetName, data.dataLength, data.headerLength);
    });
    console.log("** Done **");
}

main()
