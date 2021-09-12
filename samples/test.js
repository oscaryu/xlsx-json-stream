const ExcelStreamer = require('xlsx-json-stream');

const filename = '/home/test/Spreadsheet_20210910.xlsx';

main = async () => {
    options = {SHEETS:['USA', 'MEX']}
    await ExcelStreamer.readToJson(filename, options, (obj, rowCtr, sheetName, headerLength, dataLength) => {
        console.log(obj, rowCtr, sheetName, dataLength, headerLength);
    });
    console.log("** Done **");
}

main()
