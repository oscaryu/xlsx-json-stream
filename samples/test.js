const ExcelStreamer = require('xlsx-json-stream');

const filename = '/home/test/Spreadsheet_20210910.xlsx';

main = async () => {
    const validSheetNames = ['US', 'MEX'];
    const DATE_FIELD_LIST = ['Scheduled Date','Actual Date'];
    let options = { SHEETS: validSheetNames, DATE_FIELD_LIST, DATA_START_ROW: 2, READ_OPTIONS: { dateFormat: 'm/d/yy h:mm;@' } }; // default format is mm/dd/yyyy
    await ExcelStreamer.readToJson(filename, options, (err, data) => {
        console.log(data.rowCtr, data.sheetName, data.dataLength, data.headerLength, JSON.stringify(data.row));
    });
    console.log("** Done **");
}

main()
