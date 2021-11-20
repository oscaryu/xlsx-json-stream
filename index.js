const xlsxFile = require('read-excel-file/node');
const { parseExcelDate } = require('read-excel-file');

let MAX_EMPTY_SEQUENTIAL_ROWS = 2;
let DEFAULT_VALUE='';
let HEADER_ROW=0;
let DATA_START_ROW=1;
let INCLUDE_EMPTY=false;
let SHEETS=[];
let ASYNC_BATCH_SIZE=1;
let READ_OPTIONS = { } // dateFormat: 'm/d/yy;@' // mm/dd/yyyy

readToJson = (filename, options, cb) =>{
    return new Promise(async (resolve, reject) => {
        if (!cb && typeof options == 'function') {
            cb = options;
            console.log("No options", cb)
        } else {
            console.log("Options got passed", options)
            MAX_EMPTY_SEQUENTIAL_ROWS = options.MAX_EMPTY_SEQUENTIAL_ROWS ?? MAX_EMPTY_SEQUENTIAL_ROWS;
            DEFAULT_VALUE = options.DEFAULT_VALUE ?? DEFAULT_VALUE;
            HEADER_ROW = options.HEADER_ROW ?? HEADER_ROW;
            DATA_START_ROW = options.DATA_START_ROW ?? DATA_START_ROW;
            INCLUDE_EMPTY = options.INCLUDE_EMPTY ?? INCLUDE_EMPTY;
            ASYNC_BATCH_SIZE = options.ASYNC_BATCH_SIZE ?? ASYNC_BATCH_SIZE;
            SHEETS = options.SHEETS ?? SHEETS;
            READ_OPTIONS = options.READ_OPTIONS ?? READ_OPTIONS;
        }
        let emptyRowCtr = 0;
        if (!SHEETS || !SHEETS.length) {
            let sheets = await xlsxFile(filename, { getSheets: true })
            SHEETS = sheets.map(s=>s.name);
            console.log("No sheets", SHEETS)
        }
        for(let i=0;i<SHEETS.length;i++) {
            const sheetName = SHEETS[i];
            // console.log("Sheet#"+i, sheetName);
            await readSheet(filename, sheetName, emptyRowCtr, READ_OPTIONS, options.DATE_FIELD_LIST, cb);
        }
        resolve();
    });
}

module.exports = { readToJson };

function readSheet(filename, sheetName, emptyRowCtr, options, DATE_FIELD_LIST, cb) {
    return new Promise(async (resolve, reject) => {
        if (sheetName && sheetName.length) {
            options['sheet'] = sheetName;
        }
        try {
            await xlsxFile(filename, options).then(async (rows) => {
                const headers = rows[HEADER_ROW];
                let promises = [];
                for (let i = 1; i < rows.length; i++) {
                    if (i >= DATA_START_ROW) {
                        const row = rows[i];
                        let isEmpty = true;
                        const obj = {};
                        for (let c = 0; c < row.length; c++) {
                            const col = row[c];
                            const hdr = headers[c];
                            if (col) {
                                isEmpty = false;
                                if(typeof col == 'number' && DATE_FIELD_LIST && DATE_FIELD_LIST.length && DATE_FIELD_LIST.includes(hdr)) {
                                    try {
                                        obj[hdr] = parseExcelDate(col).toISOString();
                                        // console.log(hdr,":", col,typeof col, 'Is a Date Field:', DATE_FIELD_LIST && DATE_FIELD_LIST.length && DATE_FIELD_LIST.includes(hdr), '; Parsed:', obj[hdr])
                                    } catch (ex) {
                                        // console.log(hdr,":", col, 'Unable to parse', ex.toString())
                                        obj[hdr] = col;
                                    }
                                } else {
                                    obj[hdr] = col;
                                }
                            } else if (DEFAULT_VALUE != null) {
                                obj[hdr] = DEFAULT_VALUE;
                            }
                        }
                        if (isEmpty) {
                            emptyRowCtr += 1;
                            // console.log(i, isEmpty, row.length, headers.length, "Empty row detected");
                            if (emptyRowCtr >= MAX_EMPTY_SEQUENTIAL_ROWS) {
                                // console.log("Done");
                                break;
                            }
                            if (INCLUDE_EMPTY) {
                                promises.push(cb(null, {row:obj, rowCtr:i, sheetName, dataLength:headers.length, headerLength:row.length}));
                                if (promises.length>= ASYNC_BATCH_SIZE) {
                                    await Promise.all(promises);
                                    promises = [];
                                }
                            }
                        } else {
                            emptyRowCtr = 0;
                            // console.log(i, isEmpty, row.length, headers.length, obj);
                            promises.push(cb(null, {row:obj, rowCtr:i, sheetName, dataLength:headers.length, headerLength:row.length}));
                            if (promises.length>= ASYNC_BATCH_SIZE) {
                                await Promise.all(promises);
                                promises = [];
                            }
                    }
                    }
                }
                if (promises.length>= 0) {
                    await Promise.all(promises);
                    promises = [];
                }
                resolve();
            });
        } catch (ex) {
            const str = ex.toString();
            if (str.indexOf('Error: Sheet') > -1 && str.indexOf('not found in the') > -1) {
                console.log("Ignoring missing sheet");
            } else {
                console.log('********', ex.toString());
            }
        }
        resolve();
    });
}

