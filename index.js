const xlsxFile = require('read-excel-file/node');

let MAX_EMPTY_SEQUENTIAL_ROWS = 2;
let DEFAULT_VALUE='';
let HEADER_ROW=0;
let DATA_START_ROW=1;
let INCLUDE_EMPTY=false;
let SHEETS=[];
let ASYNC_BATCH_SIZE=1;

readToJson = (filename, options, cb) =>{
    return new Promise(async (resolve, reject) => {
        if (!cb && typeof options == 'function') {
            cb = options;
            console.log("No options", cb)
        } else {
            console.log("Options got passed")
            MAX_EMPTY_SEQUENTIAL_ROWS = options.MAX_EMPTY_SEQUENTIAL_ROWS ?? MAX_EMPTY_SEQUENTIAL_ROWS;
            DEFAULT_VALUE = options.DEFAULT_VALUE ?? DEFAULT_VALUE;
            HEADER_ROW = options.HEADER_ROW ?? HEADER_ROW;
            DATA_START_ROW = options.DATA_START_ROW ?? DATA_START_ROW;
            INCLUDE_EMPTY = options.INCLUDE_EMPTY ?? INCLUDE_EMPTY;
            ASYNC_BATCH_SIZE = options.ASYNC_BATCH_SIZE ?? ASYNC_BATCH_SIZE;
            SHEETS = options.SHEETS ?? SHEETS;
        }
        let emptyRowCtr = 0;
        if (!SHEETS || !SHEETS.length) {
            let sheets = await xlsxFile(filename, { getSheets: true })
            SHEETS = sheets.map(s=>s.name);
            console.log("No sheets", SHEETS)
        }
        for(let i=0;i<SHEETS.length;i++) {
            const sheetName = SHEETS[i];
            console.log("Sheet#"+i, sheetName);
            await readSheet(filename, sheetName, emptyRowCtr, cb, resolve);
        }
        resolve();
    });
}

module.exports = { readToJson };

function readSheet(filename, sheetName, emptyRowCtr, cb, resolve) {
    return new Promise((resolve, reject) => {
        let options = { };
        if (sheetName && sheetName.length) {
            options['sheet'] = sheetName;
        }
        xlsxFile(filename, options).then(async (rows) => {
            const headers = rows[HEADER_ROW];
            let promises = [];
            for (let i = 1; i < rows.length; i++) {
                if (i >= DATA_START_ROW) {
                    const row = rows[i];
                    let isEmpty = true;
                    const obj = {};
                    for (let c = 0; c < row.length; c++) {
                        const col = row[c];
                        if (col && col.length) {
                            isEmpty = false;
                            obj[headers[c]] = col;
                        } else if (DEFAULT_VALUE != null) {
                            obj[headers[c]] = DEFAULT_VALUE;
                        }
                    }
                    if (isEmpty) {
                        emptyRowCtr += 1;
                        console.log(i, isEmpty, row.length, headers.length, "Empty row detected");
                        if (emptyRowCtr >= MAX_EMPTY_SEQUENTIAL_ROWS) {
                            // console.log("Done");
                            break;
                        }
                        if (INCLUDE_EMPTY) {
                            promises.push(cb(obj, i, sheetName, headers.length, row.length));
                            if (promises.length>= ASYNC_BATCH_SIZE) {
                                await Promise.all(promises);
                                promises = [];
                            }
                        }
                    } else {
                        emptyRowCtr = 0;
                        console.log(i, isEmpty, row.length, headers.length, obj);
                        promises.push(cb(obj, i, sheetName, headers.length, row.length));
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
    });
}

