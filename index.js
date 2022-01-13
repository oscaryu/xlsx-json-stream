const xlsxFile = require('read-excel-file/node');
const fs = require('fs');
const XLSX = require('xlsx');
const { parseExcelDate } = require('read-excel-file');

let MAX_EMPTY_SEQUENTIAL_ROWS = 2;
let DEFAULT_VALUE='';
let HEADER_ROW=1;
let DATA_START_ROW=2;
let INCLUDE_EMPTY=false;
let SHEETS=[];
let ASYNC_BATCH_SIZE=1;
let READ_OPTIONS = { } // dateFormat: 'm/d/yy;@' // mm/dd/yyyy
let IS_LOCAL_DATES = true;

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
            IS_LOCAL_DATES = options.IS_LOCAL_DATES ?? false;
        }
        let emptyRowCtr = 0;
        /* equivalent to `var wb = XLSX.readFile(filename);` */
        const buf = fs.readFileSync(filename);
        const wb = XLSX.read(buf, {type:'buffer'});

        if (!SHEETS || !SHEETS.length) {
            let sheets = await xlsxFile(filename, { getSheets: true })
            SHEETS = sheets.map(s=>s.name);
            console.log("No sheets", SHEETS)
        }
        
        for(let i=0;i<SHEETS.length;i++) {
            const sheetName = SHEETS[i];
            const worksheet = wb.Sheets[sheetName];
            // console.log("Sheet#"+i, sheetName);
            await readSheet(worksheet, sheetName, READ_OPTIONS, IS_LOCAL_DATES, cb);
        }
        resolve();
    });
}

module.exports = { readToJson };

function readSheet(worksheet, sheetName, options, IS_LOCAL_DATES, cb) {
    return new Promise(async (resolve, reject) => {
        if (sheetName && sheetName.length) {
            options['sheet'] = sheetName;
        }
        try {
            const headerList = getHeaderList(worksheet, HEADER_ROW)
            const headers = Object.keys(headerList).map(x=>x.replace(/[0-9]/gmi,""))
            console.log('headerList', headers)
    
            let emptyRowCtr = 0;
            let promises = [];
            for(let i=DATA_START_ROW; i<Number.MAX_SAFE_INTEGER; i++) {
                const row = {};
                let isEmptyRow = true;
                headers.forEach(h=>{
                    const hdr = headerList[h+''+HEADER_ROW].w;
                    const cell_value = getCellValue(h+i, worksheet); // 'AD3'
                    // console.log(cell_value, hdr);
                    if (cell_value) {
                        if (typeof cell_value === 'object') {
                            isEmptyRow = false;
                            // console.log(i, 'is not empty')
                            row[hdr] = cell_value.v;
                            console.log('cell_value', cell_value); // { t: 'n', v: 44373, w: '6/26/21' }
                            if (cell_value.t === 'n' && (occurrenceCount(cell_value.w, '/') === 2 || occurrenceCount(cell_value.w, '.') === 2 || occurrenceCount(cell_value.w, ':') === 2)) {
                                try {
                                    const parsed = parseExcelDate(cell_value.v).toISOString();
                                    console.log(i, hdr, parsed, cell_value.w);
                                    if (IS_LOCAL_DATES) {
                                        row[hdr] = parsed.replace('Z', '');
                                    } else {
                                        row[hdr] = parsed;
                                    }
                                    // console.log(hdr,":", col,typeof col, 'Is a Date Field:', DATE_FIELD_LIST && DATE_FIELD_LIST.length && DATE_FIELD_LIST.includes(hdr), '; Parsed:', obj[hdr])
                                } catch (ex) {
                                    // console.log(hdr,":", col, 'Unable to parse', ex.toString())
                                    row[hdr] = cell_value;
                                }
                            } else {
                                row[hdr] = cell_value.v;
                            }
                        } else {
                            row[hdr] = cell_value;
                        }
                    } else if (INCLUDE_EMPTY) {
                        row[hdr] = DEFAULT_VALUE;
                    }
                })
                if (isEmptyRow) {
                    emptyRowCtr = emptyRowCtr+1;
                    // console.log(i, 'is empty.  emptyCount:', emptyCount)
                } else {
                    // console.log(i, 'is not empty')
                    emptyRowCtr = 0;
                    //row['row'] = i;
                }
                if (row !== {} && Object.keys(row).length) {
                    // console.log("Parsed Obj",i+":", row);
                    // await cb(null, obj);
                    promises.push(cb(null, {row, rowCtr:i, sheetName, dataLength:headers.length, headerLength:row.length}));
                    if (promises.length>= ASYNC_BATCH_SIZE) {
                        await Promise.all(promises);
                        promises = [];
                    }
                }
                if (emptyRowCtr>= MAX_EMPTY_SEQUENTIAL_ROWS) {
                    // console.log("BREAKING!!!", i, emptyRowCtr, MAX_EMPTY_SEQUENTIAL_ROWS)
                    resolve();
                    break;
                }
            }
            if (promises.length>= 0) {
                await Promise.all(promises);
            }
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

function getHeaderList(worksheet, lineNbr) {
    let keys = Object.keys(worksheet);
    let headerList = {};
    let headers = keys.filter(x=>x.replace(/[A-Z]/gmi,"")==lineNbr);
    headers.forEach(x=>{
        const desired_cell = worksheet[x];
        headerList[x] = desired_cell
    })
    return headerList;
}

function getCellValue(address_of_cell, worksheet) {
    const cell = worksheet[address_of_cell];
    let desired_value = undefined;
    if (cell) {
        if (cell.w && cell.v && cell.w != cell.v) {
            desired_value = cell;
        } else {
            desired_value = cell.w;
        }
    }
    return desired_value;
}

function occurrenceCount(str, char)
{
    let ctr = 0;
    for (let i = 0; i < str.length; i++) {
        if (str.charAt(i) == char) {
            ctr = ctr + 1;
        }
    }
    return ctr;
}
