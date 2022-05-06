/*jshint esversion: 8 */
/*jshint node: true */

const fs = require("fs");
const path = require('path');
const XLSX = require("xlsx");
const XLSXs = require("xlsx-style");

const args = process.argv.slice(2);
const fileWriteName = path.basename(args[0], '.json') + '.xlsx';

const funcs = {
    nameByUID: nameByUID
};

// Define fields to reference for each metadata object type:
const fieldsToReference = JSON.parse(fs.readFileSync('objectFields.json'));

main();

function main() {

    // Load metadata file
    const metadata = JSON.parse(fs.readFileSync(args[0]));



    // Initialize reference object and metametadata object
    let reference = {};
    let meta = metametadata(metadata);
    let regex = new RegExp(/:(?<func>\w*)/);


    // Iterate through metadata and create reference
    for (let objType in metadata) {
        //console.log(objType + ' ' + metadata[objType].length);
        //let fields = (fieldsToReference[objType] ? fieldsToReference[objType] : fieldsToReference['default']);
        let fields = fieldsToReference['default'];
        fields.splice(fields.length, 0, ...(fieldsToReference[objType] ? fieldsToReference[objType] : []));

        if (metadata[objType] && metadata[objType].length > 0) {
            /*
            reference[objType] = [];
            let aoa = reference[objType];
            */
            let aoa = reference[objType] = [];
            aoa.push([...fields]);

            for (let obj of metadata[objType]) {
                let a = [];

                for (let field of fields) {
                    let val;

                    let path = field.split(':');
                    let func = path.splice(1).toString();
                    path = path.toString();

                    /*
                    if (match = field.match(regex)) {
                        console.log('\nfunc:');
                        console.log(match.groups.func);
                        func = match.groups.func;
                        console.log(match);
                    };
                    */

                    console.log(func);
                    // If path has . do something else
                    path = path.split('.');
                    if (path.length > 1) {
                        val = (resolveValueByPath(obj, path));
                    } else {
                        val = (obj[path[0]] ? obj[path[0]] : '');
                    }

                    // Is val an Array?
                    if (Array.isArray(val)) {
                        let vals = [];
                        for (let element of val) {
                            try {
                                vals.push = funcs[func](element, meta);
                                
                            } catch (error) {
                                console.error(error);
                                console.log(`Func: ${func}`);
                            }
                        }
                        val = vals.join(';');
                    } else {

                    if (func) {
                        try {
                            val = funcs[func](val, meta);
                            
                        } catch (error) {
                            console.error(error);
                            console.log(`Func: ${func}`);
                        }
                    }
                }
                    a.push(val);
                }
                // TODO Prune undefined columns from a
                aoa.push(a);
            }
        }
    }
    // Initialize and build excel workbook
    let wrkBook = createWorkbook();
    for (objType in reference) {
        let aoa = reference[objType];
        appendWorksheet(sheetFromArray(aoa, true), wrkBook, objType);
    }
    saveWorkbook(wrkBook, fileWriteName);
};


function sheetFromArray(aoa, header) {
    var sheet = XLSX.utils.aoa_to_sheet(aoa);
    var range = XLSX.utils.decode_range(sheet["!ref"]);
    let colWidths = [];

    // Prettyfy the output a little, adding alternating background color to rows, ...
    for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            let cell = sheet[XLSXs.utils.encode_cell({ c: C, r: R })];

            if (cell == undefined) {
                console.log('cell is undefined');
                continue;
            }

            if (header && R == 0) {
                cell.s = { font: { bold: true } };
                cell.s.fill = { fgColor: { rgb: "a5a5e2" } };

            } else if (R == 0) {
                cell.s = (cell.s ? cell.s : {});
                cell.s.fill = { fgColor: { rgb: "d5d5f2" } };
            }

            if (R % 2 == 0 && R > 0) {
                cell.s = (cell.s ? cell.s : {});
                cell.s.fill = { fgColor: { rgb: "d5d5f2" } };
            } else if (R > 0) {
                cell.s = (cell.s ? cell.s : {});
                cell.s.fill = { fgColor: { rgb: "e4e4f6" } };
            }

            // ... determine longest string in column, ...
            if (!colWidths[C]) colWidths[C] = 1;
            if (cell.v) {
                colWidths[C] = (cell.v.length > colWidths[C]) ? cell.v.length + 2 : colWidths[C];
            };
        }
    }
    // ... and setting column widths.
    sheet["!cols"] = (sheet["!cols"]) ? sheet["!cols"] : [];
    for (let col = 0; col < colWidths.length; col++) {
        sheet["!cols"].push({ wch: colWidths[col] });
    }

    return sheet;
}


function createWorkbook() {
    return XLSX.utils.book_new();
}

function appendWorksheet(sheet, book, name) {
    XLSX.utils.book_append_sheet(book, sheet, name);
}

function saveWorkbook(book, file) {
    XLSXs.writeFile(book, file);
    console.log("✔ Reference list saved");
}

function resolveValueByPath(obj, path) {
    if (typeof path === 'string') path = path.split('.');
    if (typeof obj === 'string') return obj;
    if (Array.isArray(obj)) return obj;
    //console.log(path);
    //console.log(obj);
    try {
        return resolveValueByPath(obj[path[0]], path.slice(1));
    } catch (error) {
        console.error(error);
        console.log(obj)
        console.log(path);
    }
}

function metametadata(metadata) {
    let res = {};
    for (objType in metadata) {
        if (metadata[objType] && metadata[objType].length > 0) {

            for (obj of metadata[objType]) {

                if (obj && obj.id && obj.name) {
                    res[obj.id] = { "name": obj.name }
                }
            }
        }
    }
    return res;
}

function nameByUID(uid, meta) {
    // Check if uid is an Array
    if (meta[uid]) {
        return meta[uid].name;
    };
    return undefined;
}
/*
function nameByUID(uid, meta) {
    if (meta[uid]) {
        return meta[uid].name;
    };
    return undefined;
}
*/
