/*jshint esversion: 8 */
/*jshint node: true */

const fs = require("fs");
const path = require('path');
const XLSX = require("xlsx");
const XLSXs = require("xlsx-style");

const fieldsToReference = require("./objectFields.json");

const args = process.argv.slice(2);
const fileWriteName = path.basename(args[0], '.json') + '.xlsx';

const funcs = {
    nameByUID: nameByUID,
    expandName: expandName,
    arrayJoin: arrayJoin
};

// Define fields to reference for each metadata object type:

var meta = {};

main();

function main() {

    // Load metadata file
    const metadata = JSON.parse(fs.readFileSync(args[0]));



    // Initialize reference object and metametadata object
    let reference = {};
    meta = metametadata(metadata);
    let regex = new RegExp(/:(?<func>\w*)/);


    // Iterate through metadata and create reference
    for (let objType in metadata) {
        console.log(objType + ' ' + metadata[objType].length);

        let fields = Array.from(fieldsToReference['default']);

        fields.splice(fields.length, 0, ...(fieldsToReference[objType] ? fieldsToReference[objType] : []));

        if (!(metadata[objType] && metadata[objType].length > 0)) continue;

        let aoa = reference[objType] = [];
        let headers = fields.map((val => val.replace(/:.+/, '')));
        aoa.push([...headers]);

        for (let obj of metadata[objType]) {
            let a = [];

            for (let field of fields) {
                let thing = { "obj": obj, "path": null, "val": null };
                let func = null;

                let path = field.split(':');
                func = path.splice(1).toString();
                thing.path = path.toString();


                // If path is split with ".", resolve path to get value
                thing.path = thing.path.split('.');
                if (thing.path.length > 1) {
                    thing = (resolveValueByPath(thing));
                } else {
                    thing.val = (thing.obj[thing.path[0]] ? thing.obj[thing.path[0]] : '');
                }

                // Is val an Array?
                if (Array.isArray(thing.val)) {

                    thing.val = unpackArray(thing.val, thing.path, func);
                } else {
                    if (func) {
                        try {
                            thing.val = funcs[func](thing.val);

                        } catch (error) {
                            console.error(error);
                            console.log(`Func: ${func}`);
                        }
                    }
                }
                a.push(thing.val);
            }
            aoa.push(a);
        }
        aoa = pruneColumns(aoa);

    } // objType loop end

    // Initialize and build excel workbook
    let wrkBook = createWorkbook();
    for (objType in reference) {
        let aoa = reference[objType];
        appendWorksheet(sheetFromArray(aoa, true), wrkBook, objType);
    }
    saveWorkbook(wrkBook, fileWriteName);
};

function pruneColumns(aoa) {
    let content = aoa.slice(1);
    let isEmpty = content.map((arr) => {
        return arr.map((val) => (val == '' || val == undefined || val == null))
    })

    // Flip isEmpty to create arrays from columns
    flipIsEmpty = transpose(isEmpty);

    for (let emptyIndex = flipIsEmpty.length -1; emptyIndex >= 0; emptyIndex--) {
        aoa.forEach(arr => {
            if (!flipIsEmpty[emptyIndex].includes(false)) {
                arr.splice(emptyIndex, 1);
            }
        });
    }

    return aoa;
}

function resolveValueByPath(thing) {
    if (typeof thing.path === 'string') thing.path = thing.path.split('.');
    if (typeof thing.obj === 'string') {
        thing.val = thing.obj;
        return thing;
    }
    let what = thing.obj;
    if (Array.isArray(what)) {

        thing.val = thing.obj;
        return thing;
    }
    //console.log(path);
    //console.log(obj);
    try {
        let thePath = thing.path.slice(1);
        // thing.path = thing.path.slice(1);
        thing.obj = thing.obj[thing.path[0]];
        thing.path = thePath;

        // Sometimes we try to access properties not present in objects, e.g. categoryOptionCombos is in some, but not all, categoryCombos.
        // In such a case we return an empty string.
        if (thing.obj == null || thing.obj == undefined) {
            thing.val = '';
            thing.obj = '';
        }
        return resolveValueByPath(thing);
    } catch (error) {
        console.error(error);
        console.log(`Object: ${obj}`);
        console.log(`Path: ${path}`);
    }
}

function unpackArray(arr, path, func) {
    // Arrays in metadata usually contain objects. We usually want the "id" property.
    let newArr = arr.map(element => {
        if (typeof element === 'object') {
            let val;
            try {
                if (func && path) {
                    return funcs[func](element[path]);
                } else {
                    return element.id;
                }
            } catch (error) {
                throw error;
            }
        } else {
            try {
                if (func) {
                    return funcs[func](element);
                } else {
                    return element.id;
                }
            } catch (error) {
                throw error;
            }
        }
    });

    return newArr.join('; ');
}

/*
 * 
 */
function metametadata(metadata) {
    let res = {};
    for (objType in metadata) {
        if (metadata[objType] && metadata[objType].length > 0) {

            for (obj of metadata[objType]) {

                if (obj && obj.id && obj.name) {
                    res[obj.id] = { 'name': obj.name }
                }
            }
        }
    }
    // DashboardItems
    if (metadata.dashboards && metadata.dashboards.length > 0) {
        for (dashboard of metadata.dashboards) {
            if (dashboard.dashboardItems && dashboard.dashboardItems.length > 0) {
                for (di of dashboard.dashboardItems) {
                    try {
                        let type = di.type;
                        let uid;
                        type = type.toLowerCase();
                        if (di[type] && di[type].id) {
                            uid = di[type].id;
                        } else if (di.visualization && di.visualization.id) {
                            uid = di.visualization.id;
                        }
                        if (uid) {
                            res[di.id] = { 'name': res[uid].name }
                        }
                    } catch (error) {
                        throw error;
                    }
                }
            }
        }
    }

    return res;
}

function nameByUID(uid) {
    if (meta[uid]) {
        return meta[uid].name;
    };
    return uid;
}

function expandName(uid) {
    if (meta[uid]) {
        return `${uid} - ${meta[uid].name}`;
    };
    return uid;
}

function arrayJoin(arr) {
    console.log(arr);
    return arr;
}

function sheetFromArray(aoa, header) {
    var sheet = XLSX.utils.aoa_to_sheet(aoa);
    var range = XLSX.utils.decode_range(sheet["!ref"]);
    let colWidths = [];

    // Prettify the output a little, adding alternating background color to rows, ...
    for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            let cell = sheet[XLSXs.utils.encode_cell({ c: C, r: R })];

            if (cell == undefined) {
                //console.log('cell is undefined');
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

function transpose(aoa) {
    return aoa[0].map((col, index) => aoa.map(row => row[index]));
}
