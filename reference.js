/*jshint esversion: 8 */
/*jshint node: true */

const fs = require("fs");
const XLSX = require("xlsx");
const XLSXs = require("xlsx-style");

const args = process.argv.slice(2);

// Define fields to reference for each metadata object type:
const fieldsToReference = JSON.parse(fs.readFileSync('objectFields.json'));

main();

function main(){

    // Load metadata file
    const metadata = JSON.parse(fs.readFileSync(args[0]));



    // Initialize reference object
    let reference = {};

    // Iterate through metadata and create reference
    for (let objType in metadata) {
        console.log(objType);
        console.log(metadata[objType].length);
        let fields = (fieldsToReference[objType] ? fieldsToReference[objType] : fieldsToReference['default']);

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
                    // If field has . do something else
                    field = field.split('.');
                    if (field.length > 1) {
                        console.log('!Deep!');
                        a.push(resolveValueByPath(obj, field));
                    } else {
                        a.push(obj[field[0]] ? obj[field[0]] : '');
                    }
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
    saveWorkbook(wrkBook, 'poop.xlsx');
};


function sheetFromArray(aoa, header) {
	var sheet = XLSX.utils.aoa_to_sheet(aoa);
	var range = XLSX.utils.decode_range(sheet["!ref"]);
	let colWidths = [];

    // Prettyfy the output a little, adding alternating background color to rows, ...
	for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
			let cell = sheet[XLSXs.utils.encode_cell({c:C, r:R})];

			if (cell == undefined) {
				console.log('cell is undefined');
				continue;
			}

			if (header && R == 0) {
				cell.s = {font: {bold: true}};
				cell.s.fill = {fgColor: {rgb: "a5a5e2"}};

			} else if (R == 0) {
				cell.s = (cell.s ? cell.s : {});
				cell.s.fill = {fgColor: {rgb: "d5d5f2"}};
			}

			if (R % 2 == 0 && R > 0) {
				cell.s = (cell.s ? cell.s : {});
				cell.s.fill = {fgColor: {rgb: "d5d5f2"}};
			} else if ( R > 0 ) {
				cell.s = (cell.s ? cell.s : {});
                cell.s.fill = {fgColor: {rgb: "e4e4f6"}};
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
		sheet["!cols"].push( {wch: colWidths[col]});
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

function resolveValueByPath(obj, path){
    if (typeof path === 'string') path = path.split('.');
    if (typeof obj === 'string') return obj;
    //console.log(path);
    //console.log(obj);
    return resolveValueByPath(obj[path[0]], path.slice(1));
}
