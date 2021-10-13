const flatten = require("flat");
const excel = require("exceljs");
// const { max } = require("lodash");

let workSheetName;
let workbook = new excel.Workbook();
const json = [
    {
        study: {
            science: {
                bio: {
                    pharmacy: "bandana",
                    mbbs: {
                        general: "roshan",
                        md: "sanjay",
                    },
                },
                math: {
                    pureMath: "rukesh",
                    engineering: {
                        computer: {
                            hardware: "Aungush",
                            software: "nikesh",
                        },
                        civil: "seena",
                        mechanical: "santosh",
                    },
                },
            },
            management: {
                bba: "pratik",
                bbs: "jeena",
            },
        },
    },
    {
        study: {
            science: {
                bio: {
                    pharmacy: "rajani",
                    mbbs: {
                        general: "haris",
                        md: "shreetika",
                    },
                },
                math: {
                    pureMath: "prijal",
                    engineering: {
                        computer: {
                            hardware: "samina",
                            software: "anish",
                        },
                        civil: "rasil",
                        mechanical: "amit",
                    },
                },
            },
            management: {
                bba: "anjeela",
                bbs: "sushmin",
            },
        },
    },
];
var sheet = workbook.addWorksheet("sheet", {
    headerFooter: { firstHeader: "Hello Exceljs", firstFooter: "Hello World" },
});

let flattenJson = flatten(getSampleJson(json));
let maxDepth = 0;
let headerInformation = {};
let cellTracker = {};

function getSampleJson(json) {
    if (typeof json == "object" && Array.isArray(json)) {
        return json[0];
    }
    return json;
}

function findMaxDepth(flattenJson) {
    Object.keys(flattenJson).forEach((data) => {
        let splittedArray = data.split(".");
        // console.log(splittedArray);
        if (maxDepth < splittedArray.length) {
            maxDepth = splittedArray.length + 1;
        }
    });
}

findMaxDepth(flattenJson);
// console.log(flattenJson, "flatten json is", maxDepth);

function getHeaderInformation(flattenJson) {
    for (const data in flattenJson) {
        let splittedArray = data.split(".");
        let lastHeaderKey = splittedArray[splittedArray.length - 1];
        let rowSpan = 0;
        for (const headerKey of splittedArray) {
            let rowNumber = splittedArray.indexOf(headerKey) + 1;
            if (
                !headerInformation?.[headerKey] ||
                (headerInformation?.[headerKey] &&
                    headerInformation[headerKey].rowNumber != rowNumber)
            ) {
                // console.log(lastHeaderKey, headerKey, "apple apple");
                if (lastHeaderKey == headerKey) {
                    rowSpan = maxDepth - rowNumber - 1;
                }
                headerInformation[headerKey] = {
                    colSpan: 0,
                    rowSpan,
                    rowNumber,
                };
            } else {
                headerInformation[headerKey]["colSpan"] += 1;
            }
            // console.log(headerInformation);
        }
    }
    return headerInformation;
}

function setExcelHeader() {
    let headerInformation = getHeaderInformation(flattenJson);
    for (const header in headerInformation) {
        mergeCell(header, headerInformation[header]);
    }
    // console.log(sheet);
}

function mergeCell(header, { colSpan, rowNumber, rowSpan }) {
    // console.log("header", header, colSpan, rowNumber, rowSpan);
    let startRow = rowNumber;
    let startColumn = getColumnCell(rowNumber, colSpan, rowSpan);
    let endRow = startRow + rowSpan;
    let endColumn = startColumn + colSpan;
    console.log(
        "startRow ->",
        startRow,
        "startColumn ->",
        startColumn,
        "endRow ->",
        endRow,
        "endColumn ->",
        endColumn,
        "header ->",
        header
    );
    sheet.mergeCells(startRow, startColumn, endRow, endColumn);
    const row = sheet.getRow(rowNumber);
    const cell = row.getCell(startColumn);
    cell.value = header;
    // console.log(cell);
    // mergedCell.value = header;
}

function getColumnCell(rowNumber, colSpan, rowSpan) {
    // console.log(rowSpan, "row span is");
    let startColumn = cellTracker?.[rowNumber] ? cellTracker[rowNumber] + 1 : 1;
    for (let i = rowNumber; i <= rowNumber + rowSpan; i++) {
        cellTracker[i] = startColumn + colSpan;
    }
    // console.log(cellTracker);
    return startColumn;
}

async function generateExcel() {
    setExcelHeader();
    sheet.columns = Object.keys(flattenJson).map((jsonKey) => ({
        key: jsonKey,
    }));
    json.forEach((jsonData) => {
        sheet.addRow(flatten(jsonData));
    });
    await workbook.xlsx.writeFile("banana.xlsx");
}

generateExcel();
// let testRow = sheet.getRow(1);
// console.log(testRow.values);
// console.log(getHeaderInformation(flattenJson));

// console.log(headerInformation, "asfasdf");
// // return;
// console.log("header key is ", flattenJson);
// // if(headerInformation?.)
// // console.log(maxDepth);
// return;

// console.log(flatten(json));
