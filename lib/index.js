"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var Excel = require("exceljs");
var workbook = new Excel.Workbook();
var flatten = require("flat");
var maxDepth = 0;
var flattenJson;
var sheet;
var headerInformation = {};
var cellTracker = {};
var delimiter = ".";
function getSampleJson(json) {
    if (typeof json == 'object' && Array.isArray(json)) {
        return json[0];
    }
    return json;
}
function addWorkSheet(workSheet) {
    var worksheetTittle = workSheet.title;
    var worksheetOptions = (workSheet === null || workSheet === void 0 ? void 0 : workSheet.options) || {};
    return workbook.addWorksheet(worksheetTittle, worksheetOptions);
}
;
function setExcelHeader() {
    var headerInformation = getHeaderInformation();
    for (var header in headerInformation) {
        mergeCell(header, headerInformation[header]);
    }
    // console.log(sheet);
}
function getHeaderInformation() {
    for (var data in flattenJson) {
        var splittedArray = data.split(delimiter);
        var lastHeaderKey = splittedArray[splittedArray.length - 1];
        var rowSpan = 0;
        for (var _i = 0, splittedArray_1 = splittedArray; _i < splittedArray_1.length; _i++) {
            var headerKey = splittedArray_1[_i];
            var rowNumber = splittedArray.indexOf(headerKey) + 1;
            if (!(headerInformation === null || headerInformation === void 0 ? void 0 : headerInformation[headerKey]) ||
                ((headerInformation === null || headerInformation === void 0 ? void 0 : headerInformation[headerKey]) && headerInformation[headerKey].rowNumber != rowNumber)) {
                // console.log(lastHeaderKey, headerKey, "apple apple");
                if (lastHeaderKey == headerKey) {
                    rowSpan = maxDepth - rowNumber - 1;
                }
                headerInformation[headerKey] = {
                    colSpan: 0,
                    rowSpan: rowSpan,
                    rowNumber: rowNumber,
                };
            }
            else {
                headerInformation[headerKey]['colSpan'] += 1;
            }
            // console.log(headerInformation);
        }
    }
    return headerInformation;
}
function mergeCell(header, _a) {
    var colSpan = _a.colSpan, rowNumber = _a.rowNumber, rowSpan = _a.rowSpan;
    var startRow = rowNumber;
    var startColumn = getColumnCell(rowNumber, colSpan, rowSpan);
    var endRow = startRow + rowSpan;
    var endColumn = startColumn + colSpan;
    sheet.mergeCells(startRow, startColumn, endRow, endColumn);
    var row = sheet.getRow(rowNumber);
    var cell = row.getCell(startColumn);
    cell.value = header;
}
function getColumnCell(rowNumber, colSpan, rowSpan) {
    var startColumn = (cellTracker === null || cellTracker === void 0 ? void 0 : cellTracker[rowNumber]) ? cellTracker[rowNumber] + 1 : 1;
    for (var i = rowNumber; i <= rowNumber + rowSpan; i++) {
        cellTracker[i] = startColumn + colSpan;
    }
    return startColumn;
}
function findMaxDepth(flattenJson) {
    Object.keys(flattenJson).forEach(function (data) {
        var splittedArray = data.split('.');
        if (maxDepth < splittedArray.length) {
            maxDepth = splittedArray.length + 1;
        }
    });
}
module.exports = {
    add: function (a, b) { return a + b; },
    setDelimiter: function (delimiter) { return delimiter = delimiter; },
    generateExcel: function (sheetConfigurations) {
        if (!Array.isArray(sheetConfigurations)) {
            sheetConfigurations = [sheetConfigurations];
        }
        sheetConfigurations.forEach(function (sheetConfig) {
            sheet = addWorkSheet(sheetConfig);
            var data = Array.isArray(sheetConfig.data) ? sheetConfig.data : [sheetConfig.data];
            flattenJson = flatten(getSampleJson(data), {
                delimiter: delimiter
            });
            findMaxDepth(flattenJson);
            setExcelHeader();
            sheet.columns = Object.keys(flattenJson).map(function (jsonKey) { return ({
                key: jsonKey,
            }); });
            data.forEach(function (jsonData) {
                sheet.addRow(flatten(jsonData));
            });
        });
        return workbook;
    }
};
