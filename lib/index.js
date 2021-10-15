"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    Object.defineProperty(o, k2, { enumerable: true, get: function() { return m[k]; } });
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var Excel = __importStar(require("exceljs"));
var workbook = new Excel.Workbook();
var flat_1 = require("flat");
var maxDepth = 0;
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
function setExcelHeader(flattenJson) {
    var headerInformation = getHeaderInformation(flattenJson);
    for (var header in headerInformation) {
        mergeCell(header, headerInformation[header]);
    }
    // console.log(sheet);
}
function getHeaderInformation(flattenJson) {
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
        var splittedArray = data.split(delimiter);
        if (maxDepth < splittedArray.length) {
            maxDepth = splittedArray.length + 1;
        }
    });
}
function customDelimiter(delimiter) {
    delimiter = delimiter;
}
module.exports = {

    generateExcel: function (sheetConfigurations) {
        if (!Array.isArray(sheetConfigurations)) {
            sheetConfigurations = [sheetConfigurations];
        }
        sheetConfigurations.forEach(function (sheetConfig) {
            cellTracker = {};
            sheet = addWorkSheet(sheetConfig);
            var data = Array.isArray(sheetConfig.data) ? sheetConfig.data : [sheetConfig.data];
            var flattenJson = (0, flat_1.flatten)(getSampleJson(data), {
                delimiter: delimiter
            });
            findMaxDepth(flattenJson);
            setExcelHeader(flattenJson);
            sheet.columns = Object.keys(flattenJson).map(function (jsonKey) { return ({
                key: jsonKey,
            }); });
            data.forEach(function (jsonData) {
                sheet.addRow((0, flat_1.flatten)(jsonData));
            });
        });
        return workbook;
    }
};
