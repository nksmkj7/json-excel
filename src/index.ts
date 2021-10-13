import * as Excel from 'exceljs';
const workbook = new Excel.Workbook();
import * as flatten from 'flat';

let maxDepth: number = 0;
let flattenJson: object;
let sheet: Excel.Worksheet;
let headerInformation: headerInformationType= {};
let cellTracker: cellTrackerType = {};
let delimiter: string = ".";

interface sheet {
    title: string,
    data: object| object[],
    options?: object
}

interface headerInformationType {
    [index: string]: {
        colSpan: number,
        rowSpan: number,
        rowNumber: number
    }
}

interface cellTrackerType{
    [index:number]: number
}



function  getSampleJson(json: object[]) {
  if (typeof json == 'object' && Array.isArray(json)) {
    return json[0];
  }
  return json;
}

type workbookType = InstanceType <typeof Excel.Workbook>

function addWorkSheet (workSheet: sheet) {
    let worksheetTittle = workSheet.title;
    let worksheetOptions = workSheet?.options || {};
    return workbook.addWorksheet(worksheetTittle, worksheetOptions);
};

function setExcelHeader() {
  let headerInformation = getHeaderInformation();
  for (const header in headerInformation) {
    mergeCell(header, headerInformation[header]);
  }
  // console.log(sheet);
}

function getHeaderInformation() {
  for (const data in flattenJson) {
    let splittedArray = data.split(delimiter);
    let lastHeaderKey = splittedArray[splittedArray.length - 1];
    let rowSpan = 0;
    for (const headerKey of splittedArray) {
      let rowNumber = splittedArray.indexOf(headerKey) + 1;
      if (
        !headerInformation?.[headerKey] ||
        (headerInformation?.[headerKey] && headerInformation[headerKey].rowNumber != rowNumber)
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
        headerInformation[headerKey]['colSpan'] += 1;
      }
      // console.log(headerInformation);
    }
  }
  return headerInformation;
}

function mergeCell(header:string, { colSpan, rowNumber, rowSpan }: {colSpan:number, rowNumber:number, rowSpan:number}) {
  let startRow = rowNumber;
  let startColumn = getColumnCell(rowNumber, colSpan, rowSpan);
  let endRow = startRow + rowSpan;
  let endColumn = startColumn + colSpan;
  sheet.mergeCells(startRow, startColumn, endRow, endColumn);
  const row = sheet.getRow(rowNumber);
  const cell = row.getCell(startColumn);
  cell.value = header;
}

function getColumnCell(rowNumber:number, colSpan:number, rowSpan:number) {
  let startColumn = cellTracker?.[rowNumber] ? cellTracker[rowNumber] + 1 : 1;
  for (let i = rowNumber; i <= rowNumber + rowSpan; i++) {
    cellTracker[i] = startColumn + colSpan;
  }
  return startColumn;
}


function findMaxDepth(flattenJson: object) {
  Object.keys(flattenJson).forEach((data) => {
    let splittedArray = data.split('.');
    if (maxDepth < splittedArray.length) {
      maxDepth = splittedArray.length + 1;
    }
  });
}


module.exports = {
    setDelimiter: (delimiter:string) => delimiter = delimiter,
  generateExcel: (sheetConfigurations: sheet[]) => {
    if (!Array.isArray(sheetConfigurations)) {
      sheetConfigurations = [sheetConfigurations];
    }
    sheetConfigurations.forEach(sheetConfig => {
        sheet = addWorkSheet(sheetConfig);
        let data = Array.isArray(sheetConfig.data) ? sheetConfig.data : [sheetConfig.data];
        flattenJson = flatten(getSampleJson(data), {
          delimiter
        });
        findMaxDepth(flattenJson);
        setExcelHeader()
        sheet.columns = Object.keys(flattenJson).map((jsonKey) => ({
            key: jsonKey,
        }));
        data.forEach((jsonData) => {
            sheet.addRow(flatten(jsonData));
        });
    });
    return workbook;
  }
}


