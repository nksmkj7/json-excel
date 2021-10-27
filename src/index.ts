import * as Excel from 'exceljs';
const workbook = new Excel.Workbook();
import { flatten } from 'flat';

let maxDepth: number = 0;
let sheet: Excel.Worksheet;
let headerInformation: HeaderInformationType = {};
let cellTracker: CellTrackerType = {};
let delimiter: string = '.';

interface Sheet {
  title: string;
  data: object | object[];
  options?: {
    [index: string]: any;
  };
  delimiter?: string;
}

interface HeaderInformationType {
  [index: string]: {
    colSpan: number;
    rowSpan: number;
    rowNumber: number;
  };
}

interface CellTrackerType {
  [index: number]: number;
}

function getSampleJson(json: object[]) {
  if (typeof json === 'object' && Array.isArray(json)) {
    if (json.length <= 0) {
      return {};
    }
    return json[0];
  }
  return json;
}

type workbookType = InstanceType<typeof Excel.Workbook>;

function addWorkSheet(workSheet: Sheet) {
  const worksheetTittle = workSheet.title;
  const worksheetOptions = workSheet?.options || {};
  return workbook.addWorksheet(worksheetTittle, worksheetOptions);
}

function setExcelHeader(flattenJson: object) {
  headerInformation = getHeaderInformation(flattenJson);
  for (const header in headerInformation) {
    if (headerInformation.hasOwnProperty(header)) {
      mergeCell(header, headerInformation[header]);
    }
  }
}

function getHeaderInformation(flattenJson: object) {
  for (const data in flattenJson) {
    if (flattenJson.hasOwnProperty(data)) {
      const splittedArray = data.split(delimiter);
      const lastHeaderKey = splittedArray[splittedArray.length - 1];
      let rowSpan = 0;
      let previousHeaderKey = "";
      for (let i = 0; i < splittedArray.length; i++){
        let headerKey = splittedArray[i];
        const rowNumber =i + 1;
        let headerKeyCopy = `${headerKey}${delimiter}${previousHeaderKey}`;
        if (
          !headerInformation?.[headerKeyCopy] ||
          (headerInformation?.[headerKeyCopy] && headerInformation[headerKeyCopy].rowNumber !== rowNumber)
        ) {
          
          if (lastHeaderKey === headerKey && i === splittedArray.length - 1) {
            rowSpan = maxDepth - rowNumber - 1;
          }
        
          headerInformation[headerKeyCopy] = {
            colSpan: 0,
            rowSpan,
            rowNumber,
          };
        } else {
          headerInformation[headerKeyCopy].colSpan += 1;
        }
        previousHeaderKey = headerKeyCopy;
      }
    }
  }
  return headerInformation;
}

function mergeCell(
  header: string,
  { colSpan, rowNumber, rowSpan }: { colSpan: number; rowNumber: number; rowSpan: number },
) {
  const startRow = rowNumber;
  const startColumn = getColumnCell(rowNumber, colSpan, rowSpan);
  const endRow = startRow + rowSpan;
  const endColumn = startColumn + colSpan;
  sheet.mergeCells(startRow, startColumn, endRow, endColumn);
  const row = sheet.getRow(rowNumber);
  const cell = row.getCell(startColumn);
  cell.value = header.split(delimiter)[0];
}

function getColumnCell(rowNumber: number, colSpan: number, rowSpan: number) {
  const startColumn = cellTracker?.[rowNumber] ? cellTracker[rowNumber] + 1 : 1;
  for (let i = rowNumber; i <= rowNumber + rowSpan; i++) {
    cellTracker[i] = startColumn + colSpan;
  }
  return startColumn;
}

function findMaxDepth(flattenJson: object) {
  Object.keys(flattenJson).forEach((data) => {
    const splittedArray = data.split(delimiter);
    if (maxDepth <= splittedArray.length) {
      maxDepth = splittedArray.length + 1;
    }
  });
}

export = {
  generateExcel: (sheetConfigurations: Sheet[]) => {
    if (!Array.isArray(sheetConfigurations)) {
      sheetConfigurations = [sheetConfigurations];
    }
    const checkSheetConfiguration = (sheetConfiguration: Sheet) => {
      if (!sheetConfiguration?.title) {
        throw new Error('Sheet title is missing in one of the sheet object');
      }
      if (!sheetConfiguration?.data) {
        throw new Error('Sheet data is missing in one of the sheet object');
      }
    }
    sheetConfigurations.forEach((sheetConfig) => {
      delimiter = sheetConfigurations[0]?.delimiter ?? '.';
      cellTracker = {};
      maxDepth = 0;
      headerInformation = {};
      sheet = addWorkSheet(sheetConfig);
      checkSheetConfiguration(sheetConfig);
      const data = Array.isArray(sheetConfig?.data) ? sheetConfig?.data : [sheetConfig?.data];
      const flattenJson: object = flatten(getSampleJson(data), {
        delimiter,
      });
      findMaxDepth(flattenJson);
      setExcelHeader(flattenJson);
      sheet.columns = Object.keys(flattenJson).map((jsonKey) => ({
        key: jsonKey,
      }));
      data.forEach((jsonData) => {
          sheet.addRow(flatten(jsonData,{
            delimiter,
          }));
      });
    });
    return workbook;
  },
};
