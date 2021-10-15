import * as Excel from 'exceljs';
declare let sheet: Excel.Worksheet;
interface sheet {
    title: string;
    data: object | object[];
    options?: object;
}
declare const _default: {
    delimiter: string;
    setDelimiter: (delimiter: string) => void;
    getDelimiter: () => string;
    generateExcel: (sheetConfigurations: sheet[]) => Excel.Workbook;
};
export = _default;
