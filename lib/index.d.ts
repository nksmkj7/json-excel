import * as Excel from 'exceljs';
declare let sheet: Excel.Worksheet;
interface sheet {
    title: string;
    data: object | object[];
    options?: {
        [index: string]: any;
    };
    delimiter?: string;
}
declare const _default: {
    generateExcel: (sheetConfigurations: sheet[]) => Excel.Workbook;
};
export = _default;
