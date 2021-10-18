import * as Excel from 'exceljs';
interface Sheet {
    title: string;
    data: object | object[];
    options?: {
        [index: string]: any;
    };
    delimiter?: string;
}
declare const _default: {
    generateExcel: (sheetConfigurations: Sheet[]) => Excel.Workbook;
};
export = _default;
