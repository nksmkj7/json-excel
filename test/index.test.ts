import * as jsonExcel from "../lib/index";
import rewire from "rewire";
let myModule = rewire('../lib/index');
import { flatten } from 'flat';
let json: any = [{"study":{"science":{"bio":{"pharmacy":"ritu","mbbs":{"general":"roshan","md":"sanjay"}},"math":{"pureMath":"rukesh","engineering":{"computer":{"hardware":"Aungush","software":"nikesh"},"civil":"seena","mechanical":"santosh"}}},"management":{"bba":"pratik","bbs":"jeena"}}},{"study":{"science":{"bio":{"pharmacy":"rajani","mbbs":{"general":"haris","md":"shreetika"}},"math":{"pureMath":"prijal","engineering":{"computer":{"hardware":"samina","software":"anish"},"civil":"rasil","mechanical":"amit"}}},"management":{"bba":"anjeela","bbs":"sushmin"}}}];
import * as Excel from 'exceljs';
const workbook = new Excel.Workbook();

describe("test json to excel", () => {
    const findMaxDepth: Function = myModule.__get__('findMaxDepth');
    const getSampleJson: Function = myModule.__get__('getSampleJson')
    let sampleJson: object = getSampleJson(json)
    let flattenJson = flatten(sampleJson);
    it("should return the first element of parsed sample json", () => {
        const result: object = {
            study: {
                science: {
                    bio: {
                        pharmacy: "ritu",
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
        };
        expect(sampleJson).toEqual(result);
        
    })
    
    it("should return the maximum depth of the json including value", () => {
        findMaxDepth(flattenJson);
        let changedMaxDepth:number = myModule.__get__('maxDepth');
        expect(changedMaxDepth).toBe(7);
    })

    it("should return information of each key expected excel row number, colspan and rowspan",() => {
        let getHeaderInformation = myModule.__get__('getHeaderInformation');
        let headerInformation = getHeaderInformation(flattenJson);
        let result = {
            study: { colSpan: 9, rowSpan: 0, rowNumber: 1 },
            science: { colSpan: 7, rowSpan: 0, rowNumber: 2 },
            bio: { colSpan: 2, rowSpan: 0, rowNumber: 3 },
            pharmacy: { colSpan: 0, rowSpan: 2, rowNumber: 4 },
            mbbs: { colSpan: 1, rowSpan: 0, rowNumber: 4 },
            general: { colSpan: 0, rowSpan: 1, rowNumber: 5 },
            md: { colSpan: 0, rowSpan: 1, rowNumber: 5 },
            math: { colSpan: 4, rowSpan: 0, rowNumber: 3 },
            pureMath: { colSpan: 0, rowSpan: 2, rowNumber: 4 },
            engineering: { colSpan: 3, rowSpan: 0, rowNumber: 4 },
            computer: { colSpan: 1, rowSpan: 0, rowNumber: 5 },
            hardware: { colSpan: 0, rowSpan: 0, rowNumber: 6 },
            software: { colSpan: 0, rowSpan: 0, rowNumber: 6 },
            civil: { colSpan: 0, rowSpan: 1, rowNumber: 5 },
            mechanical: { colSpan: 0, rowSpan: 1, rowNumber: 5 },
            management: { colSpan: 1, rowSpan: 0, rowNumber: 2 },
            bba: { colSpan: 0, rowSpan: 3, rowNumber: 3 },
            bbs: { colSpan: 0, rowSpan: 3, rowNumber: 3 }
          };
        expect(headerInformation).toEqual(result);
    })

    describe("should track occupied column of specific row and return column number to start",() => {
        let getColumnCell = myModule.__get__('getColumnCell');
        let startingColumn = getColumnCell(1,9,0);
        it("should return starting column",() => {
            expect(startingColumn).toBe(1);
        })
        it("should track occupied column for given row number upto the row occupied with row span",()=>{
            let cellTracker = myModule.__get__('cellTracker');
            expect(cellTracker).toEqual({"1":10})
        })
    })

    it("should merge cells and set the cell value",() => {
        myModule.__set__('cellTracker',{});
        let mergeCell = myModule.__get__("mergeCell");
        myModule.__set__("sheet",workbook.addWorksheet("new test"));
        mergeCell("apple",{rowNumber:1,colSpan:9,rowSpan:0})
        let sheet = myModule.__get__("sheet");
        const row = sheet.getRow(1);
        expect(row.getCell(1).value).toEqual(row.getCell(9).value)
    })

    describe("When provide json data to the function", () => {
        
        let data = {
            "header": {
                "column": "test data"
            }
        };
        it("should generate excel that has testSheet as sheet name", () => {
            let generatedWorkBook = jsonExcel.generateExcel([{ title: "testSheet", data: data }]);
            expect(generatedWorkBook.getWorksheet("testSheet").name).toEqual("testSheet");
        })

        it("should generate excel that has second test Sheet as sheet name and column name with .", () => {
            let data = {
                "header": {
                    "column.fullStop": "test data"
                }
            };
            let generatedWorkBook = jsonExcel.generateExcel([{ title: "second test Sheet", data: data,delimiter: "%"}]);
            let sheet = generatedWorkBook.getWorksheet("second test Sheet")
            expect(sheet.getRow(2).values).toEqual(expect.arrayContaining(['column.fullStop']));
        })

        it("should generate excel with two sheet", () => {
            let generatedWorkBook = jsonExcel.generateExcel([{ title: "first sheet", data: data, delimiter: "%" }, { title: "second sheet", data: data, delimiter: "%" }]);
            let sheets:string[] = [];
            generatedWorkBook.eachSheet(function (worksheet) {
                sheets.push(worksheet.name);
            });
            expect(sheets).toEqual(expect.arrayContaining(['first sheet','second sheet']));
        })

        it("should apply options for respective sheet", () => {
            let generatedWorkBook = jsonExcel.generateExcel([{ title: "test sheet with options", data: data, delimiter: "%", options: { properties: { outlineLevelCol: 2 } } }]);
            let sheet = generatedWorkBook.getWorksheet("test sheet with options");
            expect(sheet.properties.outlineLevelCol).toBe(2);
        })

    })

})