# json-excel

[![Bugs](https://sonarcloud.io/api/project_badges/measure?project=nksmkj7_json-excel&metric=bugs)](https://sonarcloud.io/dashboard?id=nksmkj7_json-excel)
[![Security Rating](https://sonarcloud.io/api/project_badges/measure?project=nksmkj7_json-excel&metric=security_rating)](https://sonarcloud.io/dashboard?id=nksmkj7_json-excel)
[![Reliability Rating](https://sonarcloud.io/api/project_badges/measure?project=nksmkj7_json-excel&metric=reliability_rating)](https://sonarcloud.io/dashboard?id=nksmkj7_json-excel)
[![Maintainability Rating](https://sonarcloud.io/api/project_badges/measure?project=nksmkj7_json-excel&metric=sqale_rating)](https://sonarcloud.io/dashboard?id=nksmkj7_json-excel)
[![Lines of Code](https://sonarcloud.io/api/project_badges/measure?project=nksmkj7_json-excel&metric=ncloc)](https://sonarcloud.io/dashboard?id=nksmkj7_json-excel)

# About
Take json object as an argument and convert them into excel data. \
Json object i.e.
```json
[{"study":{"science":{"bio":{"pharmacy":"ritu","mbbs":{"general":"roshan","md":"sanjay"}},"math":{"pureMath":"rukesh","engineering":{"computer":{"hardware":"Aungush","software":"nikesh"},"civil":"seena","mechanical":"santosh"}}},"management":{"bba":"pratik","bbs":"jeena"}}},{"study":{"science":{"bio":{"pharmacy":"rajani","mbbs":{"general":"haris","md":"shreetika"}},"math":{"pureMath":"prijal","engineering":{"computer":{"hardware":"samina","software":"anish"},"civil":"rasil","mechanical":"amit"}}},"management":{"bba":"anjeela","bbs":"sushmin"}}}];
```
is converted to below like excel.\
![alt text](./image/sample.png)

#### No need to manually merge cells now !! 😊 🤩

# Installation 
```js
npm install json-excel
```
This trait helps you to skip rows that don't satisfy specified rules in rules function and generated separated excel file of skipped rows with reasons.
# Usage
```js
const excel = require('json-excel');
const workbook = excel.generateExcel([
    {
      title: 'First sheet',
      data: [
        {
          study: {
            science: {
              bio: {
                pharmacy: 'bandana',
                mbbs: {
                  general: 'roshan',
                  md: 'sanjay',
                },
              },
              math: {
                pureMath: 'rukesh',
                engineering: {
                  computer: {
                    hardware: 'Aungush',
                    software: 'nikesh',
                  },
                  civil: 'seena',
                  mechanical: 'santosh',
                },
              },
            },
            management: {
              bba: 'pratik',
              bbs: 'jeena',
            },
          },
        },
        {
          study: {
            science: {
              bio: {
                pharmacy: 'rajani',
                mbbs: {
                  general: 'haris',
                  md: 'shreetika',
                },
              },
              math: {
                pureMath: 'prijal',
                engineering: {
                  computer: {
                    hardware: 'samina',
                    software: 'anish',
                  },
                  civil: 'rasil',
                  mechanical: 'amit',
                },
              },
            },
            management: {
              bba: 'anjeela',
              bbs: 'sushmin',
            },
          },
        },
      ],
    },
  ]);
```

> generateExcel function returns exceljs workbook instance. Hence, File I/O can be achieved same as in exceljs. For example: 
```js
// write to a file
await workbook.xlsx.writeFile('sample.xlsx');
```
> For the detail reference of [ File I/O](https://www.npmjs.com/package/exceljs#file-io)

# Method
generateExcel([{ title, data, delimiter, options }]). \ 
Method generate

### title
Title is name for sheet.

