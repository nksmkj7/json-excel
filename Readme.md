# json-as-excel

[![Bugs](https://sonarcloud.io/api/project_badges/measure?project=nksmkj7_json-excel&metric=bugs)](https://sonarcloud.io/dashboard?id=nksmkj7_json-excel)
[![Security Rating](https://sonarcloud.io/api/project_badges/measure?project=nksmkj7_json-excel&metric=security_rating)](https://sonarcloud.io/dashboard?id=nksmkj7_json-excel)
[![Reliability Rating](https://sonarcloud.io/api/project_badges/measure?project=nksmkj7_json-excel&metric=reliability_rating)](https://sonarcloud.io/dashboard?id=nksmkj7_json-excel)
[![Maintainability Rating](https://sonarcloud.io/api/project_badges/measure?project=nksmkj7_json-excel&metric=sqale_rating)](https://sonarcloud.io/dashboard?id=nksmkj7_json-excel)
[![Lines of Code](https://sonarcloud.io/api/project_badges/measure?project=nksmkj7_json-excel&metric=ncloc)](https://sonarcloud.io/dashboard?id=nksmkj7_json-excel)

# About
Take json object as an argument and convert them into excel data. \
Json object i.e.
```json
[
  {
    "study": {
      "science": {
        "bio": {
          "pharmacy": "Kamran Bains",
          "mbbs": {
            "general": "Chloe-Ann Vega",
            "md": "Amayah Barajas"
          }
        },
        "math": {
          "pureMath": "Safa Blackburn",
          "engineering": {
            "computer": {
              "hardware": "Kezia Gonzalez",
              "software": "Boyd Mcbride"
            },
            "civil": "Leela Romero",
            "mechanical": "Mateusz Thornton"
          }
        }
      },
      "management": {
        "bba": "Amelie Bell",
        "bbs": "Jevon Myers"
      }
    }
  },
  {
    "study": {
      "science": {
        "bio": {
          "pharmacy": "Riley-James Duran",
          "mbbs": {
            "general": "Glen Churchill",
            "md": "Sachin Deacon"
          }
        },
        "math": {
          "pureMath": "Rufus Redfern",
          "engineering": {
            "computer": {
              "hardware": "Jonah Best",
              "software": "Zion Ingram"
            },
            "civil": "Matei Gibbs",
            "mechanical": "Kaelan Mcdonnell"
          }
        }
      },
      "management": {
        "bba": "Spike Peel",
        "bbs": "Zakariyah Gray"
      }
    }
  }
];
```
is converted to below like excel.\
![alt text](https://github.com/nksmkj7/json-excel/blob/main/image/sample.png?raw=true)

#### No need to manually merge cells now !! ???? ????

# Installation 
```js
npm install json-as-excel
```

# Usage
```js
const excel = require('json-as-excel');
const data = [
        {
          study: {
            science: {
              bio: {
                pharmacy: 'Kamran Bains',
                mbbs: {
                  general: 'Chloe-Ann Vega',
                  md: 'Amayah Barajas',
                },
              },
              math: {
                pureMath: 'Safa Blackburn',
                engineering: {
                  computer: {
                    hardware: 'Kezia Gonzalez',
                    software: 'Boyd Mcbride',
                  },
                  civil: 'Leela Romero',
                  mechanical: 'Mateusz Thornton',
                },
              },
            },
            management: {
              bba: 'Amelie Bell',
              bbs: 'Jevon Myers',
            },
          },
        },
        {
          study: {
            science: {
              bio: {
                pharmacy: 'Riley-James Duran',
                mbbs: {
                  general: 'Glen Churchill',
                  md: 'Sachin Deacon',
                },
              },
              math: {
                pureMath: 'Rufus Redfern',
                engineering: {
                  computer: {
                    hardware: 'Jonah Best',
                    software: 'Zion Ingram',
                  },
                  civil: 'Matei Gibbs',
                  mechanical: 'Kaelan Mcdonnell',
                },
              },
            },
            management: {
              bba: 'Spike Peel',
              bbs: 'Zakariyah Gray',
            },
          },
        },
      ]
const workbook = excel.generateExcel([
    {
      title: 'First sheet',
      data: data,
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
> generateExcel([{ title, data, delimiter, options }]).<br>
Method generateExcel accepts array of objects. Each object represents individual sheet. This method returns exceljs workbook instance.

### title
Title is name for sheet.

### data
Data is json object whose keys are generated as header in excel and values are placed as new row per object.

### delimiter (optional)
  `.` is used as a default delimiter. If json data consists key with `'.'`, one need to change delimiter to any other delimiter.
  ```js
  generateExcel([{title:"firstSheet", data:data, delimiter:"%"}])
  ```

### options (optional)
options are the exceljs available worksheet options i.e. [Worksheet Properties](https://www.npmjs.com/package/exceljs#worksheet-properties), [Page Setup](https://www.npmjs.com/package/exceljs#page-setup), [Headers and Footers](https://www.npmjs.com/package/exceljs#headers-and-footers)\
More detail can be obtained from [exceljs](https://www.npmjs.com/package/exceljs)
 ```js
  const options = {
    properties:{
      outlineLevelCol:2,
      tabColor:{
        argb:'FF00FF00'
      },
      defaultRowHeight:15
    },
    pageSetup:{
      fitToPage: true,
      fitToHeight: 5, 
      fitToWidth: 7
    }
  };
  generateExcel([{title:"firstSheet", data:data, delimiter:"%", options:options}])
  ```

### headerFormatter (optional)
 headerFormatter is a  function that accepts header as an argument and return computed header.
 ```
  generateExcel([{title:"firstSheet", data:data, delimiter:"%",headerFormatter: (header) => header.toUpperCase()}])
```
In above example, excel that has header, all with upper case will be generated.

## Acknowledgments
1. [ exceljs](https://www.npmjs.com/package/exceljs) 
2. [ flat ](https://www.npmjs.com/package/flat)

## Issues
> If any issue is found, please raise issue in github.

## Changelog
| Version | Changes |
| ----------- | ----------- |
| 1.0.4 | <ul><li>Installation guide update in Readme</li></ul> | |
| 1.0.5 | <ul><li>Example updated in github</li><li>Bug Fixes<ul><li>Fixed crash when sheet data is empty object</li><li>Check mandatory title and data for sheet configuration. If not provided, error is thrown</li></ul> </li></ul> | |
| 1.0.6 | <ul><li>Test updated for case when data has same nested key</li><li>Bug Fixes<ul><li>Cell merge issue when sheet data has same nested key</li></ul> </li></ul> | |
| 1.0.7 | <ul><li>Feature<ul><li>Feature to change header format is now added. For reference, look at headerFormatter option above.</li></ul> </li></ul> | |
| 1.0.8 | <ul><li>Readme updated</li></ul> | |
| 1.0.9 | <ul><li>Bug Fixes<ul><li>Worksheet name already exists issue fixed.</li></ul> </li></ul> | |

## MIT License

```
Copyright (c) 2021
```
