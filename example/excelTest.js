// const excel = require('json-as-excel');
const excel = require('../lib/index');
async function generate() {
  const workbook = excel.generateExcel([
    {
      title: 'First sheet',
      data: [
        {
          study: {
            science_test: {
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
            science_test: {
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
      ],
      headerFormatter: (header) => {
        const { startCase } = require('lodash');
        return startCase(header);
      },
    },
    {
      title: 'second sheet',
      data: [],
    },
    {
      title: 'third sheet',
      data: [],
    },
  ]);
  await workbook.xlsx.writeFile('sample.xlsx');
}
generate();
