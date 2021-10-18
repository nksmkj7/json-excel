const excel = require("../lib/index");
async function generate() {
    const workbook = excel.generateExcel([
        {
            title: "First sheet",
            data: [
                {
                    study: {
                        science: {
                            "bio.sio": {
                                pharmacy: "bandana",
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
                },
                {
                    study: {
                        science: {
                            "bio.sio": {
                                pharmacy: "rajani",
                                mbbs: {
                                    general: "haris",
                                    md: "shreetika",
                                },
                            },
                            math: {
                                pureMath: "prijal",
                                engineering: {
                                    computer: {
                                        hardware: "samina",
                                        software: "anish",
                                    },
                                    civil: "rasil",
                                    mechanical: "amit",
                                },
                            },
                        },
                        management: {
                            bba: "anjeela",
                            bbs: "sushmin",
                        },
                    },
                },
            ],
            delimiter: "%",
        },
    ]);
    await workbook.xlsx.writeFile("cat.xlsx");
}
generate();
