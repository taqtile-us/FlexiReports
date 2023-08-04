"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const generate_excel_report_1 = require("./generate-excel-report");
const Formuler_1 = require("./Formuler");
const generate = async (dataToFill, filePath) => {
    const res = await (0, generate_excel_report_1.generateExcelReport)(dataToFill, filePath);
    console.log(res, 'res');
    if (res) {
        const coords = await (0, Formuler_1.getCoordsOfFormulaCell)(res);
        console.log(coords, 'coords');
        await (0, Formuler_1.generateReportWithFormula)(res, coords);
    }
    return res;
};
exports.default = generate;
