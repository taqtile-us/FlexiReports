"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const generate_excel_report_1 = require("./generate-excel-report");
const Formuler_1 = require("./Formuler");
const generate = async (dataToFill, templatePath, reportPath) => {
    const report = await (0, generate_excel_report_1.generateExcelReport)(dataToFill, templatePath, reportPath);
    if (report) {
        const coords = await (0, Formuler_1.getCoordsOfFormulaCell)(report);
        await (0, Formuler_1.generateReportWithFormula)(report, coords);
    }
    return report;
};
exports.default = generate;
