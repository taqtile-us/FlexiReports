"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.getCoordsOfFormulaCell = exports.generateReportWithFormula = void 0;
const exceljs_1 = __importDefault(require("exceljs"));
const getComments = async (workbook) => {
    const worksheet = workbook.getWorksheet(1);
    const coords = {};
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const comment = cell._comment.note.texts[0].text;
            if (comment) {
                coords[comment] = { colNumber, rowNumber, cell };
            }
        });
    });
    return coords;
};
const getCoordsOfFormulaCell = async (reportPath) => {
    const workbook = new exceljs_1.default.Workbook();
    await workbook.xlsx.readFile(reportPath);
    const worksheet = workbook.getWorksheet(1);
    const coords = {};
    worksheet.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
            const formula = cell._value.model.formula;
            if (formula) {
                const address = cell._address;
                const formulaValue = cell._value.model.formula;
                const number = cell._row._number;
                coords[address] = { address, value: formulaValue, numberAddress: number };
            }
        });
    });
    return coords;
};
exports.getCoordsOfFormulaCell = getCoordsOfFormulaCell;
const generateReportWithFormula = async (reportPath, formulas) => {
    try {
        const workbook = new exceljs_1.default.Workbook();
        const formulasData = Object.values(formulas);
        if (!formulasData.length) {
            return;
        }
        await workbook.xlsx.readFile(reportPath);
        const worksheet = workbook.getWorksheet(1);
        formulasData.forEach((formulaData) => {
            const formula = formulaData.value.toString();
            if (formula.includes('(') && formula.includes(')')) {
                const partsOfRange = formula.split('(');
                const startRangeIndex = formula.indexOf('(') + '('.length;
                const range = formula.substring(startRangeIndex, formula.indexOf(')', startRangeIndex));
                if (range.includes(':')) {
                    const prevPosition = range.split(':');
                    const updatedRange = prevPosition[1][0] + (formulaData.numberAddress - 1);
                    prevPosition[1] = updatedRange;
                    const secondRangePart = '(' + prevPosition.join(':') + ')';
                    worksheet.getCell(formulaData.address).value = {
                        formula: partsOfRange[0] + secondRangePart,
                        date1904: false,
                    };
                }
            }
        });
        await workbook.xlsx.writeFile(reportPath);
    }
    catch (e) {
        console.log(e, 'generate report error');
        return false;
    }
};
exports.generateReportWithFormula = generateReportWithFormula;
