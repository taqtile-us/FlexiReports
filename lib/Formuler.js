var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import ExcelJS from 'exceljs';
const getComments = (workbook) => __awaiter(void 0, void 0, void 0, function* () {
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
});
const getCoordsOfFormulaCell = (path) => __awaiter(void 0, void 0, void 0, function* () {
    const workbook = new ExcelJS.Workbook();
    yield workbook.xlsx.readFile(path);
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
});
const generateReportWithFormula = (path, formulas) => __awaiter(void 0, void 0, void 0, function* () {
    try {
        const workbook = new ExcelJS.Workbook();
        const formulasData = Object.values(formulas);
        if (!formulasData.length) {
            return;
        }
        yield workbook.xlsx.readFile(path);
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
        yield workbook.xlsx.writeFile(path);
    }
    catch (e) {
        console.log(e, 'generate report error');
        return false;
    }
});
export { generateReportWithFormula, getCoordsOfFormulaCell };
