"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.writeDataToExcel = void 0;
const path_1 = __importDefault(require("path"));
// @ts-ignore
const exceljs_1 = __importDefault(require("exceljs"));
const copyChart_1 = require("./copy-excel-chart/build/copyChart");
const readCharts_1 = require("./copy-excel-chart/build/readCharts");
const writeChart_1 = require("./copy-excel-chart/build/writeChart");
const Parser_1 = require("./Parser");
const cellFont = { name: 'Arial', size: 11 };
const fs = require('fs').promises;
function readFile(path) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            const workbook = new exceljs_1.default.Workbook();
            yield workbook.xlsx.readFile(path);
            return { workSheet: workbook.getWorksheet('template'), workbook };
        }
        catch (err) {
            console.error('Error reading the file:', err);
            return null;
        }
    });
}
const createArrayIfNotExist = (list, entity) => {
    if (!list[entity]) {
        list[entity] = [];
    }
};
const copyDiagramm = (template, report, length) => __awaiter(void 0, void 0, void 0, function* () {
    const source = yield (0, readCharts_1.readCharts)(template, './temp');
    const output = yield (0, readCharts_1.readCharts)(report, './temp');
    const summary = source.summary();
    let replaceCellRefs = summary['template']['chart1'].reduce((acc, el) => {
        return Object.assign(Object.assign({}, acc), { [el]: el.replace('recommendWorksheet2', 'worksheet-Recommendation') });
    }, {});
    for (let key in replaceCellRefs) {
        replaceCellRefs[key] = (0, Parser_1.extendRange)(replaceCellRefs[key], length);
    }
    (0, copyChart_1.copyChart)(source, output, 'template', 'chart1', 'template', replaceCellRefs);
    (0, writeChart_1.writeCharts)(output, report);
});
function parse(worksheet) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            const simpleVariables = {};
            let master = null;
            const details = {};
            const formulas = { rowFormulas: [], columnFormulas: [], masterFormulas: [] };
            const staticVariables = {};
            let masterRowNumber = -1;
            worksheet === null || worksheet === void 0 ? void 0 : worksheet.eachRow((row, rowNumber) => {
                row.eachCell((cell, colNumber) => {
                    const simpleVariable = (0, Parser_1.getSimpleVariable)(cell);
                    if (simpleVariable) {
                        createArrayIfNotExist(simpleVariables, simpleVariable.variable);
                        simpleVariables[simpleVariable.variable].push({
                            address: simpleVariable.address,
                            alignment: cell.alignment,
                            variable: simpleVariable.variable
                        });
                    }
                    const complexVariable = (0, Parser_1.getComplexVariable)(cell);
                    if (complexVariable) {
                        if (complexVariable.type === 'master') {
                            master = complexVariable;
                            masterRowNumber = rowNumber;
                        }
                        if (complexVariable.type === 'detail') {
                            createArrayIfNotExist(details, complexVariable.entityName);
                            details[complexVariable.entityName].push(complexVariable);
                        }
                    }
                    const formula = (0, Parser_1.getFormula)(cell);
                    if (formula) {
                        if ((0, Parser_1.isItRowFormula)(formula.formula)) {
                            if (masterRowNumber === rowNumber) {
                                formulas.masterFormulas.push(formula);
                            }
                            else {
                                formulas.rowFormulas.push(formula);
                            }
                        }
                        else {
                            formulas.columnFormulas.push(formula);
                        }
                    }
                    if (!simpleVariable && !complexVariable && !formula) {
                        staticVariables[cell.address] = {
                            value: cell.value,
                            address: cell.address,
                            alignment: cell.alignment
                        };
                    }
                });
            });
            return { simpleVariables, master, details, formulas, staticVariables };
        }
        catch (err) {
            console.error('Error parse the file:', err);
            return {
                simpleVariables: {},
                master: {},
                details: {},
                formulas: { rowFormulas: [], columnFormulas: [], masterFormulas: [] },
                staticVariables: {}
            };
        }
    });
}
function putMasterDetail(worksheet, master, details, data, formulas, staticVariables) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            const masterDetailDeep = 1;
            let currentRowNumber = (0, Parser_1.parseIntFromString)(master.address);
            const startRowNumber = currentRowNumber;
            const column = (0, Parser_1.parseLetterFromString)(master.address);
            const detailsKeys = Object.keys(details);
            if (!column || !currentRowNumber)
                return null;
            data[master.entityName].forEach((masterEntity) => {
                // put master row
                const cell = worksheet.getCell(column + currentRowNumber);
                cell.value = masterEntity[master.fieldName];
                cell.font = cellFont;
                cell.alignment = master.alignment;
                // put master formula
                if (formulas.masterFormulas.length) {
                    const masterFormulaColumn = (0, Parser_1.parseLetterFromString)(formulas.masterFormulas[0].address);
                    if (masterFormulaColumn) {
                        formulas.masterFormulas.forEach((masterFormula) => {
                            const cell = worksheet.getCell(masterFormulaColumn + currentRowNumber);
                            cell.value = { formula: `SUM(G${currentRowNumber + 1}:G${currentRowNumber + masterEntity[detailsKeys[0]].length})` };
                            cell.font = cellFont;
                            cell.alignment = masterFormula.alignment;
                        });
                    }
                }
                currentRowNumber += 1;
                worksheet.spliceRows(currentRowNumber + masterDetailDeep, 0, []);
                // put details row
                masterEntity[detailsKeys[0]].forEach((detailEntity) => {
                    details[detailsKeys[0]].forEach((detailCoords) => {
                        if (detailEntity[detailCoords.fieldName]) {
                            const detailColumn = (0, Parser_1.parseLetterFromString)(detailCoords.address);
                            if (!detailColumn)
                                return;
                            const detailCell = worksheet.getCell(detailColumn + currentRowNumber);
                            detailCell.value = detailEntity[detailCoords.fieldName];
                            detailCell.font = cellFont;
                            detailCell.alignment = detailCoords.alignment;
                        }
                    });
                    // put row formulas
                    formulas.rowFormulas.forEach((formula) => {
                        const originalRowNumber = (0, Parser_1.parseIntFromString)(formula.formula);
                        const newRowNumber = originalRowNumber + currentRowNumber - startRowNumber - masterDetailDeep;
                        const movedFormula = (0, Parser_1.replaceSpecificNumberInFormula)(formula.formula, originalRowNumber, newRowNumber);
                        const movedAddress = (0, Parser_1.replaceSpecificNumberInFormula)(formula.address, originalRowNumber, newRowNumber);
                        const formulaCell = worksheet.getCell(movedAddress);
                        formulaCell.value = { formula: movedFormula };
                        formulaCell.alignment = formula.alignment;
                    });
                    // put static variables
                    for (let variableAddress in staticVariables) {
                        if ((0, Parser_1.parseIntFromString)(variableAddress) === startRowNumber + 1) {
                            const staticVariableColumn = (0, Parser_1.parseLetterFromString)(variableAddress);
                            const staticVariableCell = worksheet.getCell(`${staticVariableColumn}${currentRowNumber}`);
                            staticVariableCell.value = staticVariables[variableAddress].value;
                            staticVariableCell.alignment = staticVariables[variableAddress].alignment;
                        }
                    }
                    currentRowNumber += 1;
                    worksheet.spliceRows(currentRowNumber + masterDetailDeep, 0, []);
                });
            });
            // put column formulas
            const difference = currentRowNumber - startRowNumber;
            formulas.columnFormulas.forEach((formula) => {
                const originalRowNumber = (0, Parser_1.parseIntFromString)(formula.address);
                const formulaCell = worksheet.getCell((0, Parser_1.replaceSpecificNumberInFormula)(formula.address, originalRowNumber, originalRowNumber + difference));
                formulaCell.value = { formula: (0, Parser_1.addDifferenceToTheLastNumber)(formula.formula, difference) };
                formulaCell.alignment = formula.alignment;
            });
            return currentRowNumber - startRowNumber;
        }
        catch (err) {
            console.error('Error put master detail to the file:', err);
            return null;
        }
    });
}
function putSimpleVariables(worksheet, data, simpleVariables) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            for (let variable in simpleVariables) {
                if (data[variable]) {
                    simpleVariables[variable].forEach((simpleVariable) => {
                        const variableCell = worksheet.getCell(simpleVariable.address);
                        variableCell.value = data[variable];
                        variableCell.alignment = simpleVariable.alignment;
                    });
                }
            }
        }
        catch (err) {
            console.error('Error put simple variables to the file:', err);
            return null;
        }
    });
}
const buildTemplate = (dataToFill, path) => __awaiter(void 0, void 0, void 0, function* () {
    const { workbook, workSheet } = yield readFile(path);
    const { master, details, simpleVariables, formulas, staticVariables } = yield parse(workSheet);
    const masterTyped = master;
    const detailsTyped = details;
    // put simple variables
    putSimpleVariables(workSheet, dataToFill, simpleVariables);
    // put master-details
    const lenght = yield putMasterDetail(workSheet, masterTyped, detailsTyped, dataToFill, formulas, staticVariables);
    if (lenght) {
        yield workbook.xlsx.writeFile('report.xlsx');
        copyDiagramm(path, './report.xlsx', lenght);
    }
});
const writeDataToExcel = (dataToFill, templatePath) => __awaiter(void 0, void 0, void 0, function* () {
    const filePath = './report.xlsx';
    const buffer = yield buildTemplate(dataToFill, path_1.default.join(templatePath));
    // await writeFile(filePath, buffer, 'binary');
    return filePath;
});
exports.writeDataToExcel = writeDataToExcel;