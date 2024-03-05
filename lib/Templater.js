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
const exceljs_1 = __importDefault(require("exceljs"));
const copyChart_1 = require("./copy-excel-chart/build/copyChart");
const readCharts_1 = require("./copy-excel-chart/build/readCharts");
const writeChart_1 = require("./copy-excel-chart/build/writeChart");
const Parser_1 = require("./Parser");
const cellFont = { name: 'Arial', size: 11 };
const fs = require('fs').promises;
const sizeOf = require('image-size');
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
const copyDiagramm = (template, report, length, temporaryFolderPath) => __awaiter(void 0, void 0, void 0, function* () {
    try {
        yield fs.mkdir(temporaryFolderPath, { recursive: true });
        console.log(`Folder created (or already exists): ${temporaryFolderPath}`);
    }
    catch (error) {
        console.error(`Error creating folder: ${error.message}`);
    }
    try {
        const source = yield (0, readCharts_1.readCharts)(template, temporaryFolderPath);
        const output = yield (0, readCharts_1.readCharts)(report, temporaryFolderPath);
        source.worksheets.template.drawingRels.chart1 = 'rid3';
        const summary = source.summary();
        if (summary['template']['chart1']) {
            const replaceCellRefs = summary['template']['chart1'].reduce((acc, el) => {
                return Object.assign(Object.assign({}, acc), { [el]: el.replace('recommendWorksheet2', 'worksheet-Recommendation') });
            }, {});
            for (const key in replaceCellRefs) {
                replaceCellRefs[key] = (0, Parser_1.extendRange)(replaceCellRefs[key], length);
            }
            yield (0, copyChart_1.copyChart)(source, output, 'template', 'chart1', 'template', replaceCellRefs);
            yield (0, writeChart_1.writeCharts)(output, report);
        }
    }
    catch (error) {
        console.error(`Error copy diagramm: ${error.message}`);
    }
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
                            variable: simpleVariable.variable,
                        });
                    }
                    const complexVariable = (0, Parser_1.getComplexVariable)(cell);
                    if (complexVariable) {
                        if (masterRowNumber === rowNumber && !master.addedToDetails) {
                            createArrayIfNotExist(details, master.entityName);
                            details[master.entityName].push(master);
                            master.addedToDetails = true;
                        }
                        if (complexVariable.type === 'master' && masterRowNumber === -1) {
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
                            alignment: cell.alignment,
                        };
                    }
                });
            });
            if (Object.keys(details).length === 0) {
            }
            return { simpleVariables, master, details, formulas, staticVariables };
        }
        catch (err) {
            console.error('Error parse the file:', err);
            return {
                simpleVariables: {},
                master: {},
                details: {},
                formulas: { rowFormulas: [], columnFormulas: [], masterFormulas: [] },
                staticVariables: {},
            };
        }
    });
}
function putMasterDetail(worksheet, master, details, data, formulas, staticVariables) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            const masterDetailDeep = 1;
            const detailsDeep = 0;
            let currentRowNumber = (0, Parser_1.parseIntFromString)(master.address);
            const startRowNumber = currentRowNumber;
            const column = (0, Parser_1.parseLetterFromString)(master.address);
            const detailsKeys = Object.keys(details);
            if (!column || !currentRowNumber)
                return null;
            data[master.entityName.toLowerCase()].forEach((masterEntity) => {
                if (!master.addedToDetails) {
                    if (masterEntity[detailsKeys[0]].length === 0) {
                        return;
                    }
                    putMasterRow(worksheet, masterEntity, master, column, currentRowNumber);
                    putMasterFormulas(worksheet, formulas, masterEntity, detailsKeys, currentRowNumber);
                    currentRowNumber += 1;
                    worksheet.spliceRows(currentRowNumber + masterDetailDeep, 0, []);
                    masterEntity[detailsKeys[0]].forEach((detailEntity) => {
                        putDetailRow(worksheet, details, detailsKeys, detailEntity, currentRowNumber);
                        putDetailFormula(worksheet, formulas, currentRowNumber, startRowNumber, masterDetailDeep);
                        putStaticVariables(worksheet, staticVariables, currentRowNumber, startRowNumber);
                        currentRowNumber += 1;
                        worksheet.spliceRows(currentRowNumber + masterDetailDeep, 0, []);
                    });
                }
                else {
                    putDetailRow(worksheet, details, detailsKeys, masterEntity, currentRowNumber);
                    putDetailFormula(worksheet, { rowFormulas: formulas.masterFormulas }, currentRowNumber, startRowNumber, detailsDeep);
                    putStaticVariables(worksheet, staticVariables, currentRowNumber, startRowNumber - 1);
                    currentRowNumber += 1;
                    worksheet.spliceRows(currentRowNumber + masterDetailDeep, 0, []);
                }
            });
            const difference = currentRowNumber - startRowNumber;
            putColumnFormulas(worksheet, formulas, difference);
            return currentRowNumber - startRowNumber;
        }
        catch (err) {
            console.error('Error put master detail to the file:', err);
            return null;
        }
    });
}
function putMasterRow(worksheet, masterEntity, master, column, currentRowNumber) {
    const cell = worksheet.getCell(column + currentRowNumber);
    cell.value = masterEntity[master.fieldName];
    cell.font = cellFont;
    cell.alignment = master.alignment;
}
function putMasterFormulas(worksheet, formulas, masterEntity, detailsKeys, currentRowNumber) {
    if (formulas.masterFormulas.length) {
        const masterFormulaColumn = (0, Parser_1.parseLetterFromString)(formulas.masterFormulas[0].address);
        if (masterFormulaColumn) {
            formulas.masterFormulas.forEach((masterFormula) => {
                const cell = worksheet.getCell(masterFormulaColumn + currentRowNumber);
                cell.value = {
                    formula: `SUM(G${currentRowNumber + 1}:G${currentRowNumber + masterEntity[detailsKeys[0]].length})`,
                };
                cell.font = cellFont;
                cell.alignment = masterFormula.alignment;
            });
        }
    }
}
function putDetailRow(worksheet, details, detailsKeys, detailEntity, currentRowNumber) {
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
}
function putDetailFormula(worksheet, formulas, currentRowNumber, startRowNumber, masterDetailDeep) {
    formulas.rowFormulas.forEach((formula) => {
        const originalRowNumber = (0, Parser_1.parseIntFromString)(formula.formula);
        const newRowNumber = originalRowNumber + currentRowNumber - startRowNumber - masterDetailDeep;
        const movedFormula = (0, Parser_1.replaceSpecificNumberInFormula)(formula.formula, originalRowNumber, newRowNumber);
        const movedAddress = (0, Parser_1.replaceSpecificNumberInFormula)(formula.address, originalRowNumber, newRowNumber);
        const formulaCell = worksheet.getCell(movedAddress);
        formulaCell.value = { formula: movedFormula };
        formulaCell.font = cellFont;
        formulaCell.alignment = formula.alignment;
    });
}
function putColumnFormulas(worksheet, formulas, difference) {
    formulas.columnFormulas.forEach((formula) => {
        const originalRowNumber = (0, Parser_1.parseIntFromString)(formula.address);
        const formulaCell = worksheet.getCell((0, Parser_1.replaceSpecificNumberInFormula)(formula.address, originalRowNumber, originalRowNumber + difference));
        formulaCell.value = { formula: (0, Parser_1.addDifferenceToTheLastNumber)(formula.formula, difference) };
        formulaCell.alignment = formula.alignment;
    });
}
function putStaticVariables(worksheet, staticVariables, currentRowNumber, startRowNumber) {
    for (const variableAddress in staticVariables) {
        if ((0, Parser_1.parseIntFromString)(variableAddress) === startRowNumber + 1) {
            const staticVariableColumn = (0, Parser_1.parseLetterFromString)(variableAddress);
            const staticVariableCell = worksheet.getCell(`${staticVariableColumn}${currentRowNumber}`);
            staticVariableCell.value = staticVariables[variableAddress].value;
            staticVariableCell.font = cellFont;
            staticVariableCell.alignment = staticVariables[variableAddress].alignment;
        }
    }
}
function putSimpleVariables(worksheet, data, simpleVariables) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            for (const variable in simpleVariables) {
                simpleVariables[variable].forEach((simpleVariable) => {
                    const valueToPut = data[variable.toLowerCase()] ? data[variable.toLowerCase()] : '';
                    const variableCell = worksheet.getCell(simpleVariable.address);
                    variableCell.value = valueToPut;
                    variableCell.alignment = simpleVariable.alignment;
                    simpleVariable.insertedValue = valueToPut;
                });
            }
        }
        catch (err) {
            console.error('Error put simple variables to the file:', err);
            return null;
        }
    });
}
const buildTemplate = (dataToFill, path, reportPath, temporaryFolderPath) => __awaiter(void 0, void 0, void 0, function* () {
    const { workbook, workSheet } = yield readFile(path);
    const { master, details, simpleVariables, formulas, staticVariables } = yield parse(workSheet);
    const masterTyped = master;
    const detailsTyped = details;
    putSimpleVariables(workSheet, dataToFill, simpleVariables);
    for (const name in simpleVariables) {
        simpleVariables[name].forEach((variable) => {
            if (!staticVariables[variable.address]) {
                staticVariables[variable.address] = {
                    value: variable.insertedValue, address: variable.address, alignment: variable.alignment
                };
            }
        });
    }
    // put master-details
    const lenght = yield putMasterDetail(workSheet, masterTyped, detailsTyped, dataToFill, formulas, staticVariables);
    if (lenght) {
        yield workbook.xlsx.writeFile(reportPath);
        yield (0, readCharts_1.extractZip)(reportPath, `${temporaryFolderPath}/forPictures`);
        // await fixPicturesSizes(reportPath, `${temporaryFolderPath}/forPictures`);
        yield copyDiagramm(path, reportPath, lenght, temporaryFolderPath);
    }
});
const fixPicturesSizes = (reportPath, temporaryFolderPath) => __awaiter(void 0, void 0, void 0, function* () {
    try {
        const filePath = reportPath.replace(/\\/, 'g');
        const fileName = filePath
            .slice(filePath.lastIndexOf('/') + 1, filePath.length)
            .replace('.xlsx', '/');
        const sourceFolder = `${temporaryFolderPath}/${fileName}xl/media/`;
        let imagesSizes = yield fs.readdir(sourceFolder);
        imagesSizes = imagesSizes.map((name) => {
            return `${sourceFolder}${name}`;
        });
        const result = yield readFile(reportPath);
        const images = result.workSheet.getImages();
        imagesSizes = imagesSizes.map((path, index) => {
            const size = sizeOf(path);
            const nativeRow = images[index].range.tl.nativeRow;
            const nativeCol = images[index].range.tl.nativeCol;
            return Object.assign(size, { path, nativeCol, nativeRow });
        });
        images.forEach((image) => {
            image.range.tl.nativeColOff = 0;
            image.range.tl.nativeRowOff = 0;
            image.range.br.nativeRow = 0;
            image.range.br.nativeCol = 0;
        });
        imagesSizes.forEach((image) => {
            const imageId = result.workbook.addImage({
                filename: image.path,
                extension: image.type,
            });
            const { nativeCol, nativeRow, width, height } = image;
            result.workSheet.addImage(imageId, {
                tl: { nativeCol, nativeRow },
                ext: { width, height },
            });
        });
        yield result.workbook.xlsx.writeFile(reportPath);
    }
    catch (e) {
        console.log(e, 'fixPicturesSizes error');
    }
});
function convertObjectToLowercase(obj) {
    const convertedObject = {};
    for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
            const value = obj[key];
            convertedObject[key.toLowerCase()] = value;
        }
    }
    return convertedObject;
}
const writeDataToExcel = (dataToFill, templatePath, reportPath, temporaryFolderPath) => __awaiter(void 0, void 0, void 0, function* () {
    try {
        yield buildTemplate(convertObjectToLowercase(dataToFill), path_1.default.join(templatePath), reportPath, temporaryFolderPath);
    }
    catch (e) {
        console.log(e, 'write data to excel error');
    }
    return true;
});
exports.writeDataToExcel = writeDataToExcel;
