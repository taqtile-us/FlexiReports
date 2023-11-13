import path from 'path';
import {writeFile} from 'fs/promises';
import Workbook from "exceljs/index";
import excel from 'exceljs';
import {
    getSimpleVariable,
    getComplexVariable,
    getFormula,
    parseIntFromString,
    parseLetterFromString,
    isItRowFormula, replaceSpecificNumberInFormula, addDifferenceToTheLastNumber
} from "./Parser";
import {IDetail, IMaster, IDetails, IFormulas, IStaticVariables} from "./types";

const cellFont = { name: 'Arial', size: 11 };

const fs = require('fs').promises;

async function readFile(path: string) {
    try {
        const workbook: Workbook = new excel.Workbook();
        await workbook.xlsx.readFile(path);
        return {workSheet: workbook.getWorksheet('template'), workbook};
    } catch (err) {
        console.error('Error reading the file:', err);
        return null;
    }
}

const createArrayIfNotExist = (list: any, entity: string) => {
    if (!list[entity]) {
        list[entity] = [];
    }
}

async function parse(worksheet: Workbook.Worksheet) {
    try {
        const simpleVariables: any = {}
        let master: any = null;
        const details: any = {};
        const formulas: any = {rowFormulas: [], columnFormulas: []};
        const staticVariables: any = {};
        worksheet?.eachRow((row: Workbook.Worksheet.row, rowNumber: Workbook.Worksheet.rowNumber) => {
            row.eachCell((cell: Workbook.Worksheet.cell, colNumber: Workbook.Worksheet.colNumber) => {
                const simpleVariable = getSimpleVariable(cell);
                if (simpleVariable) {
                    createArrayIfNotExist(simpleVariables, simpleVariable.variable)
                    simpleVariables[simpleVariable.variable].push({address: simpleVariable.address, alignment: cell.alignment})
                }
                const complexVariable: IDetail | null = getComplexVariable(cell);
                if (complexVariable) {
                    if (complexVariable.type === 'master') {
                        master = complexVariable
                    }

                    if (complexVariable.type === 'detail') {
                        createArrayIfNotExist(details, complexVariable.entityName);
                        details[complexVariable.entityName].push(complexVariable)
                    }
                }
                const formula = getFormula(cell);

                if (formula) {
                    if (isItRowFormula(formula.formula)) {
                        formulas.rowFormulas.push(formula);
                    } else {
                        formulas.columnFormulas.push(formula);
                    }
                }

                if (!simpleVariable && !complexVariable && !formula) {
                    staticVariables[cell.address] =  {value: cell.value, address: cell.address, alignment: cell.alignment}
                }
            });
        });
        return {simpleVariables, master, details, formulas, staticVariables}
    } catch (err) {
        console.error('Error parse the file:', err);
        return {simpleVariables: {}, master: {}, details: {}, formulas: [], staticVariables: {}}
    }
}

async function putMasterDetail(worksheet: Workbook.Worksheet, master: IMaster, details: IDetails, data: any, formulas: IFormulas, staticVariables: IStaticVariables) {
    try {
        console.log(formulas, 'formulas')
        const masterDetailDeep = 1;
        let currentRowNumber: number = parseIntFromString(master.address);
        const startRowNumber: number = currentRowNumber;
        const column: string | null = parseLetterFromString(master.address);
        const detailsKeys = Object.keys(details);
        if (!column || !currentRowNumber) return null;
        data[master.entityName].forEach((masterEntity: any) => {
            // put master row
            const cell = worksheet.getCell(column + currentRowNumber);
            cell.value = masterEntity[master.fieldName];
            cell.font = cellFont;
            cell.alignment = master.alignment;
            currentRowNumber += 1;
            worksheet.spliceRows(currentRowNumber + masterDetailDeep, 0, []);
            // put details row
            masterEntity[detailsKeys[0]].forEach((detailEntity: any) => {
                details[detailsKeys[0]].forEach((detailCoords) => {
                    if (detailEntity[detailCoords.fieldName]) {
                        const detailColumn: string | null = parseLetterFromString(detailCoords.address);
                        if (!detailColumn) return
                        const detailCell = worksheet.getCell(detailColumn + currentRowNumber);
                        detailCell.value = detailEntity[detailCoords.fieldName];
                        detailCell.font = cellFont;
                        detailCell.alignment = detailCoords.alignment;
                    }
                })

                // put row formulas
                formulas.rowFormulas.forEach((formula) => {
                    const originalRowNumber = parseIntFromString(formula.formula);
                    const newRowNumber = originalRowNumber + currentRowNumber - startRowNumber - masterDetailDeep;
                    const movedFormula = replaceSpecificNumberInFormula(formula.formula, originalRowNumber, newRowNumber);
                    const movedAddress = replaceSpecificNumberInFormula(formula.address, originalRowNumber, newRowNumber);
                    const formulaCell = worksheet.getCell(movedAddress);
                    formulaCell.value = {formula: movedFormula};
                    formulaCell.alignment = formula.alignment;
                })

                // put static variables
                for (let variableAddress in staticVariables) {
                    if (parseIntFromString(variableAddress) === startRowNumber + 1) {
                        const staticVariableColumn: string | null = parseLetterFromString(variableAddress);
                        const staticVariableCell = worksheet.getCell(`${staticVariableColumn}${currentRowNumber}`);
                        staticVariableCell.value = staticVariables[variableAddress].value;
                        staticVariableCell.alignment = staticVariables[variableAddress].alignment;
                    }
                }

                currentRowNumber += 1
                worksheet.spliceRows(currentRowNumber + masterDetailDeep, 0, []);
            })
        })

        // put column formulas
        const difference = currentRowNumber - startRowNumber;
        formulas.columnFormulas.forEach((formula) => {
                    console.log(formula, 'formula')
                    const originalRowNumber = parseIntFromString(formula.address);
                    const formulaCell = worksheet.getCell(replaceSpecificNumberInFormula(formula.address, originalRowNumber, originalRowNumber + difference));
                    formulaCell.value = {formula: addDifferenceToTheLastNumber(formula.formula, difference)};
                    formulaCell.alignment = formula.alignment;
                })
    } catch (err) {
        console.error('Error put master detail to the file:', err);
        return null;
    }
}

const buildTemplate = async (dataToFill: {}, path: string) => {
    const {workbook, workSheet}: any = await readFile(path);
    const {master, details, simpleVariables, formulas, staticVariables} = await parse(workSheet);
    const masterTyped: IMaster = master;
    const detailsTyped: IDetails = details;
    // put simple variables

    // put master-details
    await putMasterDetail(workSheet, masterTyped, detailsTyped, dataToFill, formulas, staticVariables)

    await workbook.xlsx.writeFile('report.xlsx');


}
export const writeDataToExcel = async (dataToFill: any, templatePath: string) => {
    const filePath = './report.xlsx';
    const buffer: any = await buildTemplate(dataToFill, path.join(templatePath));
    // await writeFile(filePath, buffer, 'binary');

    return filePath;
};
