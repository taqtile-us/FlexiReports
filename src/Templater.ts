import path from 'path';
import {writeFile} from 'fs/promises';
import Workbook from "exceljs/index";
// @ts-ignore
import excel from 'exceljs';
// @ts-ignore
import copyExcelChart from './copy-excel-chart/copy-excel-chart'

const readCharts = copyExcelChart.readCharts
const copyChart = copyExcelChart.copyChart
const writeCharts = copyExcelChart.writeCharts
import {
    getSimpleVariable,
    getComplexVariable,
    getFormula,
    parseIntFromString,
    parseLetterFromString,
    isItRowFormula, replaceSpecificNumberInFormula, addDifferenceToTheLastNumber, extendRange
} from "./Parser";
import {IDetail, IMaster, IDetails, IFormulas, IStaticVariables, ISimpleVariables} from "./types";

const cellFont = {name: 'Arial', size: 11};

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

const copyDiagramm = async (template: string, report: string, length: number) => {
    const source = await readCharts(template, './temp')
    const output = await readCharts(report, './temp')
    const summary = source.summary();
    let replaceCellRefs = summary['template']['chart1'].reduce((acc: any, el: any) => {
        return {...acc, [el]: el.replace('recommendWorksheet2', 'worksheet-Recommendation')}
    }, {})


    for (let key in replaceCellRefs) {
        replaceCellRefs[key] = extendRange(replaceCellRefs[key], length)
    }
    copyChart(
        source,
        output,
        'template',
        'chart1',
        'template',
        replaceCellRefs,
    )

    writeCharts(output, report)
}

async function parse(worksheet: Workbook.Worksheet) {
    try {
        const simpleVariables: any = {}
        let master: any = null;
        const details: any = {};
        const formulas: IFormulas = {rowFormulas: [], columnFormulas: [], masterFormulas: []};
        const staticVariables: any = {};
        let masterRowNumber: number = -1;
        worksheet?.eachRow((row: Workbook.Worksheet.row, rowNumber: number) => {
            row.eachCell((cell: Workbook.Worksheet.cell, colNumber: Workbook.Worksheet.colNumber) => {
                const simpleVariable = getSimpleVariable(cell);
                if (simpleVariable) {
                    createArrayIfNotExist(simpleVariables, simpleVariable.variable)
                    simpleVariables[simpleVariable.variable].push({
                        address: simpleVariable.address,
                        alignment: cell.alignment,
                        variable: simpleVariable.variable
                    })
                }
                const complexVariable: IDetail | null = getComplexVariable(cell);
                if (complexVariable) {
                    if (complexVariable.type === 'master') {
                        master = complexVariable
                        masterRowNumber = rowNumber
                    }

                    if (complexVariable.type === 'detail') {
                        createArrayIfNotExist(details, complexVariable.entityName);
                        details[complexVariable.entityName].push(complexVariable)
                    }
                }
                const formula = getFormula(cell);

                if (formula) {
                    if (isItRowFormula(formula.formula)) {
                        if (masterRowNumber === rowNumber) {
                            formulas.masterFormulas.push(formula)
                        } else {
                            formulas.rowFormulas.push(formula);
                        }
                    } else {
                        formulas.columnFormulas.push(formula);
                    }
                }

                if (!simpleVariable && !complexVariable && !formula) {
                    staticVariables[cell.address] = {
                        value: cell.value,
                        address: cell.address,
                        alignment: cell.alignment
                    }
                }
            });
        });
        return {simpleVariables, master, details, formulas, staticVariables}
    } catch (err) {
        console.error('Error parse the file:', err);
        return {
            simpleVariables: {},
            master: {},
            details: {},
            formulas: {rowFormulas: [], columnFormulas: [], masterFormulas: []},
            staticVariables: {}
        }
    }
}

async function putMasterDetail(worksheet: Workbook.Worksheet, master: IMaster, details: IDetails, data: any, formulas: IFormulas, staticVariables: IStaticVariables) {
    try {
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

            // put master formula
            if (formulas.masterFormulas.length) {
                const masterFormulaColumn: string | null = parseLetterFromString(formulas.masterFormulas[0].address);
                if (masterFormulaColumn) {
                    formulas.masterFormulas.forEach((masterFormula) => {
                        const cell = worksheet.getCell(masterFormulaColumn + currentRowNumber);
                        cell.value = {formula: `SUM(G${currentRowNumber + 1}:G${currentRowNumber + masterEntity[detailsKeys[0]].length})`};
                        cell.font = cellFont;
                        cell.alignment = masterFormula.alignment;
                    })
                }

            }


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
            const originalRowNumber = parseIntFromString(formula.address);
            const formulaCell = worksheet.getCell(replaceSpecificNumberInFormula(formula.address, originalRowNumber, originalRowNumber + difference));
            formulaCell.value = {formula: addDifferenceToTheLastNumber(formula.formula, difference)};
            formulaCell.alignment = formula.alignment;
        })
        return currentRowNumber - startRowNumber;
    } catch (err) {
        console.error('Error put master detail to the file:', err);
        return null;
    }
}

async function putSimpleVariables(worksheet: Workbook.Worksheet, data: any, simpleVariables: ISimpleVariables) {
    try {
        for (let variable in simpleVariables) {
            if (data[variable]) {
                simpleVariables[variable].forEach((simpleVariable) => {
                    const variableCell = worksheet.getCell(simpleVariable.address);
                    variableCell.value = data[variable];
                    variableCell.alignment = simpleVariable.alignment;
                })
            }
        }
    } catch (err) {
        console.error('Error put simple variables to the file:', err);
        return null;
    }
}

const buildTemplate = async (dataToFill: {}, path: string) => {
    const {workbook, workSheet}: any = await readFile(path);
    const {master, details, simpleVariables, formulas, staticVariables} = await parse(workSheet);

    const masterTyped: IMaster = master;
    const detailsTyped: IDetails = details;
    // put simple variables
    putSimpleVariables(workSheet, dataToFill, simpleVariables)

    // put master-details
    const lenght: number | null = await putMasterDetail(workSheet, masterTyped, detailsTyped, dataToFill, formulas, staticVariables);
    if (lenght) {
        await workbook.xlsx.writeFile('report.xlsx');
        copyDiagramm(path, './report.xlsx', lenght)
    }


}
export const writeDataToExcel = async (dataToFill: any, templatePath: string) => {
    const filePath = './report.xlsx';
    const buffer: any = await buildTemplate(dataToFill, path.join(templatePath));
    // await writeFile(filePath, buffer, 'binary');

    return filePath;
};
