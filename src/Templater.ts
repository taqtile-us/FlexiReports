import path from 'path';
import {writeFile} from 'fs/promises';
import Workbook from "exceljs/index";
import excel from 'exceljs';
import ExcelJS from "exceljs";
import {getSimpleVariable, getComplexVariable} from "./Parser";
import {IComplexVariable} from "./types";

const fs = require('fs').promises;

async function readFile(path: string) {
    try {
        const workbook: Workbook = new excel.Workbook();
        await workbook.xlsx.readFile(path);
        return workbook.getWorksheet('template');
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
        worksheet?.eachRow((row: Workbook.Worksheet.row, rowNumber: Workbook.Worksheet.rowNumber) => {
            row.eachCell((cell: Workbook.Worksheet.cell, colNumber: Workbook.Worksheet.colNumber) => {
                const simpleVariable = getSimpleVariable(cell);
                if (simpleVariable) {
                    createArrayIfNotExist(simpleVariables, simpleVariable.variable)
                    simpleVariables[simpleVariable.variable].push(simpleVariable.address)
                }
                const complexVariable: IComplexVariable | null = getComplexVariable(cell);
                if (complexVariable) {
                    if (complexVariable.type === 'master') {
                        master = complexVariable
                    }

                    if (complexVariable.type === 'detail') {
                        createArrayIfNotExist(details, complexVariable.entityName);
                        details[complexVariable.entityName].push(complexVariable)
                    }
                }
            });
        });
        console.log(simpleVariables, 'simpleVariables')
        console.log(master, 'master')
        console.log(details, 'details')
    } catch (err) {
        console.error('Error parse the file:', err);
        return null;
    }
}

const buildTemplate = async (dataToFill: {}, path: string) => {
    const workbook = await readFile(path);
    const coordinatesForInsertion = await parse(workbook);
}
export const writeDataToExcel = async (dataToFill: any, templatePath: string) => {
    const filePath = 'uploads/report.xlsx';
    const buffer: any = await buildTemplate(dataToFill, path.join(templatePath));

    await writeFile(filePath, buffer, 'binary');

    return filePath;
};
