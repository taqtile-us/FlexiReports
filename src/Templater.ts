import path from 'path';
import Workbook from 'exceljs/index';
import { copyChart } from './copy-excel-chart/build/copyChart';
import { readCharts, extractZip } from './copy-excel-chart/build/readCharts';
import { writeCharts } from './copy-excel-chart/build/writeChart';
import {
  parseIntFromString,
  parseLetterFromString,
  replaceSpecificNumberInFormula,
  addDifferenceToTheLastNumber,
  extendRange, parse,
} from './Parser';
import { IDetail, IMaster, IDetails, IFormulas, IStaticVariables, ISimpleVariables } from './types';
import { ISizeCalculationResult } from 'image-size/dist/types/interface';
import { readFile } from "./Reader";

const cellFont = { name: 'Arial', size: 11 };

const fs = require('fs').promises;
const sizeOf = require('image-size');

const copyDiagramm = async (
  template: string,
  report: string,
  length: number,
  temporaryFolderPath: string,
) => {
  try {
    await fs.mkdir(temporaryFolderPath, { recursive: true });
    console.log(`Folder created (or already exists): ${temporaryFolderPath}`);
  } catch (error: any) {
    console.error(`Error creating folder: ${error.message}`);
  }

  try {
    const source = await readCharts(template, temporaryFolderPath);
    const output = await readCharts(report, temporaryFolderPath);
    source.worksheets.template.drawingRels.chart1 = 'rid3';
    const summary = source.summary();

    if (summary['template']['chart1']) {
      const replaceCellRefs = summary['template']['chart1'].reduce((acc: any, el: any) => {
        return { ...acc, [el]: el.replace('recommendWorksheet2', 'worksheet-Recommendation') };
      }, {});

      for (const key in replaceCellRefs) {
        replaceCellRefs[key] = extendRange(replaceCellRefs[key], length);
      }
      await copyChart(source, output, 'template', 'chart1', 'template', replaceCellRefs);

      await writeCharts(output, report);
    }
  } catch (error: any) {
    console.error(`Error copy diagramm: ${error.message}`);
  }
};

async function putMasterDetail(
  worksheet: Workbook.Worksheet,
  master: IMaster,
  details: IDetails,
  data: any,
  formulas: IFormulas,
  staticVariables: IStaticVariables,
) {
  try {
    const masterDetailDeep = 1;
    const detailsDeep = 0;
    let currentRowNumber: number = parseIntFromString(master.address);
    const startRowNumber: number = currentRowNumber;
    const column: string | null = parseLetterFromString(master.address);
    const detailsKeys = Object.keys(details);
    if (!column || !currentRowNumber) return null;

    data[master.entityName.toLowerCase()].forEach((masterEntity: any) => {
      if (!master.addedToDetails) {
        if (masterEntity[detailsKeys[0]].length === 0) {
          return;
        }
        putMasterRow(worksheet, masterEntity, master, column, currentRowNumber);
        putMasterFormulas(worksheet, formulas, masterEntity, detailsKeys, currentRowNumber);

        currentRowNumber += 1;
        worksheet.spliceRows(currentRowNumber + masterDetailDeep, 0, []);
        masterEntity[detailsKeys[0]].forEach((detailEntity: any) => {
          putDetailRow(worksheet, details, detailsKeys, detailEntity, currentRowNumber);

          putDetailFormula(worksheet, formulas, currentRowNumber, startRowNumber, masterDetailDeep);

          putStaticVariables(worksheet, staticVariables, currentRowNumber, startRowNumber);

          currentRowNumber += 1;
          worksheet.spliceRows(currentRowNumber + masterDetailDeep, 0, []);
        });
      } else {
        putDetailRow(worksheet, details, detailsKeys, masterEntity, currentRowNumber);

        putDetailFormula(
          worksheet,
          { rowFormulas: formulas.masterFormulas },
          currentRowNumber,
          startRowNumber,
          detailsDeep,
        );

        putStaticVariables(worksheet, staticVariables, currentRowNumber, startRowNumber - 1);

        currentRowNumber += 1;
        worksheet.spliceRows(currentRowNumber + masterDetailDeep, 0, []);
      }
    });

    const difference = currentRowNumber - startRowNumber;
    putColumnFormulas(worksheet, formulas, difference);
    return currentRowNumber - startRowNumber;
  } catch (err) {
    console.error('Error put master detail to the file:', err);
    return null;
  }
}

function putMasterRow(worksheet, masterEntity, master, column, currentRowNumber) {
  const cell = worksheet.getCell(column + currentRowNumber);
  cell.value = masterEntity[master.fieldName];
  cell.font = cellFont;
  cell.alignment = master.alignment;
}

function putMasterFormulas(worksheet, formulas, masterEntity, detailsKeys, currentRowNumber) {
  if (formulas.masterFormulas.length) {
    const masterFormulaColumn: string | null = parseLetterFromString(
      formulas.masterFormulas[0].address,
    );
    if (masterFormulaColumn) {
      formulas.masterFormulas.forEach((masterFormula) => {
        const cell = worksheet.getCell(masterFormulaColumn + currentRowNumber);
        cell.value = {
          formula: `SUM(G${currentRowNumber + 1}:G${
            currentRowNumber + masterEntity[detailsKeys[0]].length
          })`,
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
      const detailColumn: string | null = parseLetterFromString(detailCoords.address);
      if (!detailColumn) return;
      const detailCell = worksheet.getCell(detailColumn + currentRowNumber);
      detailCell.value = detailEntity[detailCoords.fieldName];
      detailCell.font = cellFont;
      detailCell.alignment = detailCoords.alignment;
    }
  });
}

function putDetailFormula(worksheet, formulas, currentRowNumber, startRowNumber, masterDetailDeep) {
  formulas.rowFormulas.forEach((formula) => {
    const originalRowNumber = parseIntFromString(formula.formula);
    const newRowNumber = originalRowNumber + currentRowNumber - startRowNumber - masterDetailDeep;
    const movedFormula = replaceSpecificNumberInFormula(
      formula.formula,
      originalRowNumber,
      newRowNumber,
    );
    const movedAddress = replaceSpecificNumberInFormula(
      formula.address,
      originalRowNumber,
      newRowNumber,
    );
    const formulaCell = worksheet.getCell(movedAddress);
    formulaCell.value = { formula: movedFormula };
    formulaCell.font = cellFont;
    formulaCell.alignment = formula.alignment;
  });
}

function putColumnFormulas(worksheet, formulas, difference) {
  formulas.columnFormulas.forEach((formula) => {
    const originalRowNumber = parseIntFromString(formula.address);
    const formulaCell = worksheet.getCell(
      replaceSpecificNumberInFormula(
        formula.address,
        originalRowNumber,
        originalRowNumber + difference,
      ),
    );
    formulaCell.value = { formula: addDifferenceToTheLastNumber(formula.formula, difference) };
    formulaCell.alignment = formula.alignment;
  });
}

function putStaticVariables(worksheet, staticVariables, currentRowNumber, startRowNumber) {
  for (const variableAddress in staticVariables) {
    if (parseIntFromString(variableAddress) === startRowNumber + 1) {
      const staticVariableColumn: string | null = parseLetterFromString(variableAddress);
      const staticVariableCell = worksheet.getCell(`${staticVariableColumn}${currentRowNumber}`);
      staticVariableCell.value = staticVariables[variableAddress].value;
      staticVariableCell.font = cellFont;
      staticVariableCell.alignment = staticVariables[variableAddress].alignment;
    }
  }
}

async function putSimpleVariables(
  worksheet: Workbook.Worksheet,
  data: any,
  simpleVariables: ISimpleVariables,
) {
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
  } catch (err) {
    console.error('Error put simple variables to the file:', err);
    return null;
  }
}

const buildTemplate = async (
  dataToFill: {},
  path: string,
  reportPath: string,
  temporaryFolderPath: string,
) => {
  const { workbook, workSheet }: any = await readFile(path);
  const { master, details, simpleVariables, formulas, staticVariables } = await parse(workSheet);

  const masterTyped: IMaster = master;
  const detailsTyped: IDetails = details;
  putSimpleVariables(workSheet, dataToFill, simpleVariables);
  for (const name in simpleVariables) {
    simpleVariables[name].forEach((variable) => {
      if (!staticVariables[variable.address]) {
        staticVariables[variable.address] = {
          value: variable.insertedValue, address: variable.address, alignment: variable.alignment
        }
      }
    })
  }
  // put master-details
  const lenght: number | null = await putMasterDetail(
    workSheet,
    masterTyped,
    detailsTyped,
    dataToFill,
    formulas,
    staticVariables,
  );
  if (lenght) {
    await workbook.xlsx.writeFile(reportPath);
    await extractZip(reportPath, `${temporaryFolderPath}/forPictures`);
    // await fixPicturesSizes(reportPath, `${temporaryFolderPath}/forPictures`);
    await copyDiagramm(path, reportPath, lenght, temporaryFolderPath);
  }
};

const fixPicturesSizes = async (reportPath, temporaryFolderPath) => {
  try {
    const filePath = reportPath.replace(/\\/, 'g');
    const fileName = filePath
      .slice(filePath.lastIndexOf('/') + 1, filePath.length)
      .replace('.xlsx', '/');
    const sourceFolder = `${temporaryFolderPath}/${fileName}xl/media/`;
    let imagesSizes = await fs.readdir(sourceFolder);
    imagesSizes = imagesSizes.map((name) => {
      return `${sourceFolder}${name}`;
    });
    const result: any = await readFile(reportPath);
    const images = result.workSheet.getImages();
    imagesSizes = imagesSizes.map((path, index) => {
      const size: ISizeCalculationResult = sizeOf(path);
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
    await result.workbook.xlsx.writeFile(reportPath);
  } catch (e) {
    console.log(e, 'fixPicturesSizes error');
  }
};

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

export const writeDataToExcel = async (
  dataToFill: any,
  templatePath: string,
  reportPath: string,
  temporaryFolderPath: string,
) => {
  try {
    await buildTemplate(
      convertObjectToLowercase(dataToFill),
      path.join(templatePath),
      reportPath,
      temporaryFolderPath,
    );
  } catch (e) {
    console.log(e, 'write data to excel error');
  }

  return true;
};
