import Workbook from 'exceljs/index';
import {IDetail, IFormulas} from './types';

const parseSimpleVariable = (str: string) => {
  const startDelimiter = '{{';

  const startIndex = str.indexOf(startDelimiter);

  if (startIndex === 0) {
    const variable = str.substring(
      startIndex + startDelimiter.length,
      str.length - startDelimiter.length,
    );
    return variable;
  } else {
    return null;
  }
};

const getSimpleVariable = (cell: Workbook.Worksheet.cell) => {
  let value = null;
  if (typeof cell.value === 'string') {
    value = cell.value;
  }
  if (cell.value && typeof cell.value === 'object') {
    if ('text' in cell.value) {
      value = cell.value.text;
    }
  }
  if (!value) return null;
  const variable = parseSimpleVariable(value);
  if (variable) {
    return { variable, address: cell.address };
  }
};

const parseComplexVariable = (str: string) => {
  const startDelimiter = '[[';
  const endDelimiter = ']]';

  const startIndex = str.indexOf(startDelimiter);
  const endIndex = str.indexOf(endDelimiter);

  if (startIndex === 0) {
    const entityName = str.substring(startIndex + startDelimiter.length, endIndex);
    const fieldName = str.substring(
      endIndex + endDelimiter.length + startDelimiter.length,
      str.length - endDelimiter.length,
    );
    return { entityName, fieldName };
  } else {
    return null;
  }
};

const getComplexVariable = (cell: Workbook.Worksheet.cell): IDetail | null => {
  let value = null;
  if (typeof cell.value === 'string') {
    value = cell.value;
  }
  if (cell.value && typeof cell.value === 'object') {
    if ('text' in cell.value) {
      value = cell.value.text;
    }
  }
  if (!value) return null;
  const variables: any | null = parseComplexVariable(value);
  if (variables) {
    variables.address = cell.address;
    if (cell.address[0] === 'A') {
      variables.type = 'master';
    } else {
      variables.type = 'detail';
    }
    variables.alignment = cell.alignment;
    return variables;
  }
  return null;
};

const getFormula = (cell: Workbook.Worksheet.cell) => {
  let value = null;
  if (cell.value && typeof cell.value === 'object') {
    if ('formula' in cell.value) {
      value = cell.value.formula;
    }
  }
  if (!value) return null;
  return { formula: value, address: cell.address, alignment: cell.alignment };
};

const parseIntFromString = (string: string) => {
  const numberPart = string.match(/\d+/);

  if (numberPart) {
    const number = parseInt(numberPart[0], 10);
    return number;
  } else {
    return 0;
  }
};

const parseAllIntsFromString = (string: string) => {
  const rowMatches = string.match(/[A-Z](\d+)/g);
  if (!rowMatches) return [];
  const rowNumbers = rowMatches.map((match: any) => parseInt(match.match(/\d+/)[0]));

  return rowNumbers;
};

const parseLetterFromString = (string: string) => {
  const columnLetters = string.match(/[A-Z]+/);
  const letter = columnLetters ? columnLetters[0] : null;
  return letter;
};

const isItRowFormula = (string: string) => {
  const numbers = parseAllIntsFromString(string);
  return numbers[0] == numbers[1];
};
function replaceSpecificNumberInFormula(formula: string, targetNumber: number, newNumber: number) {
  const result = formula.replace(new RegExp(String(targetNumber), 'g'), String(newNumber));
  return result;
}

function addDifferenceToTheLastNumber(formula: string, difference: number) {
  let theLastNumber: number | string = formula.split(':')[1].split(')')[0];
  theLastNumber = parseIntFromString(theLastNumber);
  const theNewNumber: number = theLastNumber + difference;
  return replaceSpecificNumberInFormula(formula, theLastNumber, theNewNumber);
}

function extendRange(originalRange: string, extendNumber: number) {
  const [sheetName, range] = originalRange.split('!');
  // @ts-ignore
  const [startRow, endRow] = range.match(/\d+/g);
  const newEndRow: number = +endRow + extendNumber;
  const newRange = originalRange.replace(endRow, newEndRow.toString());

  return newRange;
}

const createArrayIfNotExist = (list: any, entity: string) => {
  if (!list[entity]) {
    list[entity] = [];
  }
};

async function parse(worksheet: Workbook.Worksheet) {
  try {
    const simpleVariables: any = {};
    let master: any = null;
    const details: any = {};
    const formulas: IFormulas = { rowFormulas: [], columnFormulas: [], masterFormulas: [] };
    const staticVariables: any = {};
    let masterRowNumber: number = -1;
    worksheet?.eachRow((row: Workbook.Worksheet.row, rowNumber: number) => {
      row.eachCell((cell: Workbook.Worksheet.cell, colNumber: Workbook.Worksheet.colNumber) => {
        const simpleVariable = getSimpleVariable(cell);
        if (simpleVariable) {
          createArrayIfNotExist(simpleVariables, simpleVariable.variable);
          simpleVariables[simpleVariable.variable].push({
            address: simpleVariable.address,
            alignment: cell.alignment,
            variable: simpleVariable.variable,
          });
        }
        const complexVariable: IDetail | null = getComplexVariable(cell);
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
        const formula = getFormula(cell);

        if (formula) {
          if (isItRowFormula(formula.formula)) {
            if (masterRowNumber === rowNumber) {
              formulas.masterFormulas.push(formula);
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
            alignment: cell.alignment,
          };
        }
      });
    });

    if (Object.keys(details).length === 0) {
    }
    return { simpleVariables, master, details, formulas, staticVariables };
  } catch (err) {
    console.error('Error parse the file:', err);
    return {
      simpleVariables: {},
      master: {},
      details: {},
      formulas: { rowFormulas: [], columnFormulas: [], masterFormulas: [] },
      staticVariables: {},
    };
  }
}

export {
  getSimpleVariable,
  getComplexVariable,
  getFormula,
  parseIntFromString,
  parseLetterFromString,
  isItRowFormula,
  replaceSpecificNumberInFormula,
  addDifferenceToTheLastNumber,
  extendRange,
  parse,
};
