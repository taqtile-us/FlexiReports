import Workbook from "exceljs/index";
import {IDetail} from "./types";

const parseSimpleVariable = (str: string) => {
    const startDelimiter = '{{';

    const startIndex = str.indexOf(startDelimiter);

    if (startIndex === 0) {
        const variable = str.substring(
            startIndex + startDelimiter.length,
            str.length - startDelimiter.length
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
        return {variable, address: cell.address}
    }
}

const parseComplexVariable = (str: string) => {
    const startDelimiter = '[[';
    const endDelimiter = ']]'

    const startIndex = str.indexOf(startDelimiter);
    const endIndex = str.indexOf(endDelimiter);

    if (startIndex === 0) {
        const entityName = str.substring(
            startIndex + startDelimiter.length,
            endIndex
        );
        const fieldName = str.substring(
            endIndex + endDelimiter.length + startDelimiter.length,
            str.length - endDelimiter.length
        );
        return {entityName, fieldName};
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
            variables.type = 'master'
        } else {
            variables.type = 'detail'
        }
        variables.alignment = cell.alignment;
        return variables
    }
    return null;
}

const getFormula = (cell: Workbook.Worksheet.cell) => {
    let value = null;
    if (cell.value && typeof cell.value === 'object') {
        if ('formula' in cell.value) {
            value = cell.value.formula;
        }
    }
    if (!value) return null;
    return {formula: value, address: cell.address, alignment: cell.alignment}
}

const parseIntFromString = (string: string) => {
    const numberPart = string.match(/\d+/);

    if (numberPart) {
        const number = parseInt(numberPart[0], 10);
        return number;
    } else {
        return 0
    }
}

const parseAllIntsFromString = (string: string) => {

    const rowMatches = string.match(/[A-Z](\d+)/g);
    if (!rowMatches) return [];
    const rowNumbers = rowMatches.map((match: any) => parseInt(match.match(/\d+/)[0]));

    return rowNumbers
}

const parseLetterFromString = (string: string) => {
    const columnLetters = string.match(/[A-Z]+/);
    const letter = columnLetters ? columnLetters[0] : null;
    return letter;
}

const isItRowFormula = (string: string) => {
    const numbers = parseAllIntsFromString(string);
    return numbers[0] == numbers[1]
}

function replaceSpecificNumberInFormula(formula: string, targetNumber: number, newNumber: number) {
    return formula.replace(`/${targetNumber}/g`, String(newNumber));
}

function addDifferenceToTheLastNumber(formula: string, difference: number) {
    let theLastNumber: number | string = formula.split(':')[1].split(')')[0];
    theLastNumber = parseIntFromString(theLastNumber);
    const theNewNumber: number = theLastNumber + difference;
    return replaceSpecificNumberInFormula(formula, theLastNumber, theNewNumber)
}

function extendRange(originalRange: string, extendNumber: number) {
    const [sheetName, range] = originalRange.split('!');
    // @ts-ignore
    const [startRow, endRow] = range.match(/\d+/g);
    const newEndRow: number = +endRow + extendNumber;
    const newRange = originalRange.replace(endRow, newEndRow.toString())

    return newRange;
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
    extendRange
}