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
Object.defineProperty(exports, "__esModule", { value: true });
exports.parse = exports.extendRange = exports.addDifferenceToTheLastNumber = exports.replaceSpecificNumberInFormula = exports.isItRowFormula = exports.parseLetterFromString = exports.parseIntFromString = exports.getFormula = exports.getComplexVariable = exports.getSimpleVariable = void 0;
const parseSimpleVariable = (str) => {
    const startDelimiter = '{{';
    const startIndex = str.indexOf(startDelimiter);
    if (startIndex === 0) {
        const variable = str.substring(startIndex + startDelimiter.length, str.length - startDelimiter.length);
        return variable;
    }
    else {
        return null;
    }
};
const getSimpleVariable = (cell) => {
    let value = null;
    if (typeof cell.value === 'string') {
        value = cell.value;
    }
    if (cell.value && typeof cell.value === 'object') {
        if ('text' in cell.value) {
            value = cell.value.text;
        }
    }
    if (!value)
        return null;
    const variable = parseSimpleVariable(value);
    if (variable) {
        return { variable, address: cell.address };
    }
};
exports.getSimpleVariable = getSimpleVariable;
const parseComplexVariable = (str) => {
    const startDelimiter = '[[';
    const endDelimiter = ']]';
    const startIndex = str.indexOf(startDelimiter);
    const endIndex = str.indexOf(endDelimiter);
    if (startIndex === 0) {
        const entityName = str.substring(startIndex + startDelimiter.length, endIndex);
        const fieldName = str.substring(endIndex + endDelimiter.length + startDelimiter.length, str.length - endDelimiter.length);
        return { entityName, fieldName };
    }
    else {
        return null;
    }
};
const getComplexVariable = (cell) => {
    let value = null;
    if (typeof cell.value === 'string') {
        value = cell.value;
    }
    if (cell.value && typeof cell.value === 'object') {
        if ('text' in cell.value) {
            value = cell.value.text;
        }
    }
    if (!value)
        return null;
    const variables = parseComplexVariable(value);
    if (variables) {
        variables.address = cell.address;
        if (cell.address[0] === 'A') {
            variables.type = 'master';
        }
        else {
            variables.type = 'detail';
        }
        variables.alignment = cell.alignment;
        return variables;
    }
    return null;
};
exports.getComplexVariable = getComplexVariable;
const getFormula = (cell) => {
    let value = null;
    if (cell.value && typeof cell.value === 'object') {
        if ('formula' in cell.value) {
            value = cell.value.formula;
        }
    }
    if (!value)
        return null;
    return { formula: value, address: cell.address, alignment: cell.alignment };
};
exports.getFormula = getFormula;
const parseIntFromString = (string) => {
    const numberPart = string.match(/\d+/);
    if (numberPart) {
        const number = parseInt(numberPart[0], 10);
        return number;
    }
    else {
        return 0;
    }
};
exports.parseIntFromString = parseIntFromString;
const parseAllIntsFromString = (string) => {
    const rowMatches = string.match(/[A-Z](\d+)/g);
    if (!rowMatches)
        return [];
    const rowNumbers = rowMatches.map((match) => parseInt(match.match(/\d+/)[0]));
    return rowNumbers;
};
const parseLetterFromString = (string) => {
    const columnLetters = string.match(/[A-Z]+/);
    const letter = columnLetters ? columnLetters[0] : null;
    return letter;
};
exports.parseLetterFromString = parseLetterFromString;
const isItRowFormula = (string) => {
    const numbers = parseAllIntsFromString(string);
    return numbers[0] == numbers[1];
};
exports.isItRowFormula = isItRowFormula;
function replaceSpecificNumberInFormula(formula, targetNumber, newNumber) {
    const result = formula.replace(new RegExp(String(targetNumber), 'g'), String(newNumber));
    return result;
}
exports.replaceSpecificNumberInFormula = replaceSpecificNumberInFormula;
function addDifferenceToTheLastNumber(formula, difference) {
    let theLastNumber = formula.split(':')[1].split(')')[0];
    theLastNumber = parseIntFromString(theLastNumber);
    const theNewNumber = theLastNumber + difference;
    return replaceSpecificNumberInFormula(formula, theLastNumber, theNewNumber);
}
exports.addDifferenceToTheLastNumber = addDifferenceToTheLastNumber;
function extendRange(originalRange, extendNumber) {
    const [sheetName, range] = originalRange.split('!');
    // @ts-ignore
    const [startRow, endRow] = range.match(/\d+/g);
    const newEndRow = +endRow + extendNumber;
    const newRange = originalRange.replace(endRow, newEndRow.toString());
    return newRange;
}
exports.extendRange = extendRange;
const createArrayIfNotExist = (list, entity) => {
    if (!list[entity]) {
        list[entity] = [];
    }
};
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
                    const simpleVariable = getSimpleVariable(cell);
                    if (simpleVariable) {
                        createArrayIfNotExist(simpleVariables, simpleVariable.variable);
                        simpleVariables[simpleVariable.variable].push({
                            address: simpleVariable.address,
                            alignment: cell.alignment,
                            variable: simpleVariable.variable,
                        });
                    }
                    const complexVariable = getComplexVariable(cell);
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
exports.parse = parse;
