import Workbook from "exceljs/index";
import {IComplexVariable} from "./types";

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
    const variable =  parseSimpleVariable(value);
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

const getComplexVariable = (cell: Workbook.Worksheet.cell): IComplexVariable | null => {
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
        return variables
    }
    return null;
}


export {getSimpleVariable, getComplexVariable}