import ExcelJS from 'exceljs';

type ExcelWorkbook = ExcelJS.Workbook;
type ExcelComment = {
  rowNumber: number;
  colNumber: number;
  cell: any;
};
type ExcelFormula = {
  value: string;
  address: string;
  numberAddress: number;
};

type ExcelComments = {
  [key: string]: ExcelComment;
};

type ExcelFormulas = {
  [key: string]: ExcelFormula;
};

const getComments = async (workbook: ExcelWorkbook): Promise<ExcelComments> => {
  const worksheet = workbook.getWorksheet(1);
  const coords: ExcelComments = {};
  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell: any, colNumber) => {
      const comment = cell._comment.note.texts[0].text;
      if (comment) {
        coords[comment] = { colNumber, rowNumber, cell };
      }
    });
  });

  return coords;
};

const getCoordsOfFormulaCell = async (path: string): Promise<ExcelFormulas> => {
  const workbook: ExcelWorkbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path);
  const worksheet = workbook.getWorksheet(1);
  const coords: ExcelFormulas = {};
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell: any) => {
      const formula = cell._value.model.formula;
      if (formula) {
        const address = cell._address;
        const formulaValue = cell._value.model.formula;
        const number = cell._row._number;
        coords[address] = { address, value: formulaValue, numberAddress: number };
      }
    });
  });

  return coords;
};

const generateReportWithFormula = async (path: string, formulas: ExcelFormulas) => {
  try {
    const workbook: ExcelWorkbook = new ExcelJS.Workbook();
    const formulasData = Object.values(formulas);
    if (!formulasData.length) {
      return;
    }
    await workbook.xlsx.readFile(path);
    const worksheet = workbook.getWorksheet(1);
    formulasData.forEach((formulaData) => {
      const formula = formulaData.value.toString();
      if (formula.includes('(') && formula.includes(')')) {
        const partsOfRange = formula.split('(');

        const startRangeIndex = formula.indexOf('(') + '('.length;

        const range = formula.substring(startRangeIndex, formula.indexOf(')', startRangeIndex));

        if (range.includes(':')) {
          const prevPosition = range.split(':');
          const updatedRange = prevPosition[1][0] + (formulaData.numberAddress - 1);
          prevPosition[1] = updatedRange;
          const secondRangePart = '(' + prevPosition.join(':') + ')';

          worksheet.getCell(formulaData.address).value = {
            formula: partsOfRange[0] + secondRangePart,
            date1904: false,
          };
        }
      }
    });

    await workbook.xlsx.writeFile(path);
  } catch (e) {
    console.log(e, 'generate report error');
    return false;
  }
};

export { generateReportWithFormula, getCoordsOfFormulaCell };
