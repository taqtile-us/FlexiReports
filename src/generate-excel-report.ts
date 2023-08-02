import ExcelJS from 'exceljs';
import { Workbook } from 'exceljs/index.d';
import { writeDataToExcel } from './Templater';

export const generateExcelReport = async (dataToFill: any, filePath: string) => {
  try {
    const workbook: Workbook = new ExcelJS.Workbook();

    await workbook.xlsx.readFile(filePath);
    const pathToReport = await writeDataToExcel(dataToFill, filePath);

    return pathToReport;
  } catch (e) {
    console.log(e, 'generateExcelReport error');
    return false;
  }
};
