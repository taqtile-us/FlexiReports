import ExcelJS from 'exceljs';
import { Workbook } from 'exceljs/index.d';
import { writeDataToExcel } from './Templater';

export const generateExcelReport = async (dataToFill: any, templatePath: string, reportPath: string) => {
  try {
    const workbook: Workbook = new ExcelJS.Workbook();

    await workbook.xlsx.readFile(templatePath);
    const pathToReport = await writeDataToExcel(dataToFill, templatePath, reportPath);

    return pathToReport;
  } catch (e) {
    console.log(e, 'generateExcelReport error');
    return false;
  }
};
